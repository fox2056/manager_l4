use std::path::PathBuf;
use anyhow::Result;
use calamine::{Reader, Xlsx, Xls, open_workbook, DataType};
use regex::Regex;
use std::collections::HashSet;
use rust_xlsxwriter::{Workbook, Format, FormatBorder, Color, Worksheet};
use std::error::Error;

#[derive(Debug, Clone)]
pub struct EmployeeData {
    pub nazwisko: String,
    pub imie: String,
    pub pesel: String,
    pub data_od: Option<f64>,
    pub data_do: Option<f64>,
    pub na_opieke: String,
    pub pobyt_w_szpitalu: String,
    pub status: String,
    pub source: String,
}

pub struct ExcelMerger {
    pub messages: Vec<String>,
}

impl ExcelMerger {
    pub fn new() -> Self {
        Self {
            messages: Vec::new(),
        }
    }

    fn log_message(&mut self, message: String) {
        self.messages.push(message);
    }

    pub fn get_sheet_names(&mut self, path: &PathBuf) -> Vec<String> {
        let extension = path.extension()
            .and_then(|ext| ext.to_str())
            .unwrap_or("")
            .to_lowercase();

        match extension.as_str() {
            "xlsx" => {
                match open_workbook::<Xlsx<_>, _>(path) {
                    Ok(wb) => {
                        let sheet_names = wb.sheet_names().to_vec();
                        self.log_message(format!("Wczytano arkusze z pliku XLSX: {}", path.display()));
                        sheet_names
                    },
                    Err(e) => {
                        self.log_message(format!("Błąd podczas wczytywania arkuszy XLSX: {}", e));
                        Vec::new()
                    }
                }
            },
            "xls" => {
                match open_workbook::<Xls<_>, _>(path) {
                    Ok(wb) => {
                        let sheet_names = wb.sheet_names().to_vec();
                        self.log_message(format!("Wczytano arkusze z pliku XLS: {}", path.display()));
                        sheet_names
                    },
                    Err(e) => {
                        self.log_message(format!("Błąd podczas wczytywania arkuszy XLS: {}", e));
                        Vec::new()
                    }
                }
            },
            _ => {
                self.log_message(format!("Nieobsługiwany format pliku: {}", extension));
                Vec::new()
            }
        }
    }

    fn parse_date(date_str: &str) -> Option<f64> {
        if date_str.is_empty() {
            return None;
        }

        let parts: Vec<&str> = date_str.split('-').collect();
        if parts.len() != 3 {
            return None;
        }

        let year = parts[0].parse::<i32>().ok()?;
        let month = parts[1].parse::<u32>().ok()?;
        let day = parts[2].parse::<u32>().ok()?;

        if let Some(date) = chrono::NaiveDate::from_ymd_opt(year, month, day) {
            let excel_date = (date - chrono::NaiveDate::from_ymd_opt(1900, 1, 1).unwrap()).num_days() as f64 + 2.0;
            Some(excel_date)
        } else {
            None
        }
    }

    pub fn merge_files(
        &mut self,
        first_file: &PathBuf,
        second_file: &PathBuf,
        output_file: &PathBuf,
        first_sheet: &str,
        second_sheet: &str,
    ) -> Result<(), Box<dyn Error>> {
        let mut workbook = Workbook::new();
        let mut sheet = workbook.add_worksheet();
        
        let header_format = Format::new()
            .set_bold()
            .set_background_color(Color::RGB(0x4F8_1BD))
            .set_font_color(Color::RGB(0xFFFF_FF))
            .set_border(FormatBorder::Thin);
            
        let date_format = Format::new().set_num_format("dd/mm/yyyy");
        
        let (headers, data) = self.prepare_data(first_file, second_file, first_sheet, second_sheet)?;
        let common_pesels = self.find_common_pesels(&data);
        
        let filtered_data: Vec<EmployeeData> = data.into_iter()
            .filter(|employee| employee.source == "l4" && common_pesels.contains(&employee.pesel))
            .collect();
        
        self.write_headers(&mut sheet, &headers, &header_format)?;
        self.write_data(&mut sheet, &filtered_data, &date_format)?;
        
        workbook.save(output_file)?;
        
        let message = format!("\nLiczba wspólnych numerów PESEL: {}", common_pesels.len());
        self.log_message(message);
        self.log_message(format!("Utworzono plik wynikowy: {}", output_file.display()));
        
        Ok(())
    }

    fn prepare_data(
        &mut self,
        first_file: &PathBuf,
        second_file: &PathBuf,
        first_sheet: &str,
        second_sheet: &str,
    ) -> Result<(Vec<String>, Vec<EmployeeData>)> {
        let range1 = if first_file.extension().and_then(|ext| ext.to_str()).map(|s| s.to_lowercase()) == Some("xlsx".to_string()) {
            let mut workbook: Xlsx<_> = open_workbook(first_file)?;
            workbook.worksheet_range(first_sheet)
                .ok_or_else(|| anyhow::anyhow!("Nie można otworzyć pierwszego arkusza"))??
        } else {
            let mut workbook: Xls<_> = open_workbook(first_file)?;
            workbook.worksheet_range(first_sheet)
                .ok_or_else(|| anyhow::anyhow!("Nie można otworzyć pierwszego arkusza"))??
        };

        let range2 = if second_file.extension().and_then(|ext| ext.to_str()).map(|s| s.to_lowercase()) == Some("xlsx".to_string()) {
            let mut workbook: Xlsx<_> = open_workbook(second_file)?;
            workbook.worksheet_range(second_sheet)
                .ok_or_else(|| anyhow::anyhow!("Nie można otworzyć drugiego arkusza"))??
        } else {
            let mut workbook: Xls<_> = open_workbook(second_file)?;
            workbook.worksheet_range(second_sheet)
                .ok_or_else(|| anyhow::anyhow!("Nie można otworzyć drugiego arkusza"))??
        };

        let headers = vec![
            "Nazwisko".to_string(),
            "Imię".to_string(),
            "PESEL".to_string(),
            "Data od".to_string(),
            "Data do".to_string(),
            "Na opiekę".to_string(),
            "Pobyt w szpitalu".to_string(),
            "Status zaśw.".to_string(),
        ];

        let mut data = Vec::new();
        let mut pracownicy_count = 0;
        let mut l4_count = 0;

        self.log_message("\nWczytywanie danych z pliku pracowników:".to_string());
        for row in range1.rows().skip(1) {
            if let (Some(DataType::String(nazwisko)), Some(DataType::String(imie)), Some(DataType::String(pesel))) = 
                (row.get(0), row.get(1), row.get(2)) {
                
                data.push(EmployeeData {
                    nazwisko: nazwisko.to_string(),
                    imie: imie.to_string(),
                    pesel: pesel.to_string(),
                    data_od: None,
                    data_do: None,
                    na_opieke: String::new(),
                    pobyt_w_szpitalu: String::new(),
                    status: String::new(),
                    source: "pracownicy".to_string(),
                });
                pracownicy_count += 1;
            }
        }
        self.log_message(format!("Wczytano {} PESEL-i", pracownicy_count));

        self.log_message("\nWczytywanie danych z pliku L4:".to_string());
        for row in range2.rows().skip(1) {
            if let Some(DataType::String(ubezpieczony)) = row.get(0).map(|c| c.to_owned()) {
                let pesel_regex = Regex::new(r".*\s(\d{11})$")?;
                if let Some(captures) = pesel_regex.captures(&ubezpieczony) {
                    let pesel = captures.get(1).unwrap().as_str().to_string();
                    let parts: Vec<&str> = ubezpieczony.split_whitespace().collect();
                    let nazwisko = parts.first().unwrap_or(&"").to_string();
                    let imie = parts.get(1).unwrap_or(&"").to_string();

                    let data_od_str = row.get(3).and_then(|c| c.get_string()).unwrap_or_default().to_string();
                    let data_do_str = row.get(4).and_then(|c| c.get_string()).unwrap_or_default().to_string();

                    let data_od = Self::parse_date(&data_od_str);
                    let data_do = Self::parse_date(&data_do_str);

                    data.push(EmployeeData {
                        nazwisko,
                        imie,
                        pesel,
                        data_od,
                        data_do,
                        na_opieke: row.get(5).and_then(|c| c.get_string()).unwrap_or_default().to_string(),
                        pobyt_w_szpitalu: row.get(6).and_then(|c| c.get_string()).unwrap_or_default().to_string(),
                        status: row.get(7).and_then(|c| c.get_string()).unwrap_or_default().to_string(),
                        source: "l4".to_string(),
                    });
                    l4_count += 1;
                }
            }
        }
        self.log_message(format!("Wczytano {} PESEL-i", l4_count));

        Ok((headers, data))
    }

    fn find_common_pesels(&mut self, data: &[EmployeeData]) -> HashSet<String> {
        let mut common_pesels = HashSet::new();
        let mut pracownicy_pesels = HashSet::new();
        let mut l4_pesels = HashSet::new();

        for employee in data {
            let source = &employee.source;
            if source == "pracownicy" {
                pracownicy_pesels.insert(employee.pesel.clone());
            } else {
                l4_pesels.insert(employee.pesel.clone());
            }
        }

        for pesel in &pracownicy_pesels {
            if l4_pesels.contains(pesel) {
                common_pesels.insert(pesel.clone());
            }
        }

        self.log_message(format!("\nLiczba wspólnych numerów PESEL: {}", common_pesels.len()));

        common_pesels
    }

    fn write_headers(&self, sheet: &mut Worksheet, headers: &Vec<String>, header_format: &Format) -> Result<()> {
        for (col, header) in headers.iter().enumerate() {
            sheet.write_string_with_format(0, col as u16, &header.to_string(), header_format)?;
        }
        Ok(())
    }

    fn write_data(&mut self, sheet: &mut Worksheet, data: &Vec<EmployeeData>, date_format: &Format) -> Result<()> {
        sheet.set_column_width(0, 20.0)?; // Nazwisko
        sheet.set_column_width(1, 15.0)?; // Imię
        sheet.set_column_width(2, 12.0)?; // PESEL
        sheet.set_column_width(3, 12.0)?; // Data od
        sheet.set_column_width(4, 12.0)?; // Data do
        sheet.set_column_width(5, 10.0)?; // Na opiekę
        sheet.set_column_width(6, 15.0)?; // Pobyt w szpitalu
        sheet.set_column_width(7, 12.0)?; // Status zaśw.

        let mut row = 1;
        for employee in data {
            if employee.source == "l4" {
                sheet.write_string(row, 0, &employee.nazwisko)?;
                sheet.write_string(row, 1, &employee.imie)?;
                sheet.write_string(row, 2, &employee.pesel)?;
                
                if let Some(excel_date) = employee.data_od {
                    sheet.write_number_with_format(row, 3, excel_date, date_format)?;
                } else {
                    sheet.write_blank(row, 3, date_format)?;
                }
                
                if let Some(excel_date) = employee.data_do {
                    sheet.write_number_with_format(row, 4, excel_date, date_format)?;
                } else {
                    sheet.write_blank(row, 4, date_format)?;
                }
                
                sheet.write_string(row, 5, &employee.na_opieke)?;
                sheet.write_string(row, 6, &employee.pobyt_w_szpitalu)?;
                sheet.write_string(row, 7, &employee.status)?;
                
                row += 1;
            }
        }
        Ok(())
    }
} 