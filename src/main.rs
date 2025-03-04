#![windows_subsystem = "windows"]

mod excel_merger;

use eframe::egui;
use rfd::FileDialog;
use std::path::PathBuf;
use excel_merger::ExcelMerger;

struct ExcelMergerApp {
    first_file: Option<PathBuf>,
    second_file: Option<PathBuf>,
    output_file: Option<PathBuf>,
    first_sheet: Option<String>,
    second_sheet: Option<String>,
    available_sheets_1: Vec<String>,
    available_sheets_2: Vec<String>,
    selected_sheet_1_index: Option<usize>,
    selected_sheet_2_index: Option<usize>,
    log: String,
    merger: ExcelMerger,
}

impl Default for ExcelMergerApp {
    fn default() -> Self {
        let default_output = format!("L4_{}.xlsx", chrono::Local::now().format("%d-%m-%Y"));
        Self {
            first_file: None,
            second_file: None,
            output_file: Some(PathBuf::from(default_output)),
            first_sheet: None,
            second_sheet: None,
            available_sheets_1: Vec::new(),
            available_sheets_2: Vec::new(),
            selected_sheet_1_index: None,
            selected_sheet_2_index: None,
            log: String::new(),
            merger: ExcelMerger::new(),
        }
    }
}

impl ExcelMergerApp {
    fn log_message(&mut self, message: String) {
        self.log.push_str(&format!("{}\n", message));
    }
}

impl eframe::App for ExcelMergerApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        let frame_ui = egui::containers::Frame::default()
            .fill(egui::Color32::from_rgb(240, 240, 240))
            .inner_margin(20.0)
            .rounding(10.0)
            .shadow(egui::epaint::Shadow {
                extrusion: 5.0,
                color: egui::Color32::from_black_alpha(20),
            });

        egui::CentralPanel::default().frame(frame_ui).show(ctx, |ui| {
            // Pasek tytułowy
            ui.horizontal(|ui| {
                ui.add_space(10.0);
                ui.heading(egui::RichText::new("PS - L4")
                    .size(32.0)
                    .color(egui::Color32::from_rgb(41, 128, 185)));
            });
            ui.add_space(20.0);

            // Główny kontener
            egui::Frame::none()
                .fill(egui::Color32::from_rgb(255, 255, 255))
                .rounding(8.0)
                .inner_margin(16.0)
                .show(ui, |ui| {
                    // Sekcja pierwszego pliku
                    ui.group(|ui| {
                        ui.horizontal(|ui| {
                            ui.label(egui::RichText::new("Plik z listą pracowników:").size(16.0));
                            if ui.add(egui::Button::new(
                                egui::RichText::new("Wybierz plik z listą pracowników")
                                    .color(egui::Color32::WHITE)
                            ).fill(egui::Color32::from_rgb(52, 152, 219))
                            ).clicked() {
                                if let Some(path) = FileDialog::new()
                                    .add_filter("Excel", &["xlsx", "xls"])
                                    .pick_file() {
                                    self.available_sheets_1 = self.merger.get_sheet_names(&path);
                                    self.first_file = Some(path);
                                    self.selected_sheet_1_index = Some(0);
                                    if !self.available_sheets_1.is_empty() {
                                        self.first_sheet = Some(self.available_sheets_1[0].clone());
                                    }
                                }
                            }
                            if let Some(path) = &self.first_file {
                                ui.label(path.file_name().unwrap().to_string_lossy().to_string());
                            }
                        });

                        ui.horizontal(|ui| {
                            ui.label(egui::RichText::new("Wybierz arkusz:").size(16.0));
                            egui::ComboBox::from_id_source("sheet1_combo")
                                .selected_text(self.first_sheet.as_deref().unwrap_or(""))
                                .show_ui(ui, |ui| {
                                    for (idx, name) in self.available_sheets_1.iter().enumerate() {
                                        if ui.selectable_value(&mut self.selected_sheet_1_index, Some(idx), name).clicked() {
                                            self.first_sheet = Some(name.clone());
                                        }
                                    }
                                });
                        });
                    });

                    ui.add_space(10.0);

                    // Sekcja drugiego pliku
                    ui.group(|ui| {
                        ui.horizontal(|ui| {
                            ui.label(egui::RichText::new("Plik z L4:").size(16.0));
                            if ui.add(egui::Button::new(
                                egui::RichText::new("Wybierz plik z L4")
                                    .color(egui::Color32::WHITE)
                            ).fill(egui::Color32::from_rgb(52, 152, 219))
                            ).clicked() {
                                if let Some(path) = FileDialog::new()
                                    .add_filter("Excel", &["xlsx", "xls"])
                                    .pick_file() {
                                    self.available_sheets_2 = self.merger.get_sheet_names(&path);
                                    self.second_file = Some(path);
                                    self.selected_sheet_2_index = Some(0);
                                    if !self.available_sheets_2.is_empty() {
                                        self.second_sheet = Some(self.available_sheets_2[0].clone());
                                    }
                                }
                            }
                            if let Some(path) = &self.second_file {
                                ui.label(path.file_name().unwrap().to_string_lossy().to_string());
                            }
                        });

                        ui.horizontal(|ui| {
                            ui.label(egui::RichText::new("Wybierz arkusz:").size(16.0));
                            egui::ComboBox::from_id_source("sheet2_combo")
                                .selected_text(self.second_sheet.as_deref().unwrap_or(""))
                                .show_ui(ui, |ui| {
                                    for (idx, name) in self.available_sheets_2.iter().enumerate() {
                                        if ui.selectable_value(&mut self.selected_sheet_2_index, Some(idx), name).clicked() {
                                            self.second_sheet = Some(name.clone());
                                        }
                                    }
                                });
                        });
                    });

                    ui.add_space(10.0);

                    // Sekcja pliku wynikowego
                    ui.group(|ui| {
                        ui.horizontal(|ui| {
                            ui.label(egui::RichText::new("Nazwa pliku wynikowego:").size(16.0));
                            if let Some(path) = &mut self.output_file {
                                let mut text = path.to_string_lossy().to_string();
                                if ui.add(egui::TextEdit::singleline(&mut text)
                                    .desired_width(300.0)
                                    .margin(egui::vec2(5.0, 5.0))
                                ).changed() {
                                    *path = PathBuf::from(text);
                                }
                            }
                        });
                    });

                    ui.add_space(20.0);

                    // Przycisk uruchomienia
                    ui.vertical_centered(|ui| {
                        if ui.add(egui::Button::new(
                            egui::RichText::new("Uruchom")
                                .size(18.0)
                                .color(egui::Color32::WHITE)
                            )
                            .min_size(egui::vec2(200.0, 40.0))
                            .fill(egui::Color32::from_rgb(46, 204, 113))
                        ).clicked() {
                            if let (Some(first), Some(second), Some(output), Some(first_sheet), Some(second_sheet)) = (
                                &self.first_file,
                                &self.second_file,
                                &self.output_file,
                                &self.first_sheet,
                                &self.second_sheet,
                            ) {
                                if let Err(e) = self.merger.merge_files(first, second, output, first_sheet, second_sheet) {
                                    self.log_message(format!("Błąd: {}", e));
                                }
                                // Aktualizacja logów z mergera
                                let messages = self.merger.messages.clone();
                                for message in messages {
                                    self.log_message(message);
                                }
                                self.merger.messages.clear();
                            } else {
                                self.log_message("Proszę wybrać wszystkie wymagane pliki i arkusze.".to_string());
                            }
                        }
                    });
                });

            ui.add_space(20.0);

            // Sekcja logów
            ui.group(|ui| {
                ui.label(egui::RichText::new("Logi:").size(16.0));
                egui::ScrollArea::vertical()
                    .max_height(200.0)
                    .show(ui, |ui| {
                        ui.add(
                            egui::TextEdit::multiline(&mut self.log.as_str())
                                .desired_width(f32::INFINITY)
                                .desired_rows(10)
                                .text_color(egui::Color32::from_rgb(44, 62, 80))
                                .font(egui::TextStyle::Monospace)
                        );
                    });
            });

            // Stopka
            ui.with_layout(egui::Layout::bottom_up(egui::Align::Center), |ui| {
                ui.add_space(10.0);
                ui.label(
                    egui::RichText::new("Aplikację stworzył: Oleksii Sliepov")
                        .size(14.0)
                        .color(egui::Color32::from_rgb(127, 140, 141))
                );
                ui.add_space(5.0);
            });
        });
    }
}

fn main() -> Result<(), eframe::Error> {
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_inner_size([800.0, 900.0])
            .with_min_inner_size([600.0, 400.0])
            .with_title("PS - L4")
            .with_transparent(false)
            .with_decorations(true),
        ..Default::default()
    };
    
    eframe::run_native(
        "PS - L4",
        options,
        Box::new(|cc| {
            cc.egui_ctx.set_visuals(egui::Visuals::light());
            Box::new(ExcelMergerApp::default())
        }),
    )
}

