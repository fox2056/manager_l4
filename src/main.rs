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
        // Ustawienie stylu wizualnego
        let mut style = (*ctx.style()).clone();
        style.visuals.window_rounding = egui::Rounding::same(12.0);
        style.visuals.widgets.noninteractive.bg_fill = egui::Color32::from_rgb(245, 245, 245);
        style.visuals.widgets.hovered.bg_fill = egui::Color32::from_rgb(230, 240, 255);
        style.visuals.widgets.active.bg_fill = egui::Color32::from_rgb(200, 220, 255);
        ctx.set_style(style);

        // Główny panel
        egui::CentralPanel::default()
            .frame(egui::Frame::default()
                .fill(egui::Color32::from_rgb(250, 250, 252))
                .inner_margin(10.0))
            .show(ctx, |ui| {

                // Kontener na sekcje
                egui::Frame::none()
                    .fill(egui::Color32::WHITE)
                    .rounding(10.0)
                    .inner_margin(16.0)
                    .outer_margin(egui::Margin::symmetric(0.0, 10.0))
                    .shadow(egui::epaint::Shadow {
                        extrusion: 8.0,
                        color: egui::Color32::from_black_alpha(20),
                    })
                    .show(ui, |ui| {
                        ui.spacing_mut().item_spacing = egui::vec2(10.0, 15.0);

                        // Sekcja pierwszego pliku
                        ui.group(|ui| {
                            ui.horizontal(|ui| {
                                ui.label(egui::RichText::new("📋 Plik z listą pracowników").size(16.0));
                                if ui.add(
                                    egui::Button::new("Wybierz")
                                        .fill(egui::Color32::from_rgb(33, 150, 243))
                                        .rounding(6.0)
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
                                    ui.label(
                                        egui::RichText::new(path.file_name().unwrap().to_string_lossy().to_string())
                                            .color(egui::Color32::from_rgb(44, 62, 80))
                                    );
                                }
                                ui.add_space(ui.available_width());
                            });
                            ui.horizontal(|ui| {
                                ui.label(egui::RichText::new("Arkusz:").size(14.0));
                                egui::ComboBox::from_id_source("sheet1_combo")
                                    .width(ui.available_width() - 50.0)
                                    .selected_text(self.first_sheet.as_deref().unwrap_or("Wybierz arkusz"))
                                    .show_ui(ui, |ui| {
                                        for (idx, name) in self.available_sheets_1.iter().enumerate() {
                                            if ui.selectable_value(&mut self.selected_sheet_1_index, Some(idx), name).clicked() {
                                                self.first_sheet = Some(name.clone());
                                            }
                                        }
                                    });
                            });
                            ui.add_space(5.0);
                            ui.label(
                                egui::RichText::new("Wymagane kolumny: Nazwisko, Imię, Pesel")
                                    .size(12.0)
                                    .color(egui::Color32::from_rgb(120, 144, 156))
                            );
                        });

                        // Sekcja drugiego pliku
                        ui.group(|ui| {
                            ui.horizontal(|ui| {
                                ui.label(egui::RichText::new("📄 Plik z L4").size(16.0));
                                if ui.add(
                                    egui::Button::new("Wybierz")
                                        .fill(egui::Color32::from_rgb(33, 150, 243))
                                        .rounding(6.0)
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
                                    ui.label(
                                        egui::RichText::new(path.file_name().unwrap().to_string_lossy().to_string())
                                            .color(egui::Color32::from_rgb(44, 62, 80))
                                    );
                                }
                                ui.add_space(ui.available_width());
                            });
                            ui.horizontal(|ui| {
                                ui.label(egui::RichText::new("Arkusz:").size(14.0));
                                egui::ComboBox::from_id_source("sheet2_combo")
                                    .width(ui.available_width() - 50.0)
                                    .selected_text(self.second_sheet.as_deref().unwrap_or("Wybierz arkusz"))
                                    .show_ui(ui, |ui| {
                                        for (idx, name) in self.available_sheets_2.iter().enumerate() {
                                            if ui.selectable_value(&mut self.selected_sheet_2_index, Some(idx), name).clicked() {
                                                self.second_sheet = Some(name.clone());
                                            }
                                        }
                                    });
                            });
                            ui.add_space(5.0);
                            ui.label(
                                egui::RichText::new("Wymagane kolumny: Ubezpieczony	Seria i nr zaśw., Data wyst., Od, Do, Na opiekę, Pobyt w szpitalu, Status zaśw.")
                                    .size(12.0)
                                    .color(egui::Color32::from_rgb(120, 144, 156))
                            );
                        });

                        // Sekcja pliku wynikowego
                        ui.group(|ui| {
                            ui.horizontal(|ui| {
                                ui.label(egui::RichText::new("💾 Plik wynikowy").size(16.0));
                                if let Some(path) = &mut self.output_file {
                                    let mut text = path.to_string_lossy().to_string();
                                    ui.add(
                                        egui::TextEdit::singleline(&mut text)
                                            .desired_width(ui.available_width() - 50.0)
                                            .text_color(egui::Color32::from_rgb(44, 62, 80))
                                    );
                                    *path = PathBuf::from(text);
                                }
                                ui.add_space(ui.available_width());
                            });
                        });

                        // Przycisk uruchomienia
                        ui.add_space(5.0);
                        ui.vertical_centered(|ui| {
                            if ui.add(
                                egui::Button::new(
                                    egui::RichText::new("▶ Uruchom")
                                        .size(18.0)
                                        .color(egui::Color32::WHITE)
                                )
                                .fill(egui::Color32::from_rgb(46, 204, 113))
                                .min_size(egui::vec2(200.0, 50.0))
                                .rounding(8.0)
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

                // Logi
                ui.add_space(5.0);
                ui.label(egui::RichText::new("📜 Logi").size(16.0));
                egui::Frame::none()
                    .fill(egui::Color32::from_rgb(255, 255, 255))
                    .rounding(10.0)
                    .inner_margin(10.0)
                    .shadow(egui::epaint::Shadow {
                        extrusion: 4.0,
                        color: egui::Color32::from_black_alpha(10),
                    })
                    .show(ui, |ui| {
                        egui::ScrollArea::vertical()
                            .max_height(150.0)
                            .show(ui, |ui| {
                                ui.add(
                                    egui::TextEdit::multiline(&mut self.log.as_str())
                                        .desired_width(f32::INFINITY)
                                        .desired_rows(8)
                                        .text_color(egui::Color32::from_rgb(44, 62, 80))
                                        .font(egui::TextStyle::Monospace)
                                        .interactive(false)
                                );
                            });
                    });

                // Stopka
                ui.with_layout(egui::Layout::bottom_up(egui::Align::Center), |ui| {
                    ui.add_space(5.0); // Dodatkowy odstęp na dole
                    ui.horizontal(|ui| {
                        ui.label(
                            egui::RichText::new("© 2025 Oleksii Sliepov")
                                .size(12.0)
                                .color(egui::Color32::from_rgb(149, 165, 166))
                        );
                        ui.add_space(10.0);
                        if ui.add(
                            egui::Button::new(
                                egui::RichText::new("✉ sliepov@wp.pl")
                                    .size(12.0)
                                    .color(egui::Color32::from_rgb(33, 150, 243))
                            )
                            .fill(egui::Color32::TRANSPARENT)
                        ).clicked() {
                            if let Err(e) = open::that("mailto:sliepov@wp.pl") {
                                self.log_message(format!("Nie można otworzyć klienta poczty: {}", e));
                            }
                        }
                    });
                });
            });
    }
}

fn main() -> Result<(), eframe::Error> {
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_inner_size([700.0, 650.0])
            .with_min_inner_size([700.0, 650.0])
            .with_max_inner_size([700.0, 650.0])
            .with_resizable(false)
            .with_title("L4 Filter")
            .with_transparent(false)
            .with_decorations(true),
        ..Default::default()
    };

    eframe::run_native(
        "L4 Filter",
        options,
        Box::new(|cc| {
            cc.egui_ctx.set_visuals(egui::Visuals::light());
            Box::new(ExcelMergerApp::default())
        }),
    )
}