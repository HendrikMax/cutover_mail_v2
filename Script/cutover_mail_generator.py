"""
Cutover E-Mail Generator - Hauptprogramm mit GUI
Erstellt automatisch E-Mails aus Excel-Cutoverplan
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
import threading
from datetime import datetime
from typing import Optional

import excel_parser
import email_generator
import config


class CutoverMailGeneratorGUI:
    """
    Hauptfenster der Anwendung.
    """

    def __init__(self):
        """Initialisiert GUI-Elemente."""
        self.root = tk.Tk()
        self.root.title(config.GUI_TITLE)
        self.root.geometry(f"{config.GUI_WIDTH}x{config.GUI_HEIGHT}")

        # Variablen
        self.excel_file_path = tk.StringVar()
        self.selected_sheet = tk.StringVar()
        self.cutover_ident = tk.StringVar()
        self.output_path = tk.StringVar()
        self.filter_status = tk.StringVar(value="Alle")
        self.filter_ident = tk.StringVar(value="Alle")

        # Sheet-Namen Cache
        self.available_sheets = []
        self.available_status = ["Alle"]
        self.available_idents = ["Alle"]

        # GUI aufbauen
        self._create_widgets()

    def _create_widgets(self):
        """Erstellt alle GUI-Elemente."""
        # Hauptframe mit Padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Grid-Konfiguration für responsives Layout
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        row = 0

        # Excel-Dateiauswahl
        ttk.Label(main_frame, text="Excel-Datei:", font=('Arial', 10, 'bold')).grid(
            row=row, column=0, sticky=tk.W, pady=5
        )
        row += 1

        file_frame = ttk.Frame(main_frame)
        file_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        file_frame.columnconfigure(0, weight=1)

        ttk.Entry(file_frame, textvariable=self.excel_file_path, state='readonly').grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5)
        )
        ttk.Button(file_frame, text="Durchsuchen...", command=self.browse_excel_file).grid(
            row=0, column=1
        )
        row += 1

        # Tabellenblatt
        ttk.Label(main_frame, text="Tabellenblatt:").grid(
            row=row, column=0, sticky=tk.W, pady=5
        )
        row += 1

        self.sheet_combo = ttk.Combobox(
            main_frame,
            textvariable=self.selected_sheet,
            state='readonly',
            width=40
        )
        self.sheet_combo.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        self.sheet_combo.bind('<<ComboboxSelected>>', lambda e: self.load_filter_options())
        row += 1

        # Cutover-Ident
        ttk.Label(main_frame, text="Cutover-Ident:").grid(
            row=row, column=0, sticky=tk.W, pady=5
        )
        row += 1

        ident_frame = ttk.Frame(main_frame)
        ident_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        ttk.Entry(ident_frame, textvariable=self.cutover_ident, width=30).grid(
            row=0, column=0, sticky=tk.W
        )
        ttk.Label(ident_frame, text='  (z.B. "JOSEF", "DPN_ECH")', foreground='gray').grid(
            row=0, column=1, sticky=tk.W
        )
        row += 1

        # Separator
        ttk.Separator(main_frame, orient='horizontal').grid(
            row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10
        )
        row += 1

        # Ausgabepfad für E-Mail-Dateien
        ttk.Label(main_frame, text="Ausgabepfad für E-Mail-Dateien:", font=('Arial', 10, 'bold')).grid(
            row=row, column=0, sticky=tk.W, pady=5
        )
        row += 1

        output_frame = ttk.Frame(main_frame)
        output_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        output_frame.columnconfigure(0, weight=1)

        ttk.Entry(output_frame, textvariable=self.output_path, state='readonly').grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5)
        )

        ttk.Button(
            output_frame,
            text="Durchsuchen...",
            command=self.browse_output_folder
        ).grid(row=0, column=1)
        row += 1

        # Info-Text
        info_label = ttk.Label(
            main_frame,
            text="  Hinweis: EML-Dateien können in Outlook per Doppelklick geöffnet werden",
            foreground='gray',
            font=('Arial', 8)
        )
        info_label.grid(row=row, column=0, columnspan=3, sticky=tk.W, pady=(0, 5))
        row += 1

        # Separator
        ttk.Separator(main_frame, orient='horizontal').grid(
            row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10
        )
        row += 1

        # Filter-Optionen
        ttk.Label(main_frame, text="Filter-Optionen:", font=('Arial', 10, 'bold')).grid(
            row=row, column=0, sticky=tk.W, pady=5
        )
        row += 1

        filter_frame = ttk.Frame(main_frame)
        filter_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        # Status-Filter
        ttk.Label(filter_frame, text="IST-Status filtern:").grid(row=0, column=0, sticky=tk.W)
        self.status_combo = ttk.Combobox(
            filter_frame,
            textvariable=self.filter_status,
            values=self.available_status,
            state='readonly',
            width=30
        )
        self.status_combo.grid(row=0, column=1, sticky=tk.W, padx=(10, 0))

        # Ident-Filter
        ttk.Label(filter_frame, text="Aktivitäts-Ident filtern:").grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
        self.ident_combo = ttk.Combobox(
            filter_frame,
            textvariable=self.filter_ident,
            values=self.available_idents,
            state='readonly',
            width=30
        )
        self.ident_combo.grid(row=1, column=1, sticky=tk.W, padx=(10, 0), pady=(5, 0))
        row += 1

        # Separator
        ttk.Separator(main_frame, orient='horizontal').grid(
            row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10
        )
        row += 1

        # Button: E-Mails generieren
        self.generate_button = ttk.Button(
            main_frame,
            text="E-Mails generieren",
            command=self.generate_emails,
            style='Accent.TButton'
        )
        self.generate_button.grid(row=row, column=0, columnspan=3, pady=10)
        row += 1

        # Fortschrittsbalken
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            main_frame,
            variable=self.progress_var,
            maximum=100
        )
        self.progress_bar.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        row += 1

        # Fortschritts-Label
        self.progress_label = ttk.Label(main_frame, text="Bereit")
        self.progress_label.grid(row=row, column=0, columnspan=3, pady=2)
        row += 1

        # Log-Ausgabe
        ttk.Label(main_frame, text="Log:", font=('Arial', 10, 'bold')).grid(
            row=row, column=0, sticky=tk.W, pady=5
        )
        row += 1

        self.log_text = scrolledtext.ScrolledText(
            main_frame,
            height=10,
            width=70,
            state='disabled',
            wrap=tk.WORD
        )
        self.log_text.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        main_frame.rowconfigure(row, weight=1)

    def browse_excel_file(self):
        """Öffnet Datei-Dialog für Excel-Auswahl."""
        filename = filedialog.askopenfilename(
            title="Excel-Datei auswählen",
            filetypes=[("Excel-Dateien", "*.xlsx *.xls"), ("Alle Dateien", "*.*")]
        )

        if filename:
            self.excel_file_path.set(filename)
            self.load_sheet_names()
            self.log_message(f"Excel-Datei geladen: {Path(filename).name}")

    def load_sheet_names(self):
        """Lädt Tabellenblatt-Namen aus gewählter Datei."""
        file_path = self.excel_file_path.get()
        if not file_path:
            return

        try:
            # Tabellenblätter laden
            self.available_sheets = excel_parser.get_sheet_names(file_path)
            self.sheet_combo['values'] = self.available_sheets

            # Erstes Blatt automatisch auswählen
            if self.available_sheets:
                self.selected_sheet.set(self.available_sheets[0])
                self.load_filter_options()

        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Laden der Tabellenblätter:\n{e}")
            self.log_message(f"FEHLER: {e}")

    def load_filter_options(self):
        """Lädt verfügbare Idents und Status für Filter."""
        file_path = self.excel_file_path.get()
        sheet_name = self.selected_sheet.get()

        if not file_path or not sheet_name:
            return

        try:
            # Idents laden
            idents = excel_parser.get_unique_values(
                file_path,
                sheet_name,
                config.EXCEL_COLUMNS['ident']
            )
            self.available_idents = ["Alle"] + idents
            self.ident_combo['values'] = self.available_idents

            # Status-Werte laden
            status_values = excel_parser.get_unique_values(
                file_path,
                sheet_name,
                config.EXCEL_COLUMNS['ist_status']
            )
            self.available_status = ["Alle"] + status_values
            self.status_combo['values'] = self.available_status

        except:
            pass  # Fehler ignorieren, Filter ist optional

    def browse_output_folder(self):
        """Öffnet Ordner-Dialog für Ausgabepfad."""
        folder = filedialog.askdirectory(title="Ausgabeordner auswählen")
        if folder:
            self.output_path.set(folder)
            self.log_message(f"Ausgabeordner: {folder}")

    def validate_inputs(self) -> bool:
        """
        Validiert alle Eingabefelder.

        Returns:
            True wenn alle Eingaben gültig, False sonst
        """
        # Excel-Datei
        if not self.excel_file_path.get():
            messagebox.showwarning("Validierung", "Bitte wählen Sie eine Excel-Datei aus.")
            return False

        # Tabellenblatt
        if not self.selected_sheet.get():
            messagebox.showwarning("Validierung", "Bitte wählen Sie ein Tabellenblatt aus.")
            return False

        # Cutover-Ident
        if not self.cutover_ident.get().strip():
            messagebox.showwarning("Validierung", "Bitte geben Sie eine Cutover-Ident ein.")
            return False

        # Ausgabepfad
        if not self.output_path.get():
            messagebox.showwarning(
                "Validierung",
                "Bitte wählen Sie einen Ausgabeordner für die E-Mail-Dateien."
            )
            return False

        return True

    def generate_emails(self):
        """Startet E-Mail-Generierung in separatem Thread."""
        if not self.validate_inputs():
            return

        # Button deaktivieren während Verarbeitung
        self.generate_button.config(state='disabled')
        self.progress_var.set(0)
        self.progress_label.config(text="Verarbeitung läuft...")

        # In separatem Thread ausführen, damit GUI responsiv bleibt
        thread = threading.Thread(target=self._generate_emails_thread, daemon=True)
        thread.start()

    def _generate_emails_thread(self):
        """Thread-Funktion für E-Mail-Generierung."""
        try:
            # Parameter sammeln
            file_path = self.excel_file_path.get()
            sheet_name = self.selected_sheet.get()
            cutover_id = self.cutover_ident.get().strip()
            output_path = self.output_path.get()

            # Filter vorbereiten
            filters = {}

            # Status-Filter anwenden (nur wenn nicht "Alle")
            status_filter = self.filter_status.get()
            if status_filter != "Alle":
                filters[config.EXCEL_COLUMNS['ist_status']] = status_filter

            # Ident-Filter anwenden (nur wenn nicht "Alle")
            ident_filter = self.filter_ident.get()
            if ident_filter != "Alle":
                filters[config.EXCEL_COLUMNS['ident']] = ident_filter

            # Aktivitäten laden
            self.log_message(f"\nLade Aktivitäten aus: {Path(file_path).name}")
            self.log_message(f"Tabellenblatt: {sheet_name}")

            activities = excel_parser.load_activities(file_path, sheet_name, filters)

            self.log_message(f"{len(activities)} Aktivitäten gefunden")

            if len(activities) == 0:
                self.log_message("WARNUNG: Keine Aktivitäten zum Verarbeiten gefunden!")
                messagebox.showinfo(
                    "Keine Aktivitäten",
                    "Es wurden keine Aktivitäten gefunden, die den Filterkriterien entsprechen."
                )
                return

            # E-Mails generieren
            self.log_message(f"\nStarte E-Mail-Generierung (EML-Dateien)")
            self.log_message("-" * 50)

            stats = email_generator.generate_emails(
                activities,
                cutover_id,
                'msg',
                output_path,
                self.update_progress
            )

            # Ergebnis anzeigen
            self.log_message("-" * 50)
            self.log_message(f"\nFertig!")
            self.log_message(f"Erfolgreich: {stats['erfolg']}")
            self.log_message(f"Fehler: {stats['fehler']}")
            self.log_message(f"Gesamt: {stats['gesamt']}")

            # Fehlerdetails, falls vorhanden
            if stats['fehler'] > 0:
                self.log_message("\nFehler-Details:")
                for fehler in stats['fehler_liste']:
                    self.log_message(
                        f"  - {fehler['ident']} ({fehler['email']}): {fehler['fehler']}"
                    )

            # Success-Dialog
            messagebox.showinfo(
                "E-Mail-Generierung abgeschlossen",
                f"E-Mails wurden erfolgreich generiert!\n\n"
                f"Erfolgreich: {stats['erfolg']}\n"
                f"Fehler: {stats['fehler']}\n"
                f"Gesamt: {stats['gesamt']}"
            )

        except Exception as e:
            self.log_message(f"\nKRITISCHER FEHLER: {e}")
            messagebox.showerror("Fehler", f"Fehler bei der Verarbeitung:\n\n{e}")

        finally:
            # Button wieder aktivieren
            self.generate_button.config(state='normal')
            self.progress_label.config(text="Bereit")
            self.progress_var.set(0)

    def update_progress(self, current: int, total: int, message: str):
        """
        Aktualisiert Fortschrittsbalken und Log.

        Args:
            current: Aktueller Fortschritt
            total: Gesamtanzahl
            message: Log-Nachricht
        """
        # Fortschritt berechnen
        progress = (current / total) * 100
        self.progress_var.set(progress)
        self.progress_label.config(text=f"{current}/{total}")

        # Log-Nachricht hinzufügen
        self.log_message(message)

    def log_message(self, message: str):
        """
        Fügt Nachricht zum Log hinzu.

        Args:
            message: Log-Nachricht
        """
        timestamp = datetime.now().strftime(config.LOG_DATE_FORMAT)

        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)  # Scroll zum Ende
        self.log_text.config(state='disabled')

    def run(self):
        """Startet GUI-Hauptschleife."""
        self.log_message("Cutover E-Mail Generator gestartet")
        self.log_message("Bitte wählen Sie eine Excel-Datei aus")
        self.root.mainloop()


def main():
    """Hauptfunktion."""
    app = CutoverMailGeneratorGUI()
    app.run()


if __name__ == "__main__":
    main()
