import pandas as pd
from tkinter import Tk, messagebox
from tkinter.filedialog import askopenfilename, asksaveasfilename
import sys  # sys importieren

def convert_csv_to_xlsx(csv_file, xlsx_file):
    # Lese die CSV-Datei ein
    df = pd.read_csv(csv_file)

    # Exportiere die Daten in eine XLSX-Datei
    df.to_excel(xlsx_file, index=False)

if __name__ == "__main__":
    # Initialisiere tkinter und verstecke das Hauptfenster
    Tk().withdraw()

    # Zeige eine Beschreibung für die Dateiauswahl an
    messagebox.showinfo("CSV zu XLSX Konverter", "Bitte wähle die CSV-Datei aus, die du konvertieren möchtest.")

    # Wähle die CSV-Datei aus
    csv_file = askopenfilename(title="Wähle die CSV-Datei aus", filetypes=[("CSV-Dateien", "*.csv")])
    if not csv_file:
        print("Keine Datei ausgewählt. Programm wird beendet.")
        sys.exit()  # Verwende sys.exit()

    # Zeige eine Beschreibung für die Speicherortauswahl an
    messagebox.showinfo("Speichern der XLSX-Datei", "Bitte wähle den Speicherort und den Namen für die konvertierte XLSX-Datei.")

    # Wähle den Speicherort für die XLSX-Datei
    xlsx_file = asksaveasfilename(title="Speichere die XLSX-Datei", defaultextension=".xlsx", filetypes=[("Excel-Dateien", "*.xlsx")])
    if not xlsx_file:
        print("Keine Datei ausgewählt. Programm wird beendet.")
        sys.exit()  # Verwende sys.exit()

    # Konvertiere die CSV-Datei in XLSX
    convert_csv_to_xlsx(csv_file, xlsx_file)
    print(f'Die Datei wurde erfolgreich konvertiert: {xlsx_file}')
