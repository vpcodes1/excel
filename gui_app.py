#!/usr/bin/env python3
"""
GUI aplikacija za ažuriranje cena u Excel fajlovima.
Jednostavna drag-and-drop aplikacija sa grafičkim interfejsom.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from tkinterdnd2 import DND_FILES, TkinterDnD
import threading


class ExcelPriceUpdaterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Ažuriranje cena u Excel-u")
        self.root.geometry("700x650")
        self.root.resizable(False, False)

        # Varijable
        self.file_path = tk.StringVar()
        self.coefficient = tk.DoubleVar(value=1.1)
        self.selected_sheets = []
        self.all_sheets = []

        # Stilizovanje
        style = ttk.Style()
        style.theme_use('clam')

        self.create_widgets()

    def create_widgets(self):
        # Header
        header_frame = tk.Frame(self.root, bg="#366092", height=80)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)

        title_label = tk.Label(
            header_frame,
            text="📊 Ažuriranje Cena u Excel-u",
            font=("Arial", 18, "bold"),
            bg="#366092",
            fg="white"
        )
        title_label.pack(expand=True)

        # Main container
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Fajl sekcija
        file_frame = tk.LabelFrame(main_frame, text="1. Odaberi Excel fajl", font=("Arial", 10, "bold"), padx=10, pady=10)
        file_frame.pack(fill=tk.X, pady=(0, 15))

        file_entry_frame = tk.Frame(file_frame)
        file_entry_frame.pack(fill=tk.X)

        self.file_entry = tk.Entry(file_entry_frame, textvariable=self.file_path, font=("Arial", 10), state="readonly")
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        browse_btn = tk.Button(
            file_entry_frame,
            text="📁 Odaberi",
            command=self.browse_file,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=15,
            pady=5,
            relief=tk.RAISED,
            cursor="hand2"
        )
        browse_btn.pack(side=tk.RIGHT)

        # Drag and drop info
        dnd_label = tk.Label(file_frame, text="ili prevuci fajl ovde", font=("Arial", 9, "italic"), fg="gray")
        dnd_label.pack(pady=(5, 0))

        # Sheets sekcija
        sheets_frame = tk.LabelFrame(main_frame, text="2. Odaberi sheet-ove (opciono)", font=("Arial", 10, "bold"), padx=10, pady=10)
        sheets_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        sheets_info = tk.Label(
            sheets_frame,
            text="Ako ne odabereš ništa, ažuriraće se SVI sheet-ovi",
            font=("Arial", 9),
            fg="gray"
        )
        sheets_info.pack(anchor=tk.W, pady=(0, 5))

        self.sheets_listbox = tk.Listbox(
            sheets_frame,
            selectmode=tk.MULTIPLE,
            font=("Arial", 10),
            height=6,
            exportselection=False
        )
        self.sheets_listbox.pack(fill=tk.BOTH, expand=True)

        # Koeficijent sekcija
        coef_frame = tk.LabelFrame(main_frame, text="3. Unesi koeficijent", font=("Arial", 10, "bold"), padx=10, pady=10)
        coef_frame.pack(fill=tk.X, pady=(0, 15))

        coef_input_frame = tk.Frame(coef_frame)
        coef_input_frame.pack(fill=tk.X)

        tk.Label(coef_input_frame, text="Koeficijent:", font=("Arial", 10)).pack(side=tk.LEFT, padx=(0, 10))

        self.coef_entry = tk.Entry(coef_input_frame, textvariable=self.coefficient, font=("Arial", 10), width=10)
        self.coef_entry.pack(side=tk.LEFT, padx=(0, 20))

        # Primeri
        examples_text = "Primeri:  1.1 (+10%)  |  1.25 (+25%)  |  0.9 (-10%)  |  0.8 (-20%)"
        tk.Label(coef_input_frame, text=examples_text, font=("Arial", 9), fg="gray").pack(side=tk.LEFT)

        # Dugmad
        buttons_frame = tk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=(0, 10))

        self.update_btn = tk.Button(
            buttons_frame,
            text="✅ Ažuriraj Cene",
            command=self.update_prices,
            bg="#2196F3",
            fg="white",
            font=("Arial", 12, "bold"),
            padx=20,
            pady=10,
            relief=tk.RAISED,
            cursor="hand2",
            state=tk.DISABLED
        )
        self.update_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))

        self.preview_btn = tk.Button(
            buttons_frame,
            text="👁 Pregled",
            command=self.preview_changes,
            bg="#FF9800",
            fg="white",
            font=("Arial", 12, "bold"),
            padx=20,
            pady=10,
            relief=tk.RAISED,
            cursor="hand2",
            state=tk.DISABLED
        )
        self.preview_btn.pack(side=tk.RIGHT, expand=True, fill=tk.X, padx=(5, 0))

        # Status bar
        self.status_label = tk.Label(
            self.root,
            text="Spremno - Odaberi Excel fajl za početak",
            font=("Arial", 9),
            bg="#f0f0f0",
            anchor=tk.W,
            padx=10,
            pady=5
        )
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Odaberi Excel fajl",
            filetypes=[("Excel fajlovi", "*.xlsx *.xlsm"), ("Svi fajlovi", "*.*")]
        )
        if filename:
            self.load_file(filename)

    def load_file(self, filepath):
        self.file_path.set(filepath)
        self.status_label.config(text=f"Učitavam fajl: {os.path.basename(filepath)}")

        try:
            workbook = load_workbook(filename=filepath, read_only=True)
            self.all_sheets = workbook.sheetnames
            workbook.close()

            # Popuni listbox
            self.sheets_listbox.delete(0, tk.END)
            for sheet in self.all_sheets:
                self.sheets_listbox.insert(tk.END, sheet)

            # Omogući dugmad
            self.update_btn.config(state=tk.NORMAL)
            self.preview_btn.config(state=tk.NORMAL)

            self.status_label.config(text=f"Učitan fajl: {os.path.basename(filepath)} ({len(self.all_sheets)} sheet-ova)")

        except Exception as e:
            messagebox.showerror("Greška", f"Greška pri učitavanju fajla:\n{str(e)}")
            self.status_label.config(text="Greška pri učitavanju fajla")

    def get_selected_sheets(self):
        selected_indices = self.sheets_listbox.curselection()
        if not selected_indices:
            return None  # Svi sheet-ovi
        return [self.all_sheets[i] for i in selected_indices]

    def validate_inputs(self):
        if not self.file_path.get():
            messagebox.showwarning("Upozorenje", "Molim te odaberi Excel fajl!")
            return False

        try:
            coef = self.coefficient.get()
            if coef <= 0:
                raise ValueError("Koeficijent mora biti pozitivan!")
        except:
            messagebox.showwarning("Upozorenje", "Neispravan koeficijent! Unesi pozitivan broj.")
            return False

        return True

    def preview_changes(self):
        if not self.validate_inputs():
            return

        selected_sheets = self.get_selected_sheets()
        sheets_text = "SVIH sheet-ova" if selected_sheets is None else f"sheet-ova: {', '.join(selected_sheets)}"

        # Otvori preview prozor
        preview_window = tk.Toplevel(self.root)
        preview_window.title("Pregled promena")
        preview_window.geometry("600x500")

        tk.Label(
            preview_window,
            text=f"Ažuriranje {sheets_text}",
            font=("Arial", 12, "bold"),
            pady=10
        ).pack()

        tk.Label(
            preview_window,
            text=f"Koeficijent: {self.coefficient.get()}",
            font=("Arial", 10),
            pady=5
        ).pack()

        text_area = scrolledtext.ScrolledText(preview_window, font=("Courier", 9), wrap=tk.WORD)
        text_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Učitaj preview
        try:
            workbook = load_workbook(filename=self.file_path.get())
            sheets_to_check = selected_sheets if selected_sheets else workbook.sheetnames

            for sheet_name in sheets_to_check:
                sheet = workbook[sheet_name]
                text_area.insert(tk.END, f"\n{'='*50}\n")
                text_area.insert(tk.END, f"Sheet: {sheet_name}\n")
                text_area.insert(tk.END, f"{'='*50}\n\n")

                count = 0
                for row in sheet.iter_rows(min_row=2, max_row=min(10, sheet.max_row), values_only=False):
                    for cell in row:
                        if isinstance(cell.value, (int, float)) and cell.value:
                            old_val = cell.value
                            new_val = old_val * self.coefficient.get()
                            text_area.insert(tk.END, f"  {cell.coordinate}: {old_val} → {new_val:.2f}\n")
                            count += 1

                if count == 0:
                    text_area.insert(tk.END, "  (Nema numeričkih vrednosti)\n")

                text_area.insert(tk.END, f"\n  Ažurirano: {count} ćelija (prikazano prvih 10 redova)\n")

            workbook.close()
            text_area.config(state=tk.DISABLED)

        except Exception as e:
            text_area.insert(tk.END, f"\nGreška: {str(e)}")

    def update_prices(self):
        if not self.validate_inputs():
            return

        selected_sheets = self.get_selected_sheets()
        sheets_text = "SVE sheet-ove" if selected_sheets is None else f"odabrane sheet-ove ({len(selected_sheets)})"

        # Potvrda
        confirm = messagebox.askyesno(
            "Potvrda",
            f"Ažuriraću cene u {sheets_text} sa koeficijentom {self.coefficient.get()}.\n\n"
            f"Da li želiš da nastaviš?\n\n"
            f"Preporučujem da prvo napraviš kopiju fajla!",
            icon=messagebox.WARNING
        )

        if not confirm:
            return

        # Onemogući dugmad tokom obrade
        self.update_btn.config(state=tk.DISABLED)
        self.preview_btn.config(state=tk.DISABLED)
        self.status_label.config(text="Ažuriram cene...")

        # Pokreni u zasebnom thread-u
        thread = threading.Thread(target=self._update_prices_thread, args=(selected_sheets,))
        thread.start()

    def _update_prices_thread(self, selected_sheets):
        try:
            workbook = load_workbook(filename=self.file_path.get())
            sheets_to_update = selected_sheets if selected_sheets else workbook.sheetnames

            total_updated = 0

            for sheet_name in sheets_to_update:
                sheet = workbook[sheet_name]
                for row in sheet.iter_rows(min_row=2):
                    for cell in row:
                        if isinstance(cell.value, (int, float)) and cell.value:
                            cell.value = cell.value * self.coefficient.get()
                            total_updated += 1

            # Sačuvaj
            workbook.save(self.file_path.get())
            workbook.close()

            # Uspeh
            self.root.after(0, lambda: self._update_complete(total_updated))

        except Exception as e:
            self.root.after(0, lambda: self._update_error(str(e)))

    def _update_complete(self, count):
        self.update_btn.config(state=tk.NORMAL)
        self.preview_btn.config(state=tk.NORMAL)
        self.status_label.config(text=f"Uspešno ažurirano {count} ćelija!")

        messagebox.showinfo(
            "Uspeh!",
            f"Cene su uspešno ažurirane!\n\n"
            f"Ukupno ažuriranih ćelija: {count}\n\n"
            f"Fajl je sačuvan."
        )

    def _update_error(self, error):
        self.update_btn.config(state=tk.NORMAL)
        self.preview_btn.config(state=tk.NORMAL)
        self.status_label.config(text="Greška pri ažuriranju")

        messagebox.showerror("Greška", f"Greška pri ažuriranju cena:\n{error}")


def main():
    try:
        # Pokušaj sa drag-and-drop podrškom
        root = TkinterDnD.Tk()
    except:
        # Ako ne radi, koristi obični Tk
        root = tk.Tk()

    app = ExcelPriceUpdaterGUI(root)

    # Centriraj prozor
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')

    root.mainloop()


if __name__ == '__main__':
    main()
