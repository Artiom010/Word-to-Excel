import os
import re
import sqlite3
from tkinter import *
from tkinter import ttk, filedialog, messagebox
from docx import Document

# Funcțiile anterioare rămân neschimbate...

def delete_all_data():
    """
    Șterge toate datele din tabelul `produse` din baza de date SQLite.
    """
    conn = sqlite3.connect('facturi.db')
    cursor = conn.cursor()

    # Confirmăm înainte de ștergere
    confirm = messagebox.askyesno("Confirmare", "Sigur vrei să ștergi toate datele din tabel?")
    if confirm:
        cursor.execute("DELETE FROM produse")
        conn.commit()
        conn.close()
        messagebox.showinfo("Succes", "Toate datele au fost șterse din tabel.")
        refresh_data()  # Încărcăm din nou datele din baza de date și le afișăm în tabel
    else:
        messagebox.showinfo("Anulare", "Operațiunea a fost anulată.")

# Funcția pentru a adăuga un buton de ștergere în interfață
def add_delete_button():
    """
    Adaugă un buton pentru a șterge toate datele din tabelul `produse`.
    """
    delete_button = Button(root, text="Șterge toate datele", command=delete_all_data)
    delete_button.pack(pady=10)

# Codul pentru interfața principală rămâne aproape neschimbat...
# (asigură-te că apelul `add_delete_button()` este plasat într-un loc corespunzător în funcția principală)

def import_data():
    """
    Procesează fișierul selectat din directorul E:\word_toex.
    """
    file_path = entry_file_path.get()
    
    if not file_path:
        messagebox.showerror("Eroare", "Te rog selectează un fișier Word.")
        return

    # Previzualizarea și importarea datelor
    preview_and_import_file(file_path)

def refresh_data():
    """
    Încarcă și afișează datele din baza de date SQLite în tabelul Tkinter.
    """
    rows = load_data_from_db()
    for row in treeview.get_children():
        treeview.delete(row)  # Șterge datele anterioare
    for row in rows:
        treeview.insert("", "end", values=row)

# Creăm interfața Tkinter
root = Tk()
root.title("Import Facturi")

# Creăm un câmp pentru a introduce calea fișierului
Label(root, text="Selectează fișierul Word").pack(pady=10)
entry_file_path = Entry(root, width=50)
entry_file_path.pack(pady=5)

# Buton pentru a alege fișierul
Button(root, text="Browse", command=browse_file).pack(pady=10)

# Buton pentru importul datelor
Button(root, text="Importă Date", command=import_data).pack(pady=10)

# Adăugăm butonul pentru ștergerea datelor
add_delete_button()

# Creăm tabelul Tkinter pentru a vizualiza datele din SQLite
treeview = ttk.Treeview(root, columns=("ID", "Barcode", "Nume Produs", "Cantitate in Cutie", "Pret Unit", "Total"), show="headings")
treeview.heading("ID", text="ID")
treeview.heading("Barcode", text="Barcode")
treeview.heading("Nume Produs", text="Nume Produs")
treeview.heading("Cantitate in Cutie", text="Cantitate in Cutie")
treeview.heading("Pret Unit", text="Pret Unit")
treeview.heading("Total", text="Total")
treeview.pack(padx=10, pady=10)

# Încarcă și afișează datele din baza de date
refresh_data()

root.mainloop()
