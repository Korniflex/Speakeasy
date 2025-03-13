import tkinter as tk
from PIL import Image, ImageTk 
from tkinter import filedialog 
from tkinter import PhotoImage
import pandas as pd
import PyPDF2
import pytesseract
from docx import Document   
import os


import tkinter as tk
from PIL import Image, ImageTk

# Wir stellen den Root Fenster an
root = tk.Tk()                                                                  # Funktion die unser Hauptfenster erstellt
root.geometry("550x300")                                                        # wir definieren die groeße des Fensters
root.title("Speakeasy v0.1")                                                    # Wir definieren den Titel
root.resizable(width=True, height=True)                                         # Wir setzten eine dynamische Groeße ein : wenn wir ein großes Bild z.B zufuegen, wir das Fenster automatisch angepasst.

# Wir laden unser Speakeasy Logo !
logo_path = r"C:\Users\Admin\Desktop\Speakeasy\data\logo_Speakeasy_1080p.png"       # right clic, "copy Path", als string eingeben, tada.
logo = Image.open(logo_path)                                                        # Wir benutzen die Bibliothek PIL um den Bild zu oeffnen
logo = logo.resize((200, 200))                                                      # Wir skalieren das importierte bild mit .resize((length, width)) doppel Klammer !
global_logo = ImageTk.PhotoImage(logo)                                              # Wir konvertieren unser durch Pillow (PIL) geaendertes Bild in einen Format der von Tkinter uebernommen werden kann mit PhotoImage() 


# Herstekkung des Label, um das Bild anzeigen zu koennen.
label = tk.Label(root,                                                                                              # Wir setzten den Laben auf unser Hauptfenster
                 text='Willkommen bei Speakeasy.\nBitte wählen Sie eine Datei aus, um den Text zu extrahieren',     # Wir sind gut erzogen, sagen Hallo und wie wir behelflich sein koennen.
                 image=global_logo, 
                 compound='top')
label.pack(anchor=tk.CENTER)                                                                                        # Wir Positionnieren unser Label in der Mitte



# Funktion um den extrahierten Text zu speichern

def save_to_text(text, output_file= 'extracted.txt'):           # text ist die variabel die zugefuegt wird, extracted.txt wie wir den file benennen
    try:
        with open(output_file, "w", encoding="utf-8") as f:     # w steht fur write, 'utf-8' damit wir Akzents usw. mitnehmen koennen 
            f.write(text)                                       # schreibt unsere gespeicherte variabel text von read_file()
        print(f" Text gespeichert in {output_file}")
    except Exception as e:
        print(f" Fehler bei der Speicherung : {e}")