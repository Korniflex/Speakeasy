# Der Benutzer muss ein Bild/ PDF hochladen koennen.
from googletrans import Translator
import tkinter as tk
from tkinter import filedialog
from PIL import Image
import pytesseract

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

#STEP  1 TKINTER
#eine User Interface erstellen mit Tkinter mit:
# Hauptfenster (Logo?)
# Knopf um ein Bild hochzuladen
# 
#Step2 PYTESSERACT
# Bild zu String verwandeln
#String mit Tkinter in einen neues Fenster asugeben? (Dienst nur fuers testen). Dannach wird es uebersetzt.

