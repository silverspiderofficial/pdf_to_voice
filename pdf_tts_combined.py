import os
import urllib.request
import subprocess
import win32com.client
import fitz  # pip install pymupdf
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
from webbrowser import Chrome

# ------------------ بخش نصب و بررسی صدا ------------------

def has_voice(keywords):
    """بررسی وجود صدایی که شامل کلیدواژه‌ها باشد."""
    try:
        speaker = win32com.client.Dispatch("SAPI.SpVoice")
        for voice in speaker.GetVoices():
            desc = voice.GetDescription().lower()
            if any(k.lower() in desc for k in keywords):
                return True
    except Exception:
        return False
    return False





# ------------------ ابزارهای متن و صدا ------------------

def extract_text_from_pdf(pdf_path):
    """استخراج متن از PDF."""
    doc = fitz.open(pdf_path)
    return "\n".join(page.get_text() for page in doc).strip()

def detect_language(text):
    """تشخیص ساده زبان متن."""
    fa_count = sum(1 for ch in text if '\u0600' <= ch <= '\u06FF')
    en_count = sum(1 for ch in text if ch.isascii() and ch.isalpha())
    return "fa" if fa_count > en_count else "en"

def list_voices():
    """لیست همه صداهای نصب‌شده."""
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    return [voice.GetDescription() for voice in speaker.GetVoices()]

def speak_to_wav(text, voice_name, wav_path):
    """تبدیل متن به WAV با SAPI5."""
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    for voice in speaker.GetVoices():
        if voice_name.lower() in voice.GetDescription().lower():
            speaker.Voice = voice
            break
    stream = win32com.client.Dispatch("SAPI.SpFileStream")
    stream.Open(wav_path, 3, False)
    speaker.AudioOutputStream = stream
    speaker.Speak(text)
    stream.Close()


# ------------------ رابط گرافیکی ------------------

class PDFTTSApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF TO VOICE")
        self.geometry("800x500")
        self.pdf_path = tk.StringVar()
        self.voice_var = tk.StringVar()
        self.lang_var = tk.StringVar(value="auto")
        self.create_widgets()

    def create_widgets(self):
        # انتخاب PDF
        frm_file = tk.LabelFrame(self, text="CHOISE PDF")
        frm_file.pack(fill="x", padx=10, pady=10)
        tk.Entry(frm_file, textvariable=self.pdf_path, width=60).pack(side="left", padx=5, pady=5)
        tk.Button(frm_file, text="📂 CHOISE", command=self.choose_pdf).pack(side="left", padx=5)
        # انتخاب صدا
        frm_voice = tk.LabelFrame(self, text="CHOISE VOICE")
        frm_voice.pack(fill="x", padx=10, pady=5)
        self.voice_menu = ttk.Combobox(frm_voice, textvariable=self.voice_var, state="readonly")
        self.voice_menu.pack(padx=5, pady=5)

        # دکمه شروع
        
        tk.Button(self, text="🎤 GENERATE TO WAV", bg="green", fg="white", command=self.convert_pdf).pack(pady=15)
    

    def choose_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if path:
            self.pdf_path.set(path)
            text = extract_text_from_pdf(path)
            lang = detect_language(text) if self.lang_var.get() == "auto" else self.lang_var.get()
            self.lang_var.set(lang)
            self.load_voices(lang)

    def load_voices(self, lang):
        all_voices = list_voices()
        
        filtered = all_voices
        self.voice_menu["values"] = filtered
        if filtered:
            self.voice_menu.current(0)

    def convert_pdf(self):
        if not self.pdf_path.get():
            messagebox.showerror("PLEASE CHOOSE PDF!!!")
            return
        save_path = filedialog.asksaveasfilename(defaultextension=".wav", filetypes=[("WAV Files", "*.wav")])
        if not save_path:
            return
        text = extract_text_from_pdf(self.pdf_path.get())
        speak_to_wav(text, self.voice_var.get(), save_path)
        messagebox.showinfo("SUCCESS", f"FILE SAVED IN :\n{save_path}")

    
# ------------------ اجرای برنامه ------------------

if __name__ == "__main__":
    app = PDFTTSApp()
    app.mainloop()
