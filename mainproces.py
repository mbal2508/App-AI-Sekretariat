import whisper
from pathlib import Path
from tkinter import *
from tkinter import filedialog, messagebox
import os
import sys
import sqlite3
from datetime import datetime
import ctypes 
import threading
import google.generativeai as genai
import random 
from dotenv import load_dotenv

# Pastikan module untuk export Word tersedia
try:
    import docx
    from htmldocx import HtmlToDocx
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

# ==================== GEMINI API CONFIGURATION ====================
load_dotenv()
GEMINI_API_KEY = os.getenv("API_KEY")
if GEMINI_API_KEY and GEMINI_API_KEY != "YOUR_GEMINI_API_KEY":
    genai.configure(api_key=GEMINI_API_KEY)

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# ==================== MAIN UI SETTINGS ====================
window = Tk()

BG_APP = "#003246"
BG_SIDEBAR = "#001f2c" 
BG_FRAME = "#A1FAD0"
FG_TEXT = "#aec5d4"
BTN_BLUE = "#89b4fa"
BTN_GREEN = "#a6e3a1"
BTN_RED = "#f38ba8"
FONT_TITLE = ("Segoe UI", 16, "bold")
FONT_SUBTITLE = ("Segoe UI", 12, "bold")
FONT_NORMAL = ("Segoe UI", 11)

window.title("NoteTaker")
window.geometry('1000x720') 
window.config(background=BG_APP)

def set_titlebar_color(win, hex_color):
    try:
        win.update()
        r = int(hex_color[1:3], 16)
        g = int(hex_color[3:5], 16)
        b = int(hex_color[5:7], 16)
        bgr_color = (b << 16) | (g << 8) | r
        hwnd = ctypes.windll.user32.GetParent(win.winfo_id())
        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 35, ctypes.byref(ctypes.c_int(bgr_color)), 4)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 36, ctypes.byref(ctypes.c_int(0xFFFFFF)), 4)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(ctypes.c_int(1)), 4)
    except Exception:
        pass 
        
set_titlebar_color(window, BG_APP)

try:
    logo = PhotoImage(file=resource_path("NoteTaker_logo.png"))
    logo = logo.subsample(2, 2) 
    window.iconphoto(True, logo)
    has_logo = True
except Exception as e:
    has_logo = False

# ==================== GLOBAL VARIABLE ====================
path_file = ""
mapping_id_rapat = []  
mode_history = "notulen" 

docmaker_current_name = ""
docmaker_current_path = ""

# ==================== DATABASE FUNCTIONS ====================
def inisialisasi_db():
    conn = sqlite3.connect("database_sekretariat.db")
    cursor = conn.cursor()
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS notulen (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            judul TEXT NOT NULL,
            hasil_rapat TEXT NOT NULL,
            timestamp TEXT NOT NULL
        )
    """)
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS daily_notes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            judul TEXT NOT NULL,
            catatan TEXT NOT NULL,
            timestamp TEXT NOT NULL
        )
    """)
    
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS docmaker_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nama_file TEXT NOT NULL,
            alamat_file TEXT NOT NULL,
            timestamp TEXT NOT NULL
        )
    """)
    
    conn.commit()
    conn.close()

inisialisasi_db()

def simpan_rapat_ke_db(judul, hasil_rapat):
    if not hasil_rapat.strip() or hasil_rapat == "Your summarize will pop here":
        messagebox.showwarning("Peringatan", "Hasil Rapat tidak boleh kosong!")
        return False

    if not judul.strip():
        tanggal_hari_ini = datetime.now().strftime("%d%m%Y")
        angka_unik = random.randint(100, 999)
        judul = f"notulen_{tanggal_hari_ini}_{angka_unik}"

    try:
        conn = sqlite3.connect("database_sekretariat.db")
        cursor = conn.cursor()
        waktu_sekarang = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        query = "INSERT INTO notulen (judul, hasil_rapat, timestamp) VALUES (?, ?, ?)"
        cursor.execute(query, (judul, hasil_rapat, waktu_sekarang))
        conn.commit()
        messagebox.showinfo("Sukses", f"Data rapat berhasil disimpan ke database!\nJudul: {judul}")
        return True
    except sqlite3.Error as e:
        messagebox.showerror("Error Database", f"Gagal menyimpan data: {e}")
        return False
    finally:
        conn.close()

# ==================== SIDEBAR NAVIGATION SYSTEM ====================
def buka_halaman_summarizer():
    page_history.pack_forget()
    page_daily_note.pack_forget()
    page_docmaker.pack_forget()
    page_summarizer.pack(fill="both", expand=True)
    btn_nav_sum.config(bg=BTN_BLUE, fg="#11111b")
    btn_nav_note.config(bg=BG_SIDEBAR, fg=FG_TEXT)
    btn_nav_doc.config(bg=BG_SIDEBAR, fg=FG_TEXT)
    btn_nav_hist.config(bg=BG_SIDEBAR, fg=FG_TEXT)

def buka_halaman_daily_note():
    page_summarizer.pack_forget()
    page_history.pack_forget()
    page_docmaker.pack_forget()
    page_daily_note.pack(fill="both", expand=True)
    btn_nav_sum.config(bg=BG_SIDEBAR, fg=FG_TEXT)
    btn_nav_note.config(bg=BTN_BLUE, fg="#11111b")
    btn_nav_doc.config(bg=BG_SIDEBAR, fg=FG_TEXT)
    btn_nav_hist.config(bg=BG_SIDEBAR, fg=FG_TEXT)

def buka_halaman_docmaker():
    page_summarizer.pack_forget()
    page_daily_note.pack_forget()
    page_history.pack_forget()
    page_docmaker.pack(fill="both", expand=True)
    btn_nav_sum.config(bg=BG_SIDEBAR, fg=FG_TEXT)
    btn_nav_note.config(bg=BG_SIDEBAR, fg=FG_TEXT)
    btn_nav_doc.config(bg=BTN_BLUE, fg="#11111b")
    btn_nav_hist.config(bg=BG_SIDEBAR, fg=FG_TEXT)

def buka_halaman_history():
    page_summarizer.pack_forget()
    page_daily_note.pack_forget()
    page_docmaker.pack_forget()
    page_history.pack(fill="both", expand=True)
    btn_nav_sum.config(bg=BG_SIDEBAR, fg=FG_TEXT)
    btn_nav_note.config(bg=BG_SIDEBAR, fg=FG_TEXT)
    btn_nav_doc.config(bg=BG_SIDEBAR, fg=FG_TEXT)
    btn_nav_hist.config(bg=BTN_BLUE, fg="#11111b")
    memuat_data_history() 

sidebar_frame = Frame(window, bg=BG_SIDEBAR, width=200)
sidebar_frame.pack(side="left", fill="y")
sidebar_frame.pack_propagate(False)

main_content = Frame(window, bg=BG_APP)
main_content.pack(side="right", fill="both", expand=True)

if has_logo:
    lbl_brand = Label(sidebar_frame, image=logo, bg=BG_SIDEBAR, pady=20)
else:
    lbl_brand = Label(sidebar_frame, text="NoteTaker", font=FONT_TITLE, fg=BTN_BLUE, bg=BG_SIDEBAR, pady=20)
lbl_brand.pack()

btn_nav_sum = Button(sidebar_frame, text="AI Summarizer", font=FONT_NORMAL, bg=BG_SIDEBAR, fg=FG_TEXT, relief="flat", activebackground=BTN_BLUE, cursor="hand2", pady=10, command=buka_halaman_summarizer)
btn_nav_sum.pack(fill="x", padx=10, pady=5)

btn_nav_note = Button(sidebar_frame, text="Daily Note", font=FONT_NORMAL, bg=BG_SIDEBAR, fg=FG_TEXT, relief="flat", activebackground=BTN_BLUE, cursor="hand2", pady=10, command=buka_halaman_daily_note)
btn_nav_note.pack(fill="x", padx=10, pady=5)

btn_nav_doc = Button(sidebar_frame, text="DocMaker", font=FONT_NORMAL, bg=BG_SIDEBAR, fg=FG_TEXT, relief="flat", activebackground=BTN_BLUE, cursor="hand2", pady=10, command=buka_halaman_docmaker)
btn_nav_doc.pack(fill="x", padx=10, pady=5)

btn_nav_hist = Button(sidebar_frame, text="Saved History", font=FONT_NORMAL, bg=BG_SIDEBAR, fg=FG_TEXT, relief="flat", activebackground=BTN_BLUE, cursor="hand2", pady=10, command=buka_halaman_history)
btn_nav_hist.pack(fill="x", padx=10, pady=5)

# ==================== TAB 1: PAGE AI SUMMARIZER ====================
page_summarizer = Frame(main_content, bg=BG_APP)
frame_header = Frame(page_summarizer, bg=BG_APP)
frame_header.pack(fill="x", pady=15)

if has_logo:
    label_title = Label(frame_header, font=FONT_TITLE, fg=FG_TEXT, bg=BG_APP, image=logo)
else:
    label_title = Label(frame_header, text="AI Audio Summarizer", font=FONT_TITLE, fg=FG_TEXT, bg=BG_APP)
label_title.pack()

frame_preset = LabelFrame(page_summarizer, text=" AI SUMMARIZER ", font=FONT_SUBTITLE, bg=BG_APP, fg=BTN_BLUE, bd=1, padx=15, pady=15)
frame_preset.pack(fill="x", padx=30, pady=10)

def input_file():
    global path_file 
    lokasi_file = filedialog.askopenfilename(
        title="Pilih File Kamu",
        filetypes=[("Audio Files", "*.mp3 *.m4a *.wav"), ("Semua File", "*.*")]
    )
    if lokasi_file:
        path_file = lokasi_file
        file_name = os.path.basename(lokasi_file)
        lbl_info.config(text=f"{file_name}", font=FONT_SUBTITLE, bg=BG_APP, fg=BTN_GREEN, bd=1, padx=15, pady=15)
        btn_upload_file.pack_forget()
        btn_upload_file.config(text="Change File", font=FONT_NORMAL, bg=BTN_BLUE, fg="#11111b", relief="flat", cursor="hand2", command=input_file)
        btn_upload_file.pack(side="right", padx=30, pady=10)
        btn_summirizing.pack(side="left", padx=30, pady=10)
    else:
        lbl_info.config(text="Canceled by user.", font=FONT_SUBTITLE, bg=BG_APP, fg=BTN_RED, bd=1, padx=15, pady=15)

def proses_background():
    global path_file
    try:
        model_gemini = genai.GenerativeModel("gemini-2.5-flash")
        
        file_asli = Path(path_file)
        file_baru = file_asli.with_suffix(".mp3")
        if file_asli.exists() and file_asli != file_baru:
            try:
                file_asli.rename(file_baru)
            except:
                pass

        txt_result.config(state="normal")
        txt_result.delete("1.0", END)
        txt_result.insert(END, "[1/2] Sedang mentranskrip audio dengan Whisper AI...\nMohon tunggu...\n")
        txt_result.config(state="disabled")

        model_whisper = whisper.load_model("base")
        hasil_transkrip = model_whisper.transcribe(str(file_baru))
        teks_transkrip = hasil_transkrip["text"]

        txt_result.config(state="normal")
        txt_result.delete("1.0", END)
        txt_result.insert(END, "[2/2] Transkripsi selesai. Mengirim ke Gemini AI untuk diringkas...\n")
        txt_result.config(state="disabled")

        prompt_peringkas = (
            "Berikut adalah transkrip mentah dari rekaman rapat/pertemuan:\n\n"
            f"\"\"\"{teks_transkrip}\"\"\"\n\n"
            "Tugas Anda: Buatlah ringkasan notulen rapat yang rapi, ringkas, dan profesional. "
            "Format notulen harus mencakup:\n"
            "1. **Ringkasan Singkat**\n"
            "2. **Poin-Poin Utama Pembahasan**\n"
            "3. **Keputusan yang Diambil**\n"
            "4. **Tindak Lanjut / Action Items**\n\n"
            "Gunakan bahasa Indonesia yang baik dan lugas."
        )

        response = model_gemini.generate_content(prompt_peringkas)
        ringkasan_gemini = response.text

        txt_result.config(state="normal")
        txt_result.delete("1.0", END)
        txt_result.insert(END, ringkasan_gemini)
        txt_result.config(state="disabled") 

        btn_save.pack(fill="x", pady=(5, 0))

    except Exception as e:
        txt_result.config(state="normal")
        txt_result.delete("1.0", END)
        error_str = str(e)
        if "429" in error_str or "ResourceExhausted" in error_str or "quota" in error_str.lower():
            txt_result.insert(END, "❌ Gagal: Kuota atau Limit Token Gemini API Anda sudah habis (Rate Limit / Quota Exceeded).\nSilakan coba lagi nanti atau gunakan API key lain.")
        else:
            txt_result.insert(END, f"Terjadi Kesalahan:\n{error_str}")
        txt_result.config(state="disabled")
    
    btn_summirizing.config(state="normal", bg=BTN_GREEN)

def summarizing():
    global path_file 
    if not path_file:
        return
    btn_summirizing.config(state="disabled", bg="#555555")
    t = threading.Thread(target=proses_background)
    t.daemon = True
    t.start()

lbl_info = Label(frame_preset, text="No files uploaded yet", font=FONT_SUBTITLE, bg=BG_APP, fg=FG_TEXT, bd=1, padx=15, pady=15)
lbl_info.pack(fill="x", padx=30, pady=10)

btn_upload_file = Button(frame_preset, text="Upload your file here", font=FONT_NORMAL, bg=BTN_BLUE, fg="#11111b", relief="flat", cursor="hand2", command=input_file)
btn_upload_file.pack(expand=True, fill="x", padx=5)

btn_summirizing = Button(frame_preset, text="Summarize", font=FONT_NORMAL, bg=BTN_GREEN, fg="#11111b", relief="flat", cursor="hand2", command=summarizing)

frame_result = LabelFrame(page_summarizer, font=FONT_SUBTITLE, bg=BTN_GREEN, bd=1, padx=15, pady=15)
frame_result.pack(fill="both", padx=30, pady=20, expand=True)

frame_tombol_bawah = Frame(frame_result, bg=BTN_GREEN)
frame_tombol_bawah.pack(side="bottom", fill="x")

frame_text_box = Frame(frame_result, bg=BTN_GREEN)
frame_text_box.pack(side="top", fill="both", expand=True)

scroll_y = Scrollbar(frame_text_box, orient="vertical")
scroll_y.pack(side="right", fill="y")

txt_result = Text(frame_text_box, font=FONT_NORMAL, bg=BTN_GREEN, fg="#11111b", bd=0, wrap="word", yscrollcommand=scroll_y.set)
txt_result.pack(side="left", fill="both", expand=True)
scroll_y.config(command=txt_result.yview)

txt_result.insert(END, "Your summarize will pop here")
txt_result.config(state="disabled") 

def reset_halaman_summarizer():
    global path_file
    path_file = ""
    lbl_info.config(text="No files uploaded yet", fg=FG_TEXT)
    btn_summirizing.pack_forget()
    btn_upload_file.pack_forget()
    btn_upload_file.config(text="Upload your file here", bg=BTN_BLUE, fg="#11111b", command=input_file)
    btn_upload_file.pack(expand=True, fill="x", padx=5)
    btn_save.pack_forget()
    txt_result.config(state="normal")
    txt_result.delete("1.0", END)
    txt_result.insert(END, "Your summarize will pop here")
    txt_result.config(state="disabled")

# --- POP-UP OVERLAY UNTUK AI SUMMARIZER ---
overlay_sum_frame = Frame(page_summarizer, bg=BG_SIDEBAR, bd=2, relief="ridge", highlightbackground=BTN_BLUE, highlightthickness=2)

lbl_overlay_sum_title = Label(overlay_sum_frame, text="Meet Title:", font=FONT_SUBTITLE, bg=BG_SIDEBAR, fg=FG_TEXT)
lbl_overlay_sum_title.pack(pady=(20, 10), padx=30)

ent_judul_sum = Entry(overlay_sum_frame, font=FONT_NORMAL, bg=BG_APP, fg=FG_TEXT, bd=1, relief="flat", insertbackground=FG_TEXT, width=35)
ent_judul_sum.pack(pady=10, padx=30, ipady=5)

frame_overlay_sum_btns = Frame(overlay_sum_frame, bg=BG_SIDEBAR)
frame_overlay_sum_btns.pack(pady=(15, 20), padx=30, fill="x")

def sembunyikan_overlay_sum():
    overlay_sum_frame.place_forget()

def eksekusi_simpan_sum():
    judul_input = ent_judul_sum.get().strip()
    hasil_input = txt_result.get("1.0", "end-1c").strip()
    berhasil = simpan_rapat_ke_db(judul_input, hasil_input)
    if berhasil:
        ent_judul_sum.delete(0, END)
        sembunyikan_overlay_sum()
        reset_halaman_summarizer()

def tampilkan_overlay_sum():
    hasil_input = txt_result.get("1.0", "end-1c").strip()
    if not hasil_input or hasil_input == "Your summarize will pop here":
        messagebox.showwarning("Peringatan", "Belum ada hasil ringkasan yang bisa disimpan!")
        return
    overlay_sum_frame.place(relx=0.5, rely=0.5, anchor="center")
    overlay_sum_frame.lift()
    ent_judul_sum.focus()

btn_save_sum_anyway = Button(frame_overlay_sum_btns, text="Save Anyway", font=FONT_NORMAL, bg=BTN_GREEN, fg="#11111b", cursor="hand2", relief="flat", command=eksekusi_simpan_sum)
btn_save_sum_anyway.pack(side="left", expand=True, fill="x", padx=(0, 5))

btn_cancel_sum = Button(frame_overlay_sum_btns, text="Cancel", font=FONT_NORMAL, bg=BTN_RED, fg="#11111b", cursor="hand2", relief="flat", command=sembunyikan_overlay_sum)
btn_cancel_sum.pack(side="right", expand=True, fill="x", padx=(5, 0))

btn_save = Button(frame_tombol_bawah, text="SAVE", font=FONT_NORMAL, bg=BTN_BLUE, fg="#11111b", relief="flat", cursor="hand2", command=tampilkan_overlay_sum)

# ==================== TAB 2: PAGE DAILY NOTE ====================
page_daily_note = Frame(main_content, bg=BG_APP)

canvas_note = Canvas(page_daily_note, bg=BG_APP, highlightthickness=0)
scroll_daily_y = Scrollbar(page_daily_note, orient="vertical", command=canvas_note.yview, bg=BG_APP, troughcolor=BG_APP)
scrollable_daily_note = Frame(canvas_note, bg=BG_APP)

scrollable_daily_note.bind(
    "<Configure>",
    lambda e: canvas_note.configure(scrollregion=canvas_note.bbox("all"))
)

canvas_window_note = canvas_note.create_window((0, 0), window=scrollable_daily_note, anchor="nw")

def _on_canvas_configure_note(event):
    canvas_note.itemconfig(canvas_window_note, width=event.width)

canvas_note.bind("<Configure>", _on_canvas_configure_note)
canvas_note.configure(yscrollcommand=scroll_daily_y.set)

scroll_daily_y.pack(side="right", fill="y")
canvas_note.pack(side="left", fill="both", expand=True)

if has_logo:
    label_title_note = Label(scrollable_daily_note, font=FONT_TITLE, fg=FG_TEXT, bg=BG_APP, image=logo)
else:
    label_title_note = Label(scrollable_daily_note, text="Daily Note", font=FONT_TITLE, fg=FG_TEXT, bg=BG_APP)
label_title_note.pack(pady=(20, 0))

frame_text_note = LabelFrame(scrollable_daily_note, bg=BG_APP, text=" Type Here ", font=FONT_SUBTITLE, fg=BTN_BLUE, bd=2, labelanchor="nw")
frame_text_note.pack(fill="x", padx=40, pady=5)

scroll_note_y = Scrollbar(frame_text_note, orient="vertical", bg=BG_APP, troughcolor=BG_APP)
scroll_note_y.pack(side="right", fill="y")

txt_note_content = Text(frame_text_note, font=FONT_NORMAL, bg=BG_SIDEBAR, fg=FG_TEXT, bd=2, relief="groove", wrap="word", insertbackground="white", yscrollcommand=scroll_note_y.set, height=6)
txt_note_content.pack(fill="both", expand=True, padx=10, pady=10)
scroll_note_y.config(command=txt_note_content.yview)

lbl_or = Label(scrollable_daily_note, text="OR", font=("Segoe UI", 16, "bold"), fg=FG_TEXT, bg=BG_APP)
lbl_or.pack(pady=5)

frame_upload_note = LabelFrame(scrollable_daily_note, text=" Upload ", font=FONT_SUBTITLE, bg=BG_APP, fg=BTN_BLUE, bd=2, labelanchor="nw")
frame_upload_note.pack(fill="x", padx=40, pady=5)

lbl_selected_file = Label(frame_upload_note, text='"No File Selected"', font=FONT_NORMAL, bg=BG_APP, fg=FG_TEXT)
lbl_selected_file.pack(pady=(15, 5))

frame_transkrip = Frame(frame_upload_note, bg=BG_APP)
frame_transkrip.pack(fill="both", expand=True, padx=15, pady=(0, 15))

scroll_transkrip_y = Scrollbar(frame_transkrip, orient="vertical", bg=BG_APP, troughcolor=BG_APP)
scroll_transkrip_y.pack(side="right", fill="y")

txt_upload_result = Text(frame_transkrip, font=FONT_NORMAL, bg=BTN_GREEN, fg=FG_TEXT, bd=2, relief="groove", wrap="word", insertbackground="white", yscrollcommand=scroll_transkrip_y.set, height=4)
txt_upload_result.pack(fill="both", expand=True)
scroll_transkrip_y.config(command=txt_upload_result.yview)
txt_upload_result.config(state="disabled")

def proses_transkrip_note(filepath):
    btn_select_file.config(state="disabled")
    lbl_selected_file.config(text='"Sedang Transkrip..."', fg=BTN_BLUE)
    try:
        model_whisper = whisper.load_model("base")
        hasil_transkrip = model_whisper.transcribe(filepath)
        teks_transkrip = hasil_transkrip["text"]
        
        txt_upload_result.config(state="normal")
        txt_upload_result.delete("1.0", END)
        txt_upload_result.insert(END, teks_transkrip)
        txt_upload_result.config(state="disabled")
        
        lbl_selected_file.config(text='"Transkrip Selesai!"', fg=BTN_GREEN)
    except Exception as e:
        messagebox.showerror("Error Transkripsi", f"Gagal mentranskrip audio: {e}")
        lbl_selected_file.config(text='"Error Transkripsi"', fg=BTN_RED)
    finally:
        btn_select_file.config(state="normal")

def input_audio_note():
    lokasi_file = filedialog.askopenfilename(
        title="Pilih File Audio untuk Catatan",
        filetypes=[("Audio Files", "*.mp3 *.m4a *.wav"), ("Semua File", "*.*")]
    )
    if lokasi_file:
        nama_file = os.path.basename(lokasi_file)
        lbl_selected_file.config(text=f'"{nama_file}"', fg=BTN_GREEN)
        t = threading.Thread(target=proses_transkrip_note, args=(lokasi_file,))
        t.daemon = True
        t.start()

btn_select_file = Button(frame_upload_note, text="Select File", font=FONT_NORMAL, bg=BTN_BLUE, fg="#11111b", relief="flat", cursor="hand2", command=input_audio_note, width=15)
btn_select_file.pack(pady=(5, 15), before=frame_transkrip)

def tampilkan_overlay_simpan():
    catatan_manual = txt_note_content.get("1.0", "end-1c").strip()
    catatan_upload = txt_upload_result.get("1.0", "end-1c").strip()
    
    if not catatan_manual and not catatan_upload:
        messagebox.showwarning("Peringatan", "Catatan masih kosong! Ketik sesuatu atau upload file terlebih dahulu.")
        return
    overlay_frame.place(relx=0.5, rely=0.5, anchor="center")
    overlay_frame.lift()
    ent_judul_note.focus()

btn_save_trigger = Button(scrollable_daily_note, text="Save", font=("Segoe UI", 16, "bold"), bg=BTN_BLUE, fg="#11111b", relief="flat", cursor="hand2", command=tampilkan_overlay_simpan)
btn_save_trigger.pack(fill="x", padx=40, pady=(15, 30))

overlay_frame = Frame(page_daily_note, bg=BG_SIDEBAR, bd=2, relief="ridge", highlightbackground=BTN_BLUE, highlightthickness=2)

lbl_overlay_title = Label(overlay_frame, text="Beri Nama Catatan", font=FONT_SUBTITLE, bg=BG_SIDEBAR, fg=FG_TEXT)
lbl_overlay_title.pack(pady=(20, 10), padx=30)

ent_judul_note = Entry(overlay_frame, font=FONT_NORMAL, bg=BG_APP, fg=FG_TEXT, bd=1, relief="flat", insertbackground=FG_TEXT, width=35)
ent_judul_note.pack(pady=10, padx=30, ipady=5)

frame_overlay_btns = Frame(overlay_frame, bg=BG_SIDEBAR)
frame_overlay_btns.pack(pady=(15, 20), padx=30, fill="x")

def sembunyikan_overlay_simpan():
    overlay_frame.place_forget()

def eksekusi_simpan_note():
    judul = ent_judul_note.get().strip()
    catatan_manual = txt_note_content.get("1.0", "end-1c").strip()
    catatan_upload = txt_upload_result.get("1.0", "end-1c").strip()
    
    catatan = ""
    if catatan_manual and catatan_upload:
        catatan = f"{catatan_manual}\n\n--- Hasil Transkrip Audio ---\n{catatan_upload}"
    elif catatan_upload:
        catatan = catatan_upload
    else:
        catatan = catatan_manual
    
    if not judul:
        tanggal_hari_ini = datetime.now().strftime("%d%m%Y")
        angka_unik = random.randint(100, 999)
        judul = f"note_{tanggal_hari_ini}_{angka_unik}"
        
    try:
        conn = sqlite3.connect("database_sekretariat.db")
        cursor = conn.cursor()
        waktu_sekarang = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cursor.execute("INSERT INTO daily_notes (judul, catatan, timestamp) VALUES (?, ?, ?)", (judul, catatan, waktu_sekarang))
        conn.commit()
        conn.close()
        
        messagebox.showinfo("Sukses", f"Catatan berhasil disimpan ke Database!\nJudul: {judul}")
        ent_judul_note.delete(0, END)
        txt_note_content.delete("1.0", END)
        
        txt_upload_result.config(state="normal")
        txt_upload_result.delete("1.0", END)
        txt_upload_result.config(state="disabled")
        
        lbl_selected_file.config(text='"No File Selected"', fg=FG_TEXT)
        sembunyikan_overlay_simpan()
    except sqlite3.Error as e:
        messagebox.showerror("Error DB", f"Gagal menyimpan catatan: {e}")

btn_save_anyway = Button(frame_overlay_btns, text="Save Anyway", font=FONT_NORMAL, bg=BTN_GREEN, fg="#11111b", cursor="hand2", relief="flat", command=eksekusi_simpan_note)
btn_save_anyway.pack(side="left", expand=True, fill="x", padx=(0, 5))

btn_cancel = Button(frame_overlay_btns, text="Cancel", font=FONT_NORMAL, bg=BTN_RED, fg="#11111b", cursor="hand2", relief="flat", command=sembunyikan_overlay_simpan)
btn_cancel.pack(side="right", expand=True, fill="x", padx=(5, 0))


# ==================== TAB 3: PAGE DOCMAKER (NEW FEATURE) ====================
page_docmaker = Frame(main_content, bg=BG_APP)

canvas_doc = Canvas(page_docmaker, bg=BG_APP, highlightthickness=0)
scroll_doc_y = Scrollbar(page_docmaker, orient="vertical", command=canvas_doc.yview, bg=BG_APP, troughcolor=BG_APP)
scrollable_doc = Frame(canvas_doc, bg=BG_APP)

scrollable_doc.bind(
    "<Configure>",
    lambda e: canvas_doc.configure(scrollregion=canvas_doc.bbox("all"))
)

canvas_window_doc = canvas_doc.create_window((0, 0), window=scrollable_doc, anchor="nw")

def _on_canvas_configure_doc(event):
    canvas_doc.itemconfig(canvas_window_doc, width=event.width)

canvas_doc.bind("<Configure>", _on_canvas_configure_doc)
canvas_doc.configure(yscrollcommand=scroll_doc_y.set)

scroll_doc_y.pack(side="right", fill="y")
canvas_doc.pack(side="left", fill="both", expand=True)

if has_logo:
    label_title_doc = Label(scrollable_doc, font=FONT_TITLE, fg=FG_TEXT, bg=BG_APP, image=logo)
else:
    label_title_doc = Label(scrollable_doc, text="DocMaker", font=FONT_TITLE, fg=FG_TEXT, bg=BG_APP)
label_title_doc.pack(pady=(20, 10))

# Bagian "Describe Your Doc" (sesuai referensi image_1e3b6f.png)
frame_desc_doc = LabelFrame(scrollable_doc, bg=BG_APP, text=" Describe Your Doc ", font=FONT_SUBTITLE, fg=BTN_BLUE, bd=2, labelanchor="nw")
frame_desc_doc.pack(fill="x", padx=40, pady=5)

scroll_desc_y = Scrollbar(frame_desc_doc, orient="vertical", bg=BG_APP, troughcolor=BG_APP)
scroll_desc_y.pack(side="right", fill="y")

txt_doc_desc = Text(frame_desc_doc, font=FONT_NORMAL, bg=BG_SIDEBAR, fg=FG_TEXT, bd=2, relief="groove", wrap="word", insertbackground="white", yscrollcommand=scroll_desc_y.set, height=6)
txt_doc_desc.pack(fill="both", expand=True, padx=10, pady=10)
scroll_desc_y.config(command=txt_doc_desc.yview)

def proses_generate_doc():
    global docmaker_current_name, docmaker_current_path
    
    deskripsi = txt_doc_desc.get("1.0", END).strip()
    if not deskripsi:
        messagebox.showwarning("Peringatan", "Mohon deskripsikan isi surat yang ingin dibuat!")
        btn_generate_doc.config(state="normal", text="Generate Document")
        return
        
    if not DOCX_AVAILABLE:
        messagebox.showerror("Error Module", "Library python-docx atau htmldocx tidak ditemukan.\nSilakan jalankan: pip install python-docx htmldocx")
        btn_generate_doc.config(state="normal", text="Generate Document")
        return

    try:
        model_gemini = genai.GenerativeModel("gemini-2.5-flash")
        
        prompt_doc = (
            "Buat dokumen/surat berdasarkan deskripsi ini:\n"
            f"\"{deskripsi}\"\n\n"
            "PENTING: Kembalikan HANYA teks dalam format HTML body (jangan sertakan tag <html>, <head>, atau <body>). "
            "Gunakan tag dasar seperti <h1>, <h2>, <p>, <b>, <i>, <br> saja. Jangan gunakan format markdown ```html"
        )
        
        response = model_gemini.generate_content(prompt_doc)
        html_content = response.text.replace("```html", "").replace("```", "").strip()

        # Konversi ke DOCX
        document = docx.Document()
        new_parser = HtmlToDocx()
        new_parser.add_html_to_document(html_content, document)
        
        # Buat folder jika belum ada
        dir_simpan = os.path.join(os.getcwd(), "GeneratedDocs")
        if not os.path.exists(dir_simpan):
            os.makedirs(dir_simpan)
            
        timestamp_str = datetime.now().strftime("%d%m%Y_%H%M%S")
        nama_file = f"Surat_Generated_{timestamp_str}.docx"
        alamat_file = os.path.join(dir_simpan, nama_file)
        
        document.save(alamat_file)
        
        # Set variable global untuk db
        docmaker_current_name = nama_file
        docmaker_current_path = alamat_file
        
        # Update UI Result
        txt_doc_result.config(state="normal")
        txt_doc_result.delete("1.0", END)
        txt_doc_result.insert(END, nama_file)
        txt_doc_result.config(state="disabled")
        
        btn_save_doc.config(state="normal")
        
        messagebox.showinfo("Berhasil", f"Surat selesai dibuat!\nNama: {nama_file}\nSilakan cek di folder GeneratedDocs.")
        
    except Exception as e:
        messagebox.showerror("Error", f"Terjadi kesalahan saat memproses surat:\n{e}")
    finally:
        btn_generate_doc.config(state="normal", text="Generate Document")

def mulai_generate_doc():
    btn_generate_doc.config(state="disabled", text="Generating...")
    btn_save_doc.config(state="disabled")
    
    txt_doc_result.config(state="normal")
    txt_doc_result.delete("1.0", END)
    txt_doc_result.config(state="disabled")
    
    t = threading.Thread(target=proses_generate_doc)
    t.daemon = True
    t.start()

btn_generate_doc = Button(scrollable_doc, text="Generate Document", font=FONT_NORMAL, bg=BTN_BLUE, fg="#11111b", relief="flat", cursor="hand2", command=mulai_generate_doc)
btn_generate_doc.pack(pady=10)

# Bagian "Result" (sesuai referensi image_1e3b6f.png)
frame_result_doc = LabelFrame(scrollable_doc, text=" Result ", font=FONT_SUBTITLE, bg=BG_APP, fg=BTN_BLUE, bd=2, labelanchor="nw")
frame_result_doc.pack(fill="x", padx=40, pady=15)

txt_doc_result = Text(frame_result_doc, font=FONT_NORMAL, bg=BTN_GREEN, fg="#11111b", bd=0, height=2)
txt_doc_result.pack(fill="both", expand=True, padx=10, pady=20)
txt_doc_result.insert(END, "File_name_result will appear here")
txt_doc_result.tag_configure("center", justify='center')
txt_doc_result.tag_add("center", "1.0", "end")
txt_doc_result.config(state="disabled")

def simpan_doc_ke_db():
    global docmaker_current_name, docmaker_current_path
    if not docmaker_current_name or not docmaker_current_path:
        return
        
    try:
        conn = sqlite3.connect("database_sekretariat.db")
        cursor = conn.cursor()
        waktu_sekarang = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cursor.execute("INSERT INTO docmaker_history (nama_file, alamat_file, timestamp) VALUES (?, ?, ?)", 
                       (docmaker_current_name, docmaker_current_path, waktu_sekarang))
        conn.commit()
        conn.close()
        
        messagebox.showinfo("Sukses", "Data lokasi surat berhasil disimpan ke History!")
        txt_doc_desc.delete("1.0", END)
        txt_doc_result.config(state="normal")
        txt_doc_result.delete("1.0", END)
        txt_doc_result.insert(END, "File_name_result will appear here")
        txt_doc_result.tag_add("center", "1.0", "end")
        txt_doc_result.config(state="disabled")
        btn_save_doc.config(state="disabled")
        
        docmaker_current_name = ""
        docmaker_current_path = ""
    except sqlite3.Error as e:
        messagebox.showerror("Error DB", f"Gagal menyimpan data: {e}")

# Bagian "SAVE" (sesuai referensi image_1e3b6f.png)
btn_save_doc = Button(scrollable_doc, text="S A V E", font=("Segoe UI", 14, "bold"), bg=BG_APP, fg=BTN_GREEN, bd=2, relief="groove", activebackground=BTN_GREEN, activeforeground="#11111b", cursor="hand2", command=simpan_doc_ke_db)
btn_save_doc.pack(pady=(5, 30), ipadx=30, ipady=5)
btn_save_doc.config(state="disabled")


# ==================== TAB 4: PAGE SAVED HISTORY ====================
page_history = Frame(main_content, bg=BG_APP)

frame_header_hist = Frame(page_history, bg=BG_APP)
frame_header_hist.pack(fill="x", pady=10)
lbl_title_hist = Label(frame_header_hist, text="Saved Records Database", font=FONT_TITLE, fg=FG_TEXT, bg=BG_APP)
lbl_title_hist.pack()

frame_hist_tabs = Frame(page_history, bg=BG_APP)
frame_hist_tabs.pack(fill="x", padx=30, pady=(0, 10))

def ganti_mode_history(mode_baru):
    global mode_history
    mode_history = mode_baru
    
    if mode_history == "notulen":
        btn_tab_notulen.config(bg=BTN_BLUE, fg="#11111b")
        btn_tab_note.config(bg=BG_SIDEBAR, fg=FG_TEXT)
        btn_tab_docmaker.config(bg=BG_SIDEBAR, fg=FG_TEXT)
        frame_kiri_list.config(text=" Daftar Notulen Rapat ")
        frame_kanan_detail.config(text=" Isi Dokumen Rapat ")
    elif mode_history == "note":
        btn_tab_notulen.config(bg=BG_SIDEBAR, fg=FG_TEXT)
        btn_tab_note.config(bg=BTN_BLUE, fg="#11111b")
        btn_tab_docmaker.config(bg=BG_SIDEBAR, fg=FG_TEXT)
        frame_kiri_list.config(text=" Daftar Catatan Pribadi ")
        frame_kanan_detail.config(text=" Isi Catatan Pribadi ")
    else: # mode docmaker
        btn_tab_notulen.config(bg=BG_SIDEBAR, fg=FG_TEXT)
        btn_tab_note.config(bg=BG_SIDEBAR, fg=FG_TEXT)
        btn_tab_docmaker.config(bg=BTN_BLUE, fg="#11111b")
        frame_kiri_list.config(text=" Daftar Dokumen Surat ")
        frame_kanan_detail.config(text=" Detail Surat Generated ")
        
    memuat_data_history()
    
    txt_detail_view.config(state="normal")
    txt_detail_view.delete("1.0", END)
    txt_detail_view.insert(END, "Pilih salah satu item di sebelah kiri untuk melihat detail isi dokumen.")
    txt_detail_view.config(state="disabled")

btn_tab_notulen = Button(frame_hist_tabs, text="Notulen Rapat", font=FONT_NORMAL, bg=BTN_BLUE, fg="#11111b", relief="flat", cursor="hand2", command=lambda: ganti_mode_history("notulen"))
btn_tab_notulen.pack(side="left", padx=(0, 10))

btn_tab_note = Button(frame_hist_tabs, text="Catatan Pribadi", font=FONT_NORMAL, bg=BG_SIDEBAR, fg=FG_TEXT, relief="flat", cursor="hand2", command=lambda: ganti_mode_history("note"))
btn_tab_note.pack(side="left", padx=(0, 10))

btn_tab_docmaker = Button(frame_hist_tabs, text="Dokumen Surat", font=FONT_NORMAL, bg=BG_SIDEBAR, fg=FG_TEXT, relief="flat", cursor="hand2", command=lambda: ganti_mode_history("docmaker"))
btn_tab_docmaker.pack(side="left")

frame_kiri_list = LabelFrame(page_history, text=" Daftar Notulen Rapat ", font=FONT_SUBTITLE, bg=BG_APP, fg=BTN_BLUE, bd=1, padx=10, pady=10)
frame_kiri_list.pack(side="left", fill="y", padx=(30, 10), pady=10)

frame_kanan_detail = LabelFrame(page_history, text=" Isi Dokumen Rapat ", font=FONT_SUBTITLE, bg=BG_APP, fg=BTN_BLUE, bd=1, padx=10, pady=10)
frame_kanan_detail.pack(side="right", fill="both", expand=True, padx=(10, 30), pady=10)

scroll_list_y = Scrollbar(frame_kiri_list, orient="vertical", bg=BG_APP, troughcolor=BG_APP)
scroll_list_y.pack(side="right", fill="y")

listbox_rapat = Listbox(frame_kiri_list, font=FONT_NORMAL, bg=BG_SIDEBAR, fg=FG_TEXT, bd=0, width=28, selectbackground=BTN_BLUE, selectforeground="#11111b", yscrollcommand=scroll_list_y.set)
listbox_rapat.pack(side="top", fill="both", expand=True)
scroll_list_y.config(command=listbox_rapat.yview)

scroll_detail_y = Scrollbar(frame_kanan_detail, orient="vertical", bg=BG_APP, troughcolor=BG_APP)
scroll_detail_y.pack(side="right", fill="y")

txt_detail_view = Text(frame_kanan_detail, font=FONT_NORMAL, bg=BTN_GREEN, fg="#11111b", bd=0, wrap="word", yscrollcommand=scroll_detail_y.set)
txt_detail_view.pack(fill="both", expand=True)
scroll_detail_y.config(command=txt_detail_view.yview)
txt_detail_view.insert(END, "Pilih salah satu item di sebelah kiri untuk melihat detail isi dokumen.")
txt_detail_view.config(state="disabled")

def memuat_data_history():
    global mapping_id_rapat
    listbox_rapat.delete(0, END)
    mapping_id_rapat = [] 
    
    if mode_history == "notulen":
        tabel = "notulen"
    elif mode_history == "note":
        tabel = "daily_notes"
    else:
        tabel = "docmaker_history"
        
    try:
        conn = sqlite3.connect("database_sekretariat.db")
        cursor = conn.cursor()
        
        if mode_history == "docmaker":
            cursor.execute(f"SELECT id, nama_file, timestamp FROM {tabel} ORDER BY id DESC")
        else:
            cursor.execute(f"SELECT id, judul, timestamp FROM {tabel} ORDER BY id DESC")
            
        rows = cursor.fetchall()
        
        for row in rows:
            tanggal_saja = row[2].split(" ")[0]
            listbox_rapat.insert(END, f" {row[1]} ({tanggal_saja})")
            mapping_id_rapat.append(row[0]) 
            
        conn.close()
    except sqlite3.Error as e:
        messagebox.showerror("Error Database", f"Gagal memuat data: {e}")

def buka_file_word_terpilih(path):
    if os.path.exists(path):
        try:
            os.startfile(path)
        except AttributeError:
            # Utk OS Non-Windows
            if sys.platform == "darwin":
                os.system(f"open '{path}'")
            else:
                os.system(f"xdg-open '{path}'")
        except Exception as e:
            messagebox.showerror("Gagal Membuka File", f"Kesalahan: {e}")
    else:
        messagebox.showerror("File Tidak Ditemukan", "File surat sudah dipindahkan atau dihapus dari komputer Anda.")

btn_open_file = Button(frame_kanan_detail, text="Buka File Word", font=FONT_NORMAL, bg=BTN_GREEN, fg="#11111b", relief="flat", cursor="hand2")
# Kita akan sembunyikan dulu secara default
btn_open_file.pack_forget()

def on_select_listbox(event):
    if not listbox_rapat.curselection():
        return
    
    indeks_terpilih = listbox_rapat.curselection()[0]
    id_db = mapping_id_rapat[indeks_terpilih]
    
    if mode_history == "notulen":
        tabel = "notulen"
    elif mode_history == "note":
        tabel = "daily_notes"
    else:
        tabel = "docmaker_history"
    
    try:
        conn = sqlite3.connect("database_sekretariat.db")
        cursor = conn.cursor()
        
        btn_open_file.pack_forget()
        txt_detail_view.config(state="normal")
        txt_detail_view.delete("1.0", END)
        
        if mode_history == "docmaker":
            cursor.execute(f"SELECT nama_file, alamat_file, timestamp FROM {tabel} WHERE id=?", (id_db,))
            data = cursor.fetchone()
            
            if data:
                teks_tampil = f"=== INFORMASI DOKUMEN ===\n\nNama File:\n{data[0]}\n\nLokasi / Path File:\n{data[1]}\n\nWaktu Generate:\n{data[2]}\n"
                txt_detail_view.insert(END, teks_tampil)
                
                btn_open_file.config(command=lambda path=data[1]: buka_file_word_terpilih(path))
                btn_open_file.pack(side="bottom", fill="x", pady=(10, 0))
        else:
            if mode_history == "notulen":
                cursor.execute(f"SELECT judul, hasil_rapat, timestamp FROM {tabel} WHERE id=?", (id_db,))
            else:
                cursor.execute(f"SELECT judul, catatan, timestamp FROM {tabel} WHERE id=?", (id_db,))
                
            data = cursor.fetchone()
            if data:
                teks_tampil = f"--- [{data[2]}] ---\n\nJUDUL:\n{data[0]}\n\nISI DOKUMEN:\n{data[1]}"
                txt_detail_view.insert(END, teks_tampil)
                
        txt_detail_view.config(state="disabled")
        conn.close()
    except sqlite3.Error as e:
        messagebox.showerror("Error Database", f"Gagal mengambil detail dokumen: {e}")

listbox_rapat.bind("<<ListboxSelect>>", on_select_listbox)

def hapus_rapat_terpilih():
    if not listbox_rapat.curselection():
        messagebox.showwarning("Peringatan", "Pilih dokumen di sebelah kiri yang ingin dihapus terlebih dahulu.")
        return
        
    indeks_terpilih = listbox_rapat.curselection()[0]
    id_db = mapping_id_rapat[indeks_terpilih]
    
    konfirmasi = messagebox.askyesno("Konfirmasi Hapus", "Apakah Anda yakin ingin menghapus data ini dari History? (Ini tidak menghapus file di komputer Anda)")
    if not konfirmasi:
        return
        
    if mode_history == "notulen":
        tabel = "notulen"
    elif mode_history == "note":
        tabel = "daily_notes"
    else:
        tabel = "docmaker_history"
        
    try:
        conn = sqlite3.connect("database_sekretariat.db")
        cursor = conn.cursor()
        cursor.execute(f"DELETE FROM {tabel} WHERE id=?", (id_db,))
        conn.commit()
        conn.close()
        
        memuat_data_history()
        
        txt_detail_view.config(state="normal")
        txt_detail_view.delete("1.0", END)
        txt_detail_view.insert(END, "Dokumen berhasil dihapus dari history.")
        txt_detail_view.config(state="disabled")
        btn_open_file.pack_forget()
        
    except sqlite3.Error as e:
        messagebox.showerror("Error Database", f"Gagal menghapus data: {e}")

btn_hapus_rapat = Button(frame_kiri_list, text="Hapus Data Terpilih", font=FONT_NORMAL, bg=BTN_RED, fg="#11111b", relief="flat", cursor="hand2", command=hapus_rapat_terpilih)
btn_hapus_rapat.pack(side="bottom", fill="x", pady=(10, 0))

# ==================== MAIN LOOP ====================
buka_halaman_summarizer() 
window.mainloop()