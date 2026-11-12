import streamlit as st
import pandas as pd
import pdfplumber
import requests
import base64
import mimetypes
import json
import os
import re
import io
import docx
import datetime
import math
import openpyxl
import gspread
import mimetypes
from google.oauth2.service_account import Credentials
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docxtpl import DocxTemplate
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Modul Google Drive API
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# Helper Format Tanggal Hari Ini ke Bahasa Indonesia
NAMA_BULAN = [
    "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
]
now = datetime.datetime.now()
tanggal_hari_ini = f"{now.day} {NAMA_BULAN[now.month]} {now.year}"

# 1. KONFIGURASI HALAMAN (HARUS PALING PERTAMA DARI ELEMEN STREAMLIT)
st.set_page_config(
    page_title="Master Rekap & Generator Dokumen BPJS",
    page_icon="📄",
    layout="wide"
)

# 2. INJECT CUSTOM CSS TEMA BIRU & ORANYE
st.markdown("""
    <style>
        /* 1. Sembunyikan scrollbar di sidebar & rapatkan elemen */
        [data-testid="stSidebar"] > div:first-child {
            overflow: hidden !important;
        }
        [data-testid="stSidebar"] [data-testid="stVerticalBlock"] {
            gap: 0.5rem !important;
        }

        /* 2. Background Sidebar & Border Oranye */
        [data-testid="stSidebar"] {
            background-color: #F0F4F9 !important;
            border-right: 3px solid #FF7A00 !important;
        }

        /* 3. TEKS JUDUL & LABEL SIDEBAR: Biru Navy Gelap */
        [data-testid="stSidebar"] label, 
        [data-testid="stSidebar"] p, 
        [data-testid="stSidebar"] h1, 
        [data-testid="stSidebar"] h2, 
        [data-testid="stSidebar"] h3 {
            color: #0A2540 !important;
            font-weight: 700 !important;
        }

        /* 4. ISIAN DALAM BOX INPUT SIDEBAR: Background Putih, Teks Navy Solid */
        [data-testid="stSidebar"] input, 
        [data-testid="stSidebar"] div[data-baseweb="select"] * {
            background-color: #FFFFFF !important;
            color: #0A2540 !important;
            font-weight: 600 !important;
        }

        /* 5. JUDUL DI AREA UTAMA: Biru Muda Cyan */
        h1, h2, h3, [data-testid="stHeader"] {
            color: #4CC9FE !important;
            font-weight: 700 !important;
        }

        /* 6. Garis Pemisah Warna Oranye */
        hr {
            border-top: 2px solid #FF7A00 !important;
            margin: 1rem 0 !important;
        }

        /* 7. Custom Tombol Primary */
        div.stButton > button[kind="primary"] {
            background-color: #0F4C81 !important;
            color: #FFFFFF !important;
            border: 2px solid #FF7A00 !important;
            border-radius: 8px !important;
            font-weight: bold !important;
            transition: all 0.3s ease !important;
        }
        div.stButton > button[kind="primary"]:hover {
            background-color: #FF7A00 !important;
            color: #FFFFFF !important;
            border-color: #0F4C81 !important;
            box-shadow: 0px 4px 12px rgba(255, 122, 0, 0.4) !important;
        }
        div.stButton > button[kind="primary"]:active {
            background-color: #E66E00 !important;
            color: #FFFFFF !important;
        }

        /* 8. Radio Button Accent (Oranye) */
        div[role="radiogroup"] label[data-baseweb="radio"] div:first-child {
            background-color: #FF7A00 !important;
        }
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# FUNGSI KEAMANAN & HELPER
# ==============================================================================
def safe_int(val):
    try:
        if pd.isna(val) or val is None or str(val).strip() == "" or str(val).strip().lower() == "none":
            return 0
        return int(float(val))
    except (ValueError, TypeError):
        return 0

# Helper pembulatan khusus (0.01 - 0.50 tetap/turun, 0.51 - 0.99 naik 1)
def custom_round(val):
    try:
        val_float = float(val)
        integer_part = int(val_float)
        decimal_part = val_float - integer_part
        if decimal_part > 0.50:
            return integer_part + 1
        else:
            return integer_part
    except (ValueError, TypeError):
        return 0

# ==============================================================================
# FUNGSI HELPER UPLOAD GOOGLE DRIVE UNIVERSAL (DENGAN DYNAMIC SUBFOLDER)
# ==============================================================================
def get_or_create_subfolder(drive_service, parent_id, folder_name):
    try:
        query = f"'{parent_id}' in parents and name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        results = drive_service.files().list(
            q=query, 
            fields="files(id, name)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        items = results.get('files', [])

        if items:
            return items[0]['id']
        else:
            folder_metadata = {
                'name': folder_name,
                'mimeType': 'application/vnd.google-apps.folder',
                'parents': [parent_id]
            }
            folder = drive_service.files().create(
                body=folder_metadata, 
                fields='id',
                supportsAllDrives=True
            ).execute()
            
            folder_id = folder.get('id')
            
            # Pindahkan permission agar tidak mentok di Service Account
            try:
                drive_service.permissions().create(
                    fileId=folder_id,
                    body={'type': 'anyone', 'role': 'writer'},
                    supportsAllDrives=True
                ).execute()
            except Exception:
                pass

            return folder_id
    except Exception as e:
        return parent_id

def upload_to_dynamic_drive_folder(file_bytes, file_name, jenis_bpjs, tahun_str, bulan_str, cabang_str):
    """
    Mengunggah berkas via Google Apps Script Web App Bridge.
    (100% Menggunakan Kuota Drive Akun Utama - Bebas Error 403 Storage Quota).
    """
    try:
        gas_url = st.secrets.get("GAS_WEB_APP_URL", "")
        if not gas_url:
            return False, "GAS_WEB_APP_URL belum dikonfigurasi di secrets.toml!"

        mime_type, _ = mimetypes.guess_type(file_name)
        if not mime_type:
            mime_type = 'application/octet-stream'

        # Konversi file bytes ke format Base64 string
        base64_data = base64.b64encode(file_bytes).decode('utf-8')
        cabang_clean = cabang_str if cabang_str.startswith("Cabang") else f"Cabang {cabang_str.replace('Cab.', '').strip()}"

        payload = {
            "fileName": file_name,
            "mimeType": mime_type,
            "fileData": base64_data,
            "jenis": jenis_bpjs,
            "tahun": str(tahun_str),
            "bulan": str(bulan_str),
            "cabang": cabang_clean
        }

        # Kirim data ke Google Apps Script
        response = requests.post(gas_url, json=payload, timeout=60)
        res_json = response.json()

        if res_json.get("status") == "success":
            return True, res_json.get("fileId")
        else:
            return False, res_json.get("message")

    except Exception as e:
        return False, str(e)

# Fungsi Sinkronisasi Data BPJS Ketenagakerjaan
def sync_to_google_sheet(df_edited, nama_tab_sheet, is_icon=False):
    try:
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
        client = gspread.authorize(creds)
        
        if is_icon:
            sheet_url = st.secrets.get("SPREADSHEET_URL_ICON", st.secrets["SPREADSHEET_URL"])
            col_npp_search = 8   # Kolom H
            col_spk = 5          # Kolom E
            col_tk = 9           # Kolom I
            col_jkk = 14         # Kolom N
            col_jht = 15         # Kolom O
            col_jkm = 16         # Kolom P
            col_jpn = 17         # Kolom Q
        else:
            sheet_url = st.secrets["SPREADSHEET_URL"]
            col_npp_search = 9   # Kolom I
            col_spk = 5          # Kolom E
            col_tk = 10          # Kolom J
            col_jkk = 15         # Kolom O
            col_jht = 16         # Kolom P
            col_jkm = 17         # Kolom Q
            col_jpn = 18         # Kolom R
            
        sh = client.open_by_url(sheet_url)
        worksheet = sh.worksheet(nama_tab_sheet)
        npp_list = worksheet.col_values(col_npp_search)
        
        updated_count = 0
        records = df_edited.to_dict(orient="records")
        
        for row in records:
            npp_pdf = str(row["NPP"]).strip()
            if npp_pdf in npp_list:
                row_idx = npp_list.index(npp_pdf) + 1
                worksheet.update_cell(row_idx, col_spk, str(row["Nomor Kontrak"]))
                worksheet.update_cell(row_idx, col_tk, safe_int(row["Jumlah Tenaga Kerja"]))
                worksheet.update_cell(row_idx, col_jkk, custom_round(row["Total Iuran JKK"]))
                worksheet.update_cell(row_idx, col_jht, custom_round(row["Total Iuran JHT"]))
                worksheet.update_cell(row_idx, col_jkm, custom_round(row["Total Iuran JKM"]))
                worksheet.update_cell(row_idx, col_jpn, custom_round(row["Total Iuran JPN"]))
                updated_count += 1
                
        return True, updated_count
    except Exception as e:
        return False, str(e)

# Fungsi Auto-Sync Khusus BPJS Kesehatan (Direct Index H, I, K)
def sync_to_google_sheet_kesehatan(df_data, target_tab_name, is_icon=False):
    """
    Mengisi/meng-update data BPJS Kesehatan secara langsung
    berdasarkan lokasi Kolom D (SPK), H (VA), I (TK), dan K (Total Iuran).
    """
    try:
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
        client = gspread.authorize(creds)
        
        sheet_url = st.secrets.get("SPREADSHEET_URL_ICON") if is_icon else st.secrets.get("SPREADSHEET_URL_KES", st.secrets["SPREADSHEET_URL"])
        sh = client.open_by_url(sheet_url)
        worksheet = sh.worksheet(target_tab_name)
        
        all_values = worksheet.get_all_values()
        if not all_values:
            return False, "Worksheet target kosong!"

        # Didefinisikan posisi kolom secara presisi (0-based indexing Python):
        spk_col_idx = 3      # Kolom D (Nomor SPK / Kontrak)
        va_col_idx = 7       # Kolom H (Nomor VA Mandiri)
        tk_col_idx = 8       # Kolom I (JUMLAH TK DI INVOICE)
        iuran_col_idx = 10   # Kolom K (IURAN)
        
        start_row_idx = 4    # Baris data dimulai dari baris ke-5/6 (0-indexed = 4)

        updated_count = 0
        records = df_data.to_dict(orient="records")

        for row in records:
            val_spk = str(row.get("Nomor Kontrak", "")).strip()
            val_va = str(row.get("Nomor VA", "")).strip()
            val_iuran = float(row.get("Total Iuran", 0))
            val_tk = int(row.get("Jumlah Tenaga Kerja", 0))

            clean_pdf_va = re.sub(r'\D', '', val_va)
            if not clean_pdf_va:
                continue

            matched_row_idx = -1

            # Cari baris yang Nomor VA-nya cocok dari baris data ke bawah
            for r_idx in range(start_row_idx, len(all_values)):
                current_row = all_values[r_idx]
                if len(current_row) > va_col_idx:
                    sheet_va = str(current_row[va_col_idx]).strip()
                    clean_sheet_va = re.sub(r'\D', '', sheet_va)
                    
                    if clean_pdf_va and clean_pdf_va in clean_sheet_va:
                        matched_row_idx = r_idx + 1  # Baris gspread (1-based index)
                        break

            # Eksekusi Update ke Google Sheet jika VA cocok
            if matched_row_idx != -1:
                # 1. Update Kolom D (Nomor SPK) -> Index 4 (1-based)
                if val_spk and val_spk != "-":
                    worksheet.update_cell(matched_row_idx, spk_col_idx + 1, val_spk)
                # 2. Update Kolom I (Jumlah TK) -> Index 9 (1-based)
                worksheet.update_cell(matched_row_idx, tk_col_idx + 1, val_tk)
                # 3. Update Kolom K (Iuran) -> Index 11 (1-based)
                worksheet.update_cell(matched_row_idx, iuran_col_idx + 1, val_iuran)
                
                updated_count += 1

        return True, f"{updated_count} baris data VA"

    except Exception as e:
        return False, str(e)

# Helper Ekstraksi Periode
def extract_periode(tanggal_str):
    if not tanggal_str:
        return "-"
    parts = tanggal_str.strip().split()
    if len(parts) >= 3:
        return f"{parts[-2]} {parts[-1]}"
    elif len(parts) == 2:
        return tanggal_str
    return tanggal_str

# Helper Terbilang
def terbilang(n):
    satuan = ["", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas"]
    n = int(abs(n))
    if n < 12:
        return satuan[n]
    elif n < 20:
        return terbilang(n - 10) + " Belas"
    elif n < 100:
        return terbilang(n // 10) + " Puluh " + terbilang(n % 10)
    elif n < 200:
        return "Seratus " + terbilang(n - 100)
    elif n < 1000:
        return terbilang(n // 100) + " Ratus " + terbilang(n % 100)
    elif n < 2000:
        return "Seribu " + terbilang(n - 1000)
    elif n < 1000000:
        return terbilang(n // 1000) + " Ribu " + terbilang(n % 1000)
    elif n < 1000000000:
        return terbilang(n // 1000000) + " Juta " + terbilang(n % 1000000)
    elif n < 1000000000000:
        return terbilang(n // 1000000000) + " Miliar " + terbilang(n % 1000000000)
    return str(n)

# Helper Border Cell Word
def set_cell_border(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(r'''
        <w:tcBorders %s>
            <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
            <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
            <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        </w:tcBorders>
    ''' % nsdecls('w'))
    tcPr.append(tcBorders)

# ==============================================================================
# DATABASE KONTRAK & SIDEBAR
# ==============================================================================
@st.cache_data
def load_database():
    possible_paths = [
        "database_kontrak.json",
        "database.json",
        os.path.join("project-bpjs", "database_kontrak.json")
    ]
    for path in possible_paths:
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
    return {}

db_kontrak = load_database()

st.sidebar.header("📌 Pengaturan Dokumen Word")
wilayah_list = list(db_kontrak.keys()) if db_kontrak else ["Cab. Palembang", "Cab. Medan", "Cab. Sibolga Nias", "Cab. Bangka Belitung"]
selected_wilayah = st.sidebar.selectbox("Pilih Cabang / Wilayah", wilayah_list)

wilayah_display = re.sub(r'(?i)\bcab\.?\s*', '', selected_wilayah).strip()
data_cabang = db_kontrak.get(selected_wilayah, {})

nama_kota = data_cabang.get("kota", wilayah_display.title())
default_manager = data_cabang.get("manager", "NAMA MANAGER")
kontrak_options = data_cabang.get("spk", ["-"])

if not kontrak_options:
    kontrak_options = ["-"]

st.sidebar.markdown("---")
st.sidebar.header("📝 Informasi Header Surat")
nomor_surat = st.sidebar.text_input("Nomor Surat", "001/SP/HPI/X/2026")
tanggal_surat = st.sidebar.text_input("Tanggal Surat", value=tanggal_hari_ini) 
nama_manager = st.sidebar.text_input("Nama Manager Cabang", value=default_manager)

if os.path.exists("logo.png"):
    st.sidebar.markdown("---")
    col1, col2, col3 = st.sidebar.columns([1, 3, 1])
    with col2:
        st.image("logo.png", use_container_width=True)

# ==============================================================================
# TAB UTAMA
# ==============================================================================
tab1, tab2 = st.tabs(["🏢 BPJS Ketenagakerjaan (EPS/NPP/F2A)", "🏥 BPJS Kesehatan (Billing Statement)"])

# ==============================================================================
# TAB 1: BPJS KETENAGAKERJAAN
# ==============================================================================
with tab1:
    st.subheader("Rekap PDF & Generator Surat BPJS Ketenagakerjaan")
    uploaded_files_tk = st.file_uploader(
        "Silahkan upload NPP/EPS/F2A (Bisa banyak & campur format sekaligus)", 
        type=["pdf"], 
        accept_multiple_files=True,
        key="uploader_tk"
    )

    if uploaded_files_tk:
        st.info(f"Memproses {len(uploaded_files_tk)} file BPJS Ketenagakerjaan...")
        all_data_tk = []
        
        def parse_amount(val_str):
            if not val_str:
                return 0.0
            clean_str = str(val_str).strip()
            if '.' in clean_str and ',' in clean_str:
                if clean_str.rfind(',') > clean_str.rfind('.'):
                    clean_str = clean_str.replace('.', '').replace(',', '.')
                else:
                    clean_str = clean_str.replace(',', '')
            elif ',' in clean_str and '.' not in clean_str:
                clean_str = clean_str.replace('.', '').replace(',', '.')
            else:
                clean_str = clean_str.replace(',', '')
            try:
                return float(clean_str)
            except ValueError:
                return 0.0

        def lengkapi_npp_dan_kode_iuran(data_tk):
            """Otomatis melengkapi Kode Iuran dari NPP (4+NPP+000) atau sebaliknya"""
            npp = str(data_tk.get("NPP", "-")).strip()
            kode = str(data_tk.get("Kode Iuran", "-")).strip()

            if npp != "-" and npp.isdigit() and (kode == "-" or not kode):
                data_tk["Kode Iuran"] = f"4{npp}000"
            elif kode != "-" and kode.isdigit() and (npp == "-" or not npp):
                if len(kode) > 4:
                    data_tk["NPP"] = kode[1:-3]
            return data_tk

        def parse_f2a_pdf(pdf, default_spk):
            """Membaca PDF Rincian Tenaga Kerja / Form F2A multi-halaman"""
            data_tk = {
                "Nama Perusahaan": "-",
                "NPP": "-",
                "Nomor Kontrak": str(default_spk),
                "Kode Iuran": "-",
                "Jumlah Tenaga Kerja": 0,
                "Total Iuran JKK": 0.0,
                "Total Iuran JHT": 0.0,
                "Total Iuran JKM": 0.0,
                "Total Iuran JPN": 0.0,
                "Total Pembayaran": 0.0
            }

            max_no_tk = 0
            first_page_text = pdf.pages[0].extract_text() or ""
            lines = first_page_text.split('\n')
            for idx, line in enumerate(lines):
                if "NPP" in line and ("Nama Perusahaan" in line or "Nama Unit" in line):
                    if idx + 1 < len(lines):
                        val_line = lines[idx + 1].strip()
                        parts = val_line.split()
                        if parts and parts[0].isdigit():
                            data_tk["NPP"] = str(parts[0])
                            sisa_teks = " ".join(parts[1:]).strip()
                        else:
                            sisa_teks = val_line

                        split_unit = re.split(r'(?=HALEYORA|ICON)', sisa_teks, flags=re.IGNORECASE)
                        split_unit = [s.strip() for s in split_unit if s.strip() and not s.strip().isdigit()]
                        if split_unit:
                            data_tk["Nama Perusahaan"] = split_unit[-1]
                        elif sisa_teks:
                            data_tk["Nama Perusahaan"] = sisa_teks
                    break

            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        if not row:
                            continue
                        
                        first_col = str(row[0]).strip() if row[0] else ""
                        if first_col.isdigit():
                            no_urut = int(first_col)
                            if no_urut > max_no_tk:
                                max_no_tk = no_urut

                        row_joined = " ".join([str(c) for c in row if c])
                        if "Jumlah Seluruhnya" in row_joined:
                            numeric_cells = []
                            for c in row:
                                if c and re.search(r'\d', str(c)) and "Jumlah" not in str(c):
                                    numeric_cells.append(parse_amount(c))

                            if len(numeric_cells) >= 8:
                                jkk = numeric_cells[2]
                                jkm = numeric_cells[3]
                                jht = numeric_cells[4] + numeric_cells[5]
                                jpn = numeric_cells[6] + numeric_cells[7]
                                total = numeric_cells[-1] if len(numeric_cells) >= 11 else (jkk + jkm + jht + jpn)

                                data_tk["Total Iuran JKK"] = jkk
                                data_tk["Total Iuran JKM"] = jkm
                                data_tk["Total Iuran JHT"] = jht
                                data_tk["Total Iuran JPN"] = jpn
                                data_tk["Total Pembayaran"] = total

            data_tk["Jumlah Tenaga Kerja"] = max_no_tk
            return lengkapi_npp_dan_kode_iuran(data_tk)

        def parse_npp_eps_pdf(pdf, default_spk):
            """Membaca PDF Billing NPP / EPS"""
            raw_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    raw_text += text + "\n"

            clean_text = raw_text.replace('\n', ' ').replace('|', '')

            data_tk = {
                "Nama Perusahaan": "-",
                "NPP": "-",
                "Nomor Kontrak": str(default_spk),
                "Kode Iuran": "-",
                "Jumlah Tenaga Kerja": 0, 
                "Total Iuran JKK": 0.0,
                "Total Iuran JHT": 0.0,
                "Total Iuran JKM": 0.0,
                "Total Iuran JPN": 0.0,
                "Total Pembayaran": 0.0
            }

            # 1. Nama Perusahaan
            match_nama_sipp = re.search(r'Nama.?\s:\s*(.*?)\s(?:Total|Jumlah)\s*Tenaga', clean_text, re.IGNORECASE)
            match_nama_eps = re.search(r'PEMBAYARAN\s*UNTUK\s+(.*?)\s+NPP', clean_text, re.IGNORECASE)
            match_nama_eps_ringkas = re.search(r'Nama\s+Perusahaan\s*:\s*(.*?)(?:\s+Total\s+Tenaga|\s+Bulan|$)', clean_text, re.IGNORECASE)

            if match_nama_sipp:
                nama_pt = match_nama_sipp.group(1).replace('Perusahaan', '').replace('Kerja', '').strip()
                data_tk["Nama Perusahaan"] = re.sub(r'\s+', ' ', nama_pt)
            elif match_nama_eps:
                data_tk["Nama Perusahaan"] = re.sub(r'\s+', ' ', match_nama_eps.group(1).strip())
            elif match_nama_eps_ringkas:
                data_tk["Nama Perusahaan"] = re.sub(r'\s+', ' ', match_nama_eps_ringkas.group(1).strip())

            # 2. NPP
            match_npp = re.search(r'NPP\s*:\s*(\d+)', clean_text, re.IGNORECASE)
            if match_npp:
                data_tk["NPP"] = str(match_npp.group(1).strip())

            # 3. Kode Iuran
            match_kode = re.search(r'(?:KODE\s*IURAN\s*:?|KODE\s*PEMBAYARAN)\s*([0-9\s]{10,20})', clean_text, re.IGNORECASE)
            if match_kode: 
                data_tk["Kode Iuran"] = match_kode.group(1).replace(' ', '').strip()

            # 4. Jumlah Tenaga Kerja
            match_tk = re.search(r'(?:JUMLAH|Total)\s*TENAGA\s*KERJA\s*:\s*(\d+)', clean_text, re.IGNORECASE)
            if match_tk:
                data_tk["Jumlah Tenaga Kerja"] = int(match_tk.group(1).strip())

            # 5. Rincian Iuran
            match_jkk = re.search(r'JKK.*?(?:Rp\.?|Rp)?\s*([\d,.]+)', clean_text, re.IGNORECASE)
            if match_jkk:
                data_tk["Total Iuran JKK"] = parse_amount(match_jkk.group(1))

            match_jht = re.search(r'JHT.*?(?:Rp\.?|Rp)?\s*([\d,.]+)', clean_text, re.IGNORECASE)
            if match_jht:
                data_tk["Total Iuran JHT"] = parse_amount(match_jht.group(1))

            match_jkm = re.search(r'JKM.*?(?:Rp\.?|Rp)?\s*([\d,.]+)', clean_text, re.IGNORECASE)
            if match_jkm:
                data_tk["Total Iuran JKM"] = parse_amount(match_jkm.group(1))

            match_jpn = re.search(r'(?:JPN|JP)\s.*?(?:Rp\.?|Rp)?\s*([\d,.]+)', clean_text, re.IGNORECASE)
            if match_jpn:
                data_tk["Total Iuran JPN"] = parse_amount(match_jpn.group(1))

            # 6. Total Pembayaran
            match_total = re.search(r'(?:TOTAL\s*YANG\s*DIBAYARKAN|TOTAL\s*IURAN)\s*(?:\(\s*PEMBULATAN\s*\))?\s*:?\s*(?:Rp\.?|Rp)?\s*([\d,.]+)', clean_text, re.IGNORECASE)
            if match_total:
                data_tk["Total Pembayaran"] = parse_amount(match_total.group(1))

            return lengkapi_npp_dan_kode_iuran(data_tk)

        default_spk = kontrak_options[0] if (kontrak_options and len(kontrak_options) > 0) else "-"

        for file in uploaded_files_tk:
            try:
                with pdfplumber.open(file) as pdf:
                    first_text = (pdf.pages[0].extract_text() or "").lower()
                    if "nomor referensi" in first_text or "iuran jht tk" in first_text or "jumlah seluruhnya" in first_text:
                        data_tk = parse_f2a_pdf(pdf, default_spk)
                    else:
                        data_tk = parse_npp_eps_pdf(pdf, default_spk)

                all_data_tk.append(data_tk)
                
            except Exception as e:
                st.error(f"Terjadi kesalahan saat membaca file {file.name}: {e}")

        if all_data_tk:
            df_raw = pd.DataFrame(all_data_tk)
            
            st.subheader("✏️ Edit Data & Nomor Kontrak (SPK)")
            st.caption("💡 Klik 2x sel di bawah ini untuk mengedit data secara langsung.")

            with st.expander("📌 Lihat Daftar Nomor Kontrak (SPK) Cabang Ini"):
                for opt in kontrak_options:
                    st.code(opt, language="text")

            column_config = {
                "Nama Perusahaan": st.column_config.TextColumn("Nama Perusahaan / Unit"),
                "NPP": st.column_config.TextColumn("NPP"),
                "Nomor Kontrak": st.column_config.TextColumn("Nomor Kontrak (SPK)", width="large"),
                "Kode Iuran": st.column_config.TextColumn("Kode Iuran"),
                "Jumlah Tenaga Kerja": st.column_config.NumberColumn("Jml TK", format="%d"),
                "Total Iuran JKK": st.column_config.NumberColumn("Total JKK", format="Rp %.2f"),
                "Total Iuran JHT": st.column_config.NumberColumn("Total JHT", format="Rp %.2f"),
                "Total Iuran JKM": st.column_config.NumberColumn("Total JKM", format="Rp %.2f"),
                "Total Iuran JPN": st.column_config.NumberColumn("Total JPN", format="Rp %.2f"),
                "Total Pembayaran": st.column_config.NumberColumn("Total Pembayaran (Rp)", format="Rp %.2f")
            }

            edited_df = st.data_editor(
                df_raw,
                column_config=column_config,
                num_rows="dynamic",
                use_container_width=True,
                key=f"editor_combine_{selected_wilayah}_{len(uploaded_files_tk)}"
            )

            edited_df.reset_index(drop=True, inplace=True)
            edited_df["NO"] = [idx + 1 for idx in range(len(edited_df))]

            total_iuran_sum = float(edited_df["Total Pembayaran"].sum())
            total_tk_sum = int(edited_df["Jumlah Tenaga Kerja"].sum())
            terbilang_iuran = (terbilang(total_iuran_sum).strip() + " Rupiah") if total_iuran_sum > 0 else "-"

            st.markdown(f"### **Total Anggaran Iuran:** Rp {total_iuran_sum:,.0f}".replace(",", "."))
            st.markdown(f"**Terbilang:** *{terbilang_iuran}*")

            st.subheader("📥 Unduh Hasil Dokumen")
            col1, col2 = st.columns(2)

            # 1. Download Excel
            with col1:
                st.markdown("##### 📊 Export Master Excel")
                clean_filename_wilayah = wilayah_display.replace(".", "").replace(" ", "_")
                excel_filename = f"Rekap_BPJS_TK_{clean_filename_wilayah}.xlsx"
                
                output_excel = io.BytesIO()
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Rekap BPJS TK"

                font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                font_data = Font(name="Calibri", size=11)
                thin_border = Border(
                    left=Side(style='thin', color='D9D9D9'),
                    right=Side(style='thin', color='D9D9D9'),
                    top=Side(style='thin', color='D9D9D9'),
                    bottom=Side(style='thin', color='D9D9D9')
                )

                headers = ["NO", "Nama Perusahaan", "NPP", "Nomor Kontrak", "Kode Iuran", "Jml TK", "Total JKK", "Total JHT", "Total JKM", "Total JPN", "Total Pembayaran"]
                ws.append(headers)
                for col_num in range(1, len(headers) + 1):
                    cell = ws.cell(row=1, column=col_num)
                    cell.font = font_header
                    cell.fill = fill_header
                    cell.alignment = Alignment(horizontal="center", vertical="center")

                records = edited_df.to_dict(orient="records")
                for r_idx, row in enumerate(records, start=2):
                    ws.cell(row=r_idx, column=1, value=int(row["NO"])).alignment = Alignment(horizontal="center")
                    ws.cell(row=r_idx, column=2, value=str(row["Nama Perusahaan"])).alignment = Alignment(horizontal="left")
                    ws.cell(row=r_idx, column=3, value=str(row["NPP"])).alignment = Alignment(horizontal="center")
                    ws.cell(row=r_idx, column=4, value=str(row["Nomor Kontrak"])).alignment = Alignment(horizontal="left")
                    ws.cell(row=r_idx, column=5, value=str(row["Kode Iuran"])).alignment = Alignment(horizontal="center")
                    ws.cell(row=r_idx, column=6, value=safe_int(row["Jumlah Tenaga Kerja"])).alignment = Alignment(horizontal="center")
                    
                    for c_off, col_name in enumerate(["Total Iuran JKK", "Total Iuran JHT", "Total Iuran JKM", "Total Iuran JPN", "Total Pembayaran"], start=7):
                        c_cell = ws.cell(row=r_idx, column=c_off, value=float(row[col_name]))
                        c_cell.number_format = '"Rp "#,##0.00'
                        c_cell.alignment = Alignment(horizontal="right")

                    for c_idx in range(1, 12):
                        c = ws.cell(row=r_idx, column=c_idx)
                        c.font = font_data
                        c.border = thin_border

                for col in ws.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = get_column_letter(col[0].column)
                    ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

                wb.save(output_excel)
                output_excel.seek(0)

                st.download_button(
                    label="📥 Download Master Excel (.xlsx)",
                    data=output_excel.getvalue(),
                    file_name=excel_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # 2. Download Word Docx
            with col2:
                st.markdown("##### 📝 Export Surat Word (.docx)")
                template_path = os.path.join("templates", "Template_DEFAULT_ketenagakerjaan.docx")
                if not os.path.exists(template_path):
                    template_path = os.path.join("templates", "Template_DEFAULT.docx")

                if os.path.exists(template_path):
                    doc_tpl = DocxTemplate(template_path)
                    periode_surat = extract_periode(tanggal_surat)

                    context = {
                        "Kota": nama_kota,
                        "Wilayah": wilayah_display,
                        "Tanggal": tanggal_surat,
                        "Periode": periode_surat,
                        "NomorSurat": nomor_surat,
                        "Terbilang": terbilang_iuran,
                        "NamaManager": nama_manager,
                        "TOTAL_TK": str(total_tk_sum),
                        "TOTAL_IURAN": f"Rp {total_iuran_sum:,.0f}".replace(",", ".")
                    }
                    doc_tpl.render(context)
                    
                    temp_stream = io.BytesIO()
                    doc_tpl.save(temp_stream)
                    temp_stream.seek(0)

                    doc_native = docx.Document(temp_stream)

                    target_p = None
                    for p in doc_native.paragraphs:
                        if "sebesar" in p.text.lower() or "anggaran" in p.text.lower():
                            target_p = p
                            break

                    for tbl in list(doc_native.tables):
                        tbl._element.getparent().remove(tbl._element)

                    if target_p:
                        spacer_top = doc_native.add_paragraph()
                        spacer_top.paragraph_format.space_before = Pt(4)
                        spacer_top.paragraph_format.space_after = Pt(4)
                        target_p._p.addnext(spacer_top._p)
                        insert_point = spacer_top
                    else:
                        insert_point = target_p

                    new_table = doc_native.add_table(rows=0, cols=5)
                    new_table.alignment = WD_TABLE_ALIGNMENT.CENTER

                    if insert_point:
                        insert_point._p.addnext(new_table._tbl)

                    hdr_cells = new_table.add_row().cells
                    headers_docx = ["NO", "NPP", "Nomor Kontrak", "Jml TK", "Total Iuran"]
                    widths = [Inches(0.5), Inches(1.3), Inches(3.1), Inches(0.8), Inches(1.8)]
                    
                    for idx, text in enumerate(headers_docx):
                        hdr_cells[idx].text = text
                        hdr_cells[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        hdr_cells[idx].paragraphs[0].runs[0].font.bold = True
                        set_cell_border(hdr_cells[idx])
                        hdr_cells[idx].width = widths[idx]

                    records = edited_df.to_dict(orient="records")
                    for row_data in records:
                        val_iuran = float(row_data["Total Pembayaran"])
                        formatted_iuran = f"Rp\t{val_iuran:,.0f}".replace(",", ".")

                        row_cells = new_table.add_row().cells
                        row_cells[0].text = str(row_data["NO"])
                        row_cells[1].text = str(row_data["NPP"])
                        row_cells[2].text = str(row_data["Nomor Kontrak"])
                        row_cells[3].text = str(safe_int(row_data["Jumlah Tenaga Kerja"]))
                        
                        p_iuran = row_cells[4].paragraphs[0]
                        p_iuran.text = formatted_iuran
                        p_iuran.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        p_iuran.paragraph_format.tab_stops.add_tab_stop(Inches(1.75), docx.enum.text.WD_TAB_ALIGNMENT.RIGHT)

                        row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        row_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        row_cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                        row_cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

                        for c_idx in range(5):
                            set_cell_border(row_cells[c_idx])
                            row_cells[c_idx].width = widths[c_idx]

                    total_cells = new_table.add_row().cells
                    total_cells[0].text = "Jumlah"
                    total_cells[3].text = str(total_tk_sum)

                    p_tot_iuran = total_cells[4].paragraphs[0]
                    p_tot_iuran.text = f"Rp\t{total_iuran_sum:,.0f}".replace(",", ".")
                    p_tot_iuran.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    p_tot_iuran.paragraph_format.tab_stops.add_tab_stop(Inches(1.75), docx.enum.text.WD_TAB_ALIGNMENT.RIGHT)
                    if p_tot_iuran.runs:
                        p_tot_iuran.runs[0].font.bold = True

                    total_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    total_cells[0].paragraphs[0].runs[0].font.bold = True
                    total_cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    total_cells[3].paragraphs[0].runs[0].font.bold = True

                    for c_idx in range(5):
                        set_cell_border(total_cells[c_idx])
                        total_cells[c_idx].width = widths[c_idx]

                    spacer_bottom = doc_native.add_paragraph()
                    spacer_bottom.paragraph_format.space_before = Pt(6)
                    spacer_bottom.paragraph_format.space_after = Pt(6)
                    new_table._tbl.addnext(spacer_bottom._p)

                    doc_io = io.BytesIO()
                    doc_native.save(doc_io)
                    doc_io.seek(0)
                    
                    word_filename = f"Surat_BPJS_TK_{clean_filename_wilayah}.docx"
                    st.download_button(
                        label="📥 Download Surat Word (.docx)",
                        data=doc_io.getvalue(),
                        file_name=word_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                else:
                    st.error("Template Word Ketenagakerjaan belum ditemukan di `templates/`!")

            # ==============================================================================
            # AUTO-SYNC GOOGLE SHEETS KHUSUS BPJS TK
            # ==============================================================================
            st.markdown("---")
            st.subheader("🚀 Auto-Sync ke Google Sheets (BPJS TK)")

            target_sheet_choice = st.radio(
                "Pilih Target Google Sheet:",
                ["Sheet Rekap Cabang HPI", "Sheet Rekap Cabang ICON"],
                horizontal=True,
                key="target_sheet_choice_tk"
            )

            try:
                scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
                creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
                client = gspread.authorize(creds)
                
                sheet_url = st.secrets.get("SPREADSHEET_URL_ICON") if target_sheet_choice == "Sheet Rekap Cabang ICON" else st.secrets["SPREADSHEET_URL"]
                sh = client.open_by_url(sheet_url)
                
                available_tabs = [ws.title for ws in sh.worksheets()]
                
                tab_sheet_target_tk = st.selectbox(
                    "Pilih Tab Google Sheet Target:", 
                    available_tabs,
                    key="target_sheet_tk"
                )
            except Exception as e:
                st.warning(f"Tidak dapat memuat daftar tab secara otomatis: {e}")
                tab_sheet_target_tk = st.text_input("Nama Tab Google Sheet Target (Manual)", "01. REKAP 11.26", key="target_sheet_tk")

            if st.button("⚡ Sync Data PDF ke Google Sheet Sekarang", type="primary", key="btn_sync_tk"):
                with st.spinner("Sedang mencocokkan NPP dan mengirim data ke Google Sheet"):
                    is_icon_selected = (target_sheet_choice == "Sheet Rekap Cabang ICON")
                    success, res = sync_to_google_sheet(edited_df, tab_sheet_target_tk, is_icon=is_icon_selected)
                    if success:
                        st.success(f"BERHASIL! {res} data baris NPP berhasil di-Update di Google Sheet!")
                    else:
                        st.error(f"Gagal melakukan sync: {res}")

# ==============================================================================
# TAB 2: BPJS KESEHATAN
# ==============================================================================
with tab2:
    st.subheader("Rekap PDF Billing Statement & Generator Surat BPJS Kesehatan")
    uploaded_files_kes = st.file_uploader(
        "Upload File PDF Billing Statement BPJS Kesehatan (Bisa banyak file sekaligus)", 
        type=["pdf"], 
        accept_multiple_files=True, 
        key="uploader_kes"
    )

    if uploaded_files_kes:
        st.info(f"Memproses {len(uploaded_files_kes)} file BPJS Kesehatan...")
        all_data_kes = []

        for file in uploaded_files_kes:
            try:
                with pdfplumber.open(file) as pdf:
                    raw_text = ""
                    for page in pdf.pages:
                        text = page.extract_text()
                        if text:
                            raw_text += text + "\n"

                clean_text = raw_text.replace('\n', ' ').replace('|', '')
                default_spk = kontrak_options[0] if (kontrak_options and len(kontrak_options) > 0) else "-"

                data_kes = {
                    "Nama Perusahaan": "-",
                    "Nomor VA": "-",
                    "Nomor Kontrak": str(default_spk),
                    "Jumlah Tenaga Kerja": 0,
                    "Total Iuran": 0.0
                }

                # 1. NAMA PERUSAHAAN
                match_nama = re.search(r'Yth\.\s*(.*?)(?:\n\s*JL\.|\s*JL\.|\s*Nomor Virtual|\s*Pelaporan|\n\n|$)', raw_text, re.IGNORECASE | re.DOTALL)
                if match_nama:
                    nama_bu = match_nama.group(1).replace('\n', ' ').strip()
                    nama_bu = re.sub(r'(?:Surat Tagihan|Billing Statement|Pelaporan Penagihan|JKN-KIS).*$', '', nama_bu, flags=re.IGNORECASE).strip()
                    data_kes["Nama Perusahaan"] = re.sub(r'\s+', ' ', nama_bu)

                # 2. NOMOR VA (DIUTAMAKAN MANDIRI / MDR)
                match_va_mdr = re.search(r'(?:MDR|Mandiri)\s*:?\s*([0-9]{13,16})', clean_text, re.IGNORECASE)
                if match_va_mdr:
                    data_kes["Nomor VA"] = match_va_mdr.group(1).strip()
                else:
                    match_va_gen = re.search(r'(?:BNI|BRI|BTN|Virtual Account)\s*:?\s*([0-9]{13,16})', clean_text, re.IGNORECASE)
                    if match_va_gen:
                        data_kes["Nomor VA"] = match_va_gen.group(1).strip()

                # 3. JUMLAH TENAGA KERJA
                match_pegawai = re.search(r'(?:Rp\.?|Rp)?\s*[\d.]+\s*(?:Rp\.?|Rp)?\s*(\d+)\s+\d+\s+\d{2}/\d{2}/\d{4}', clean_text, re.IGNORECASE)
                if match_pegawai:
                    data_kes["Jumlah Tenaga Kerja"] = int(match_pegawai.group(1).strip())
                else:
                    match_peg_alt = re.search(r'(?:Jumlah Pegawai|Number of Employees).*?\s+(\d+)', clean_text, re.IGNORECASE)
                    if match_peg_alt:
                        data_kes["Jumlah Tenaga Kerja"] = int(match_peg_alt.group(1).strip())

                # 4. TOTAL IURAN
                match_total = re.search(r'(?:TOTAL TAGIHAN|Total Balance).?(?:Rp\.?|Rp)\s([\d.]+)', clean_text, re.IGNORECASE)
                if not match_total:
                    match_total = re.search(r'(?:Premi/Iuran|Contribution).?(?:Rp\.?|Rp)\s([\d.]+)', clean_text, re.IGNORECASE)

                if match_total:
                    raw_angka = match_total.group(1).strip()
                    clean_angka = raw_angka.replace('.', '')
                    data_kes["Total Iuran"] = float(clean_angka)

                all_data_kes.append(data_kes)

            except Exception as e:
                st.error(f"Terjadi kesalahan saat membaca file {file.name}: {e}")

        if all_data_kes:
            df_kes_raw = pd.DataFrame(all_data_kes)
            
            st.subheader("✏️ Edit Data & Nomor Kontrak (SPK) BPJS Kesehatan")
            st.caption("💡 Klik 2x sel di bawah ini untuk mengedit data secara langsung.")

            with st.expander("📌 Lihat Daftar Nomor Kontrak (SPK) Cabang Ini"):
                for opt in kontrak_options:
                    st.code(opt, language="text")

            column_config_kes = {
                "Nama Perusahaan": st.column_config.TextColumn("Nama Perusahaan / BU"),
                "Nomor VA": st.column_config.TextColumn("Nomor VA"),
                "Nomor Kontrak": st.column_config.TextColumn("Nomor Kontrak (SPK)", width="large"),
                "Jumlah Tenaga Kerja": st.column_config.NumberColumn("Jml TK", format="%d"),
                "Total Iuran": st.column_config.NumberColumn("Total Tagihan (Rp)", format="Rp %.2f")
            }

            edited_df_kes = st.data_editor(
                df_kes_raw,
                column_config=column_config_kes,
                num_rows="dynamic",
                use_container_width=True,
                key=f"editor_combine_kes_{selected_wilayah}_{len(uploaded_files_kes)}"
            )

            edited_df_kes.reset_index(drop=True, inplace=True)
            edited_df_kes["NO"] = [idx + 1 for idx in range(len(edited_df_kes))]

            total_iuran_kes = float(edited_df_kes["Total Iuran"].sum())
            total_tk_kes = int(edited_df_kes["Jumlah Tenaga Kerja"].sum())
            terbilang_kes = (terbilang(total_iuran_kes).strip() + " Rupiah") if total_iuran_kes > 0 else "-"

            st.markdown(f"### **Total Tagihan BPJS Kesehatan:** Rp {total_iuran_kes:,.0f}".replace(",", "."))
            st.markdown(f"**Terbilang:** *{terbilang_kes}*")

            st.subheader("📥 Unduh Hasil Dokumen BPJS Kesehatan")
            col1, col2 = st.columns(2)

            # 1. Export Excel BPJS Kesehatan
            with col1:
                st.markdown("##### 📊 Export Master Excel")
                clean_filename_wilayah = wilayah_display.replace(".", "").replace(" ", "_")
                excel_filename_kes = f"Rekap_BPJS_Kesehatan_{clean_filename_wilayah}.xlsx"
                
                output_excel_kes = io.BytesIO()
                wb_kes = openpyxl.Workbook()
                ws_kes = wb_kes.active
                ws_kes.title = "Rekap BPJS Kesehatan"

                font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                font_data = Font(name="Calibri", size=11)
                thin_border = Border(
                    left=Side(style='thin', color='D9D9D9'),
                    right=Side(style='thin', color='D9D9D9'),
                    top=Side(style='thin', color='D9D9D9'),
                    bottom=Side(style='thin', color='D9D9D9')
                )

                headers_kes = ["NO", "Nama Perusahaan", "Nomor VA", "Nomor Kontrak", "Jml TK", "Total Tagihan"]
                ws_kes.append(headers_kes)
                for col_num in range(1, len(headers_kes) + 1):
                    cell = ws_kes.cell(row=1, column=col_num)
                    cell.font = font_header
                    cell.fill = fill_header
                    cell.alignment = Alignment(horizontal="center", vertical="center")

                records_kes = edited_df_kes.to_dict(orient="records")
                for r_idx, row in enumerate(records_kes, start=2):
                    ws_kes.cell(row=r_idx, column=1, value=int(row["NO"])).alignment = Alignment(horizontal="center")
                    ws_kes.cell(row=r_idx, column=2, value=str(row["Nama Perusahaan"])).alignment = Alignment(horizontal="left")
                    ws_kes.cell(row=r_idx, column=3, value=str(row["Nomor VA"])).alignment = Alignment(horizontal="center")
                    ws_kes.cell(row=r_idx, column=4, value=str(row["Nomor Kontrak"])).alignment = Alignment(horizontal="left")
                    ws_kes.cell(row=r_idx, column=5, value=safe_int(row["Jumlah Tenaga Kerja"])).alignment = Alignment(horizontal="center")
                    
                    c_cell = ws_kes.cell(row=r_idx, column=6, value=float(row["Total Iuran"]))
                    c_cell.number_format = '"Rp "#,##0.00'
                    c_cell.alignment = Alignment(horizontal="right")

                    for c_idx in range(1, 7):
                        c = ws_kes.cell(row=r_idx, column=c_idx)
                        c.font = font_data
                        c.border = thin_border

                for col in ws_kes.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = get_column_letter(col[0].column)
                    ws_kes.column_dimensions[col_letter].width = max(max_len + 3, 12)

                wb_kes.save(output_excel_kes)
                output_excel_kes.seek(0)

                st.download_button(
                    label="📥 Download Master Excel BPJS Kesehatan",
                    data=output_excel_kes.getvalue(),
                    file_name=excel_filename_kes,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # 2. Export Word Docx BPJS Kesehatan
            with col2:
                st.markdown("##### 📝 Export Surat Word (.docx)")
                template_kes_path = os.path.join("templates", "Template_DEFAULT_kesehatan.docx")
                if not os.path.exists(template_kes_path):
                    template_kes_path = os.path.join("templates", "Template_DEFAULT.docx")

                if os.path.exists(template_kes_path):
                    doc_tpl = DocxTemplate(template_kes_path)
                    periode_kes = extract_periode(tanggal_surat)

                    context = {
                        "Kota": nama_kota,
                        "Wilayah": wilayah_display,
                        "Tanggal": tanggal_surat,
                        "Periode": periode_kes,
                        "NomorSurat": nomor_surat,
                        "Terbilang": terbilang_kes,
                        "NamaManager": nama_manager,
                        "TOTAL_TK": str(total_tk_kes),
                        "TOTAL_IURAN": f"Rp {total_iuran_kes:,.0f}".replace(",", ".")
                    }
                    doc_tpl.render(context)
                    
                    temp_stream = io.BytesIO()
                    doc_tpl.save(temp_stream)
                    temp_stream.seek(0)

                    doc_native = docx.Document(temp_stream)

                    target_p = None
                    for p in doc_native.paragraphs:
                        if "sebesar" in p.text.lower() or "anggaran" in p.text.lower():
                            target_p = p
                            break

                    for tbl in list(doc_native.tables):
                        tbl._element.getparent().remove(tbl._element)

                    if target_p:
                        spacer_top_kes = doc_native.add_paragraph()
                        spacer_top_kes.paragraph_format.space_before = Pt(4)
                        spacer_top_kes.paragraph_format.space_after = Pt(4)
                        target_p._p.addnext(spacer_top_kes._p)
                        insert_point_kes = spacer_top_kes
                    else:
                        insert_point_kes = target_p

                    # Tabel 6 Kolom
                    new_table = doc_native.add_table(rows=0, cols=6)
                    new_table.alignment = WD_TABLE_ALIGNMENT.CENTER

                    if insert_point_kes:
                        insert_point_kes._p.addnext(new_table._tbl)

                    hdr_cells = new_table.add_row().cells
                    headers_kes_docx = ["NO", "NAMA BADAN USAHA", "NOMOR SPK", "NOMOR VA", "JML TK", "TOTAL TAGIHAN"]
                    widths_kes = [Inches(0.4), Inches(2.2), Inches(2.2), Inches(1.3), Inches(0.6), Inches(1.3)]
                    
                    for idx, text in enumerate(headers_kes_docx):
                        hdr_cells[idx].text = text
                        hdr_cells[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        hdr_cells[idx].paragraphs[0].runs[0].font.bold = True
                        set_cell_border(hdr_cells[idx])
                        hdr_cells[idx].width = widths_kes[idx]

                    for r_data in records_kes:
                        val_tagihan = float(r_data["Total Iuran"])
                        formatted_tagihan = f"Rp\t{val_tagihan:,.0f}".replace(",", ".")

                        row_cells = new_table.add_row().cells
                        row_cells[0].text = str(r_data["NO"])
                        row_cells[1].text = str(r_data["Nama Perusahaan"])
                        row_cells[2].text = str(r_data["Nomor Kontrak"])
                        row_cells[3].text = str(r_data["Nomor VA"])
                        row_cells[4].text = str(safe_int(r_data["Jumlah Tenaga Kerja"]))
                        
                        p_tagihan = row_cells[5].paragraphs[0]
                        p_tagihan.text = formatted_tagihan
                        p_tagihan.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        p_tagihan.paragraph_format.tab_stops.add_tab_stop(Inches(1.25), docx.enum.text.WD_TAB_ALIGNMENT.RIGHT)

                        row_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        row_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                        row_cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                        row_cells[3].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                        row_cells[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

                        for c_idx in range(6):
                            set_cell_border(row_cells[c_idx])
                            row_cells[c_idx].width = widths_kes[c_idx]

                    # Baris Total (Jumlah)
                    total_cells = new_table.add_row().cells
                    total_cells[0].text = "Jumlah"
                    total_cells[4].text = str(total_tk_kes)

                    p_tot_kes = total_cells[5].paragraphs[0]
                    p_tot_kes.text = f"Rp\t{total_iuran_kes:,.0f}".replace(",", ".")
                    p_tot_kes.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    p_tot_kes.paragraph_format.tab_stops.add_tab_stop(Inches(1.25), docx.enum.text.WD_TAB_ALIGNMENT.RIGHT)
                    if p_tot_kes.runs:
                        p_tot_kes.runs[0].font.bold = True

                    total_cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    total_cells[0].paragraphs[0].runs[0].font.bold = True
                    total_cells[4].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    total_cells[4].paragraphs[0].runs[0].font.bold = True

                    for c_idx in range(6):
                        set_cell_border(total_cells[c_idx])
                        total_cells[c_idx].width = widths_kes[c_idx]

                    spacer_bottom_kes = doc_native.add_paragraph()
                    spacer_bottom_kes.paragraph_format.space_before = Pt(6)
                    spacer_bottom_kes.paragraph_format.space_after = Pt(6)
                    new_table._tbl.addnext(spacer_bottom_kes._p)

                    doc_io_kes = io.BytesIO()
                    doc_native.save(doc_io_kes)
                    doc_io_kes.seek(0)
                    
                    word_filename_kes = f"Surat_BPJS_Kesehatan_{clean_filename_wilayah}.docx"
                    st.download_button(
                        label="📥 Download Surat Word (.docx)",
                        data=doc_io_kes.getvalue(),
                        file_name=word_filename_kes,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                else:
                    st.error("Template Word belum ditemukan di `templates/`!")

            # ==============================================================================
            # AUTO-SYNC GOOGLE SHEETS KHUSUS BPJS KESEHATAN
            # ==============================================================================
            st.markdown("---")
            st.subheader("🚀 Auto-Sync ke Google Sheets (BPJS Kesehatan)")

            target_sheet_choice_ks = st.radio(
                "Pilih Target Google Sheet:",
                ["Sheet Rekap Utama HPI", "Sheet Khusus Cabang ICON"],
                horizontal=True,
                key="target_sheet_choice_ks"
            )

            # Ambil daftar nama tab secara otomatis dari Google Sheet (dengan penanganan error 503)
            try:
                scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
                creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scope)
                client = gspread.authorize(creds)
                
                sheet_url_ks = st.secrets.get("SPREADSHEET_URL_ICON") if target_sheet_choice_ks == "Sheet Khusus Cabang ICON" else st.secrets.get("SPREADSHEET_URL_KES", st.secrets["SPREADSHEET_URL"])
                sh_ks = client.open_by_url(sheet_url_ks)
                
                # Mengambil daftar tab
                available_tabs_ks = [ws.title for ws in sh_ks.worksheets()]
                
                tab_sheet_target_ks = st.selectbox(
                    "Pilih Tab Google Sheet Target:", 
                    available_tabs_ks,
                    key="target_sheet_ks"
                )
            except Exception as e:
                # Menampilkan pemberitahuan ramah jika server Google 503
                st.warning("⚠️ Server Google Sheet API sedang sibuk (Error 503). Silakan ketik nama tab secara manual di bawah ini:")
                tab_sheet_target_ks = st.text_input("Nama Tab Google Sheet Target (Manual)", "1. NOV 26 SPSH SESUAI BILLI", key="target_sheet_ks")

            if st.button("⚡ Sync Data PDF BPJS Kesehatan ke Google Sheet Sekarang", type="primary", key="btn_sync_ks"):
                with st.spinner("Sedang mencocokkan Nomor VA dan mengirim data ke Google Sheet"):
                    is_icon_selected_ks = (target_sheet_choice_ks == "Sheet Khusus Cabang ICON")
                    
                    success, res = sync_to_google_sheet_kesehatan(edited_df_kes, tab_sheet_target_ks, is_icon=is_icon_selected_ks)
                    
                    if success:
                        st.success(f"BERHASIL! {res} berhasil di-Update ke Google Sheet!")
                    else:
                        st.error(f"Gagal melakukan sync: {res}")

# ==============================================================================
# KOTAK KHUSUS UPLOAD BERKAS KE GOOGLE DRIVE (PDF, RAR, ZIP, EXCEL, WORD, DLL)
# ==============================================================================
st.markdown("---")
st.subheader("☁️ Upload & Arsip Berkas Pendukung ke Google Drive")
st.caption("Pilih hirarki folder tujuan, lalu upload berkas pendukung (PDF, RAR, ZIP, Excel, Word, dll).")

with st.expander("📂 Buka Form Upload Berkas Google Drive", expanded=True):
    col_d1, col_d2 = st.columns(2)
    
    with col_d1:
        drive_kategori = st.selectbox(
            "Target Kategori Folder:", 
            ["BPJS Ketenagakerjaan", "BPJS Kesehatan", "DPLK"], 
            key="drive_kategori"
        )
        drive_cabang = st.selectbox(
            "Target Cabang / Wilayah:", 
            wilayah_list, 
            key="drive_cabang"
        )
        
    with col_d2:
        drive_tahun = st.selectbox(
            "Target Tahun Folder:", 
            ["2026", "2027", "2028"], 
            index=0, 
            key="drive_tahun"
        )
        drive_bulan = st.selectbox(
            "Target Bulan Folder:", 
            [
                "01. Januari", "02. Februari", "03. Maret", "04. April", 
                "05. Mei", "06. Juni", "07. Juli", "08. Agustus", 
                "09. September", "10. Oktober", "11. November", "12. Desember"
            ],
            index=10, 
            key="drive_bulan"
        )

    # File Uploader Bebas Jenis Berkas
    drive_uploaded_files = st.file_uploader(
        "Pilih file yang akan diunggah ke Google Drive (Bisa upload sekaligus banyak file):",
        type=["pdf", "rar", "zip", "xlsx", "xls", "docx", "csv"],
        accept_multiple_files=True,
        key="uploader_google_drive"
    )

    if st.button("📤 Upload Berkas ke Google Drive Sekarang", type="primary", key="btn_upload_drive"):
        if drive_uploaded_files:
            with st.spinner("Sedang mengunggah berkas ke Google Drive..."):
                success_count = 0
                for f in drive_uploaded_files:
                    f_bytes = f.getvalue()
                    ok, res = upload_to_dynamic_drive_folder(
                        file_bytes=f_bytes,
                        file_name=f.name,
                        jenis_bpjs=drive_kategori,
                        tahun_str=drive_tahun,
                        bulan_str=drive_bulan,
                        cabang_str=drive_cabang
                    )
                    if ok:
                        success_count += 1
                    else:
                        st.error(f"Gagal mengunggah {f.name}: {res}")
                
                if success_count > 0:
                    st.success(f"BERHASIL! {success_count} berkas berhasil diarsipkan ke Google Drive di folder: {drive_kategori} / {drive_tahun} / {drive_bulan} / {drive_cabang}")
        else:
            st.warning("Pilih minimal 1 file terlebih dahulu sebelum menekan tombol upload!")
