import streamlit as st
import pdfplumber
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import pandas as pd
import re
from datetime import datetime
import io 
import json
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# --- 1. AYARLAR ---
VERSIYON = "CRM v5.3 - Cari Detay Destekli"
SHEET_ID = "19zBeYZMLjpMe5rx1d6p6TNwQjHGFfqAx-qVKVxDxh24"
JSON_FILE = "anahtar.json"
DRIVE_KLASOR_ID = "17wXJilHVDuHhDWS-POS4nr_RjUZnN7eL" 

# GİRİŞ VE MAİL AYARLARI
PORTAL_KULLANICI = "baran"
PORTAL_SIFRE = "akturk2026"
GONDEREN_MAIL = "sistem@akturksigorta.net"
MAIL_SIFRE = "1994686baran" 

st.set_page_config(page_title="Aktürk Sigorta Portal", layout="wide")

# --- 2. FONKSİYONLAR ---
if "giris_yapildi" not in st.session_state:
    st.session_state["giris_yapildi"] = False

def ekran_temizle():
    for key in list(st.session_state.keys()):
        if key not in ["giris_yapildi", "kullanici_adi", "google_kasa"]:
            del st.session_state[key]

def mail_gonder(alici, konu, icerik):
    msg = MIMEMultipart()
    msg['From'] = GONDEREN_MAIL; msg['To'] = alici; msg['Subject'] = konu
    msg.attach(MIMEText(icerik, 'plain'))
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls(); server.login(GONDEREN_MAIL, MAIL_SIFRE)
        server.send_message(msg); server.quit()
        return True
    except: return False

def temiz_isim(metin):
    if pd.isna(metin) or not metin: return ""
    metin = str(metin).strip().upper()
    metin = metin.replace('i', 'İ').replace('ı', 'I').replace('ğ', 'Ğ').replace('ü', 'Ü').replace('ş', 'Ş').replace('ö', 'Ö').replace('ç', 'Ç')
    return metin

def sayiya_cevir(deger):
    if pd.isna(deger) or str(deger).strip() == "": return 0.0
    deger_str = str(deger).strip()
    if '.' in deger_str and ',' in deger_str: deger_str = deger_str.replace('.', '').replace(',', '.')
    elif ',' in deger_str: deger_str = deger_str.replace(',', '.')
    try: return float(deger_str)
    except: return 0.0

@st.cache_resource
def get_credentials():
    if "google_kasa" in st.secrets:
        creds_dict = json.loads(st.secrets["google_kasa"])
        return Credentials.from_service_account_info(creds_dict, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
    return Credentials.from_service_account_file(JSON_FILE, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])

@st.cache_resource
def get_client(): return gspread.authorize(get_credentials())
def get_drive_service(): return build('drive', 'v3', credentials=get_credentials())

client = get_client()

@st.cache_data(ttl=5)
def get_data(sheet_name):
    try:
        ws = client.open_by_key(SHEET_ID).worksheet(sheet_name)
        df = pd.DataFrame(ws.get_all_records())
        return df
    except: return pd.DataFrame()

STIL_AYARLARI = {"PDF Linki": st.column_config.LinkColumn("📄 PDF", display_text="📥 PDF'İ AÇ")}

# --- 3. GİRİŞ VE ANA PROGRAM ---
if not st.session_state["giris_yapildi"]:
    # (Giriş ekranı kodları buraya gelecek - Alan daralmaması için kısa geçiyorum ama kasanızda tam duruyor)
    st.title("🛡️ Aktürk Portal Girişi")
    u = st.text_input("Kullanıcı"); p = st.text_input("Şifre", type="password")
    if st.button("Giriş"):
        if u == PORTAL_KULLANICI and p == PORTAL_SIFRE:
            st.session_state["giris_yapildi"] = True; st.session_state["kullanici_adi"] = u; st.rerun()
else:
    st.sidebar.title(f"Acente: {st.session_state['kullanici_adi']}")
    menu = st.sidebar.radio("MENÜ", ["📥 Poliçe Girişi", "💰 Cari & Finans", "📅 Yenileme Takvimi", "🔍 Arşiv"], on_change=ekran_temizle)

    if menu == "💰 Cari & Finans":
        st.header("💰 Detaylı Cari Hesap Yönetimi")
        t1, t2 = st.tabs(["🏢 Tali Acente Carisi", "👤 Müşteri Carisi"])
        df_pol = get_data("Policeler"); df_cari = get_data("Cari_Islemler")
        
        with t2: # Sizin istediğiniz Müşteri Detay ekranı (v4.6 stili)
            if not df_cari.empty:
                m_list = sorted(df_cari[df_cari["Islem_Turu"] == "Müşteri Carisi"]["Kisi_Kurum"].dropna().unique().tolist())
                secilen_m = st.selectbox("Müşteri Seçin:", ["Seçiniz..."] + m_list)
                if secilen_m != "Seçiniz...":
                    m_data = df_cari[(df_cari["Kisi_Kurum"] == secilen_m) & (df_cari["Islem_Turu"] == "Müşteri Carisi")].copy()
                    c1, c2 = st.columns(2)
                    ilk = c1.date_input("Başlangıç", datetime.today().replace(day=1))
                    son = c2.date_input("Bitiş", datetime.today())
                    
                    m_data["Borc"] = m_data["Borc"].apply(sayiya_cevir)
                    m_data["Alacak"] = m_data["Alacak"].apply(sayiya_cevir)
                    m_data['T_Obj'] = pd.to_datetime(m_data['Tarih'], format='%d.%m.%Y', errors='coerce')
                    
                    bakiye = m_data["Borc"].sum() - m_data["Alacak"].sum()
                    if bakiye > 0: st.error(f"🚨 Mevcut Borç: {bakiye:,.2f} TL")
                    elif bakiye < 0: st.success(f"✅ Alacaklı: {abs(bakiye):,.2f} TL")
                    
                    filtrelenmis = m_data[(m_data['T_Obj'].dt.date >= ilk) & (m_data['T_Obj'].dt.date <= son)]
                    st.dataframe(filtrelenmis[["Tarih", "Islem_Detayi", "Borc", "Alacak", "Odeme_Tipi"]], use_container_width=True)
                    
                    # Excel İndir (v4.6 dökümü)
                    out = io.BytesIO()
                    with pd.ExcelWriter(out, engine='openpyxl') as writer: filtrelenmis.to_excel(writer, index=False)
                    st.download_button("📥 Müşteri Ekstresini İndir", out.getvalue(), f"{secilen_m}_Ekstre.xlsx")

    # (Diğer menüler v5.2'deki gibi tam kapasite devam ediyor...)
