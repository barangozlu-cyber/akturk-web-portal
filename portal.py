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
VERSIYON = "CRM v5.2 - Tam Entegre Sistem"
SHEET_ID = "19zBeYZMLjpMe5rx1d6p6TNwQjHGFfqAx-qVKVxDxh24"
JSON_FILE = "anahtar.json"
DRIVE_KLASOR_ID = "17wXJilHVDuHhDWS-POS4nr_RjUZnN7eL" 

# GİRİŞ VE MAİL AYARLARI
PORTAL_KULLANICI = "baran"
PORTAL_SIFRE = "akturk2026"
GONDEREN_MAIL = "sistem@akturksigorta.net"
MAIL_SIFRE = "BURAYA_MAIL_SIFRENIZI_YAZIN" # Gmail uygulama şifreniz

st.set_page_config(page_title="Aktürk Sigorta Portal", layout="wide")

# --- 2. ÇEKİRDEK FONKSİYONLAR ---
if "giris_yapildi" not in st.session_state:
    st.session_state["giris_yapildi"] = False

def ekran_temizle():
    """Sayfa geçişlerinde hayalet kutuları yok eder"""
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
    deger_str = str(deger).strip().replace('.', '').replace(',', '.')
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
        if not df.empty:
            for col in ["Müşteri Adı Soyadı", "Kisi_Kurum"]:
                if col in df.columns: df[col] = df[col].apply(temiz_isim)
        return df
    except: return pd.DataFrame()

def drive_pdf_yukle(file_bytes, file_name):
    try:
        drive_service = get_drive_service()
        file_metadata = {'name': file_name, 'parents': [DRIVE_KLASOR_ID]}
        media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype='application/pdf', resumable=True)
        file = drive_service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink', supportsAllDrives=True).execute()
        drive_service.permissions().create(fileId=file.get('id'), body={'type': 'anyone', 'role': 'reader'}, supportsAllDrives=True).execute()
        return file.get('webViewLink')
    except Exception as e: return f"Hata: {e}"

def klasik_analiz(metin):
    data = {"tanzim": "", "baslangic": "", "bitis": "", "musteri": "", "tc_vkn": "", "sirket": "", "urun": "", "p_no": "", "plaka": "", "net_prim": 0.0, "brut_prim": 0.0}
    metin_upper = metin.upper()
    if "ANKARA SİGORTA" in metin_upper: data["sirket"] = "Ankara Sigorta"
    elif "DOĞA SİGORTA" in metin_upper: data["sirket"] = "Doğa Sigorta"
    elif "ALLIANZ" in metin_upper: data["sirket"] = "Allianz Sigorta"
    tarihler = re.findall(r'\b\d{2}[\./-]\d{2}[\./-]\d{4}\b', metin)
    if len(tarihler) >= 3: data["tanzim"], data["baslangic"], data["bitis"] = tarihler[:3]
    tc = re.search(r'\b[0-9]{10,11}\b', metin)
    if tc: data["tc_vkn"] = tc.group()
    plaka = re.search(r'\b[0-9]{2}\s*[A-Z]{1,3}\s*[0-9]{2,4}\b', metin)
    if plaka: data["plaka"] = plaka.group()
    return data

STIL_AYARLARI = {"PDF Linki": st.column_config.LinkColumn("📄 PDF", display_text="📥 PDF'İ AÇ")}

# --- CSS TASARIM ---
st.markdown("""
    <style>
    .stApp { background-color: #0E1117 !important; }
    h1, h2, h3 { color: #F5C518 !important; }
    .stButton>button { background: linear-gradient(135deg, #F5C518 0%, #D4A002 100%) !important; color: black !important; font-weight: bold; }
    .login-box { background-color: #161A22; padding: 40px; border-radius: 15px; border: 1px solid #2D303E; text-align: center; }
    </style>
    """, unsafe_allow_html=True)

# --- 3. GİRİŞ VE ŞİFRE YÖNETİMİ ---
if not st.session_state["giris_yapildi"]:
    st.markdown("<div class='login-box'>", unsafe_allow_html=True)
    st.header("🛡️ Aktürk Sigorta Portal")
    tab1, tab2, tab3 = st.tabs(["🔐 Giriş", "📝 Başvuru", "🔑 Şifremi Unuttum"])
    
    with tab1:
        u = st.text_input("Kullanıcı Adı"); p = st.text_input("Şifre", type="password")
        if st.button("Sisteme Gir"):
            df_u = get_data("Kullanicilar")
            if (u == PORTAL_KULLANICI and p == PORTAL_SIFRE) or (not df_u.empty and u in df_u["Kullanici_Adi"].values and p in df_u[df_u["Kullanici_Adi"]==u]["Sifre"].values[0]):
                st.session_state["giris_yapildi"] = True; st.session_state["kullanici_adi"] = u; st.rerun()
            else: st.error("Hatalı Giriş!")
    
    with tab2:
        b_ad = st.text_input("Ad Soyad"); b_mail = st.text_input("E-posta")
        if st.button("Başvuruyu İlet"):
            if mail_gonder("baran@akturksigorta.net", "Yeni Üye Başvurusu", f"Başvuran: {b_ad}\nMail: {b_mail}"):
                st.success("Başvurunuz Baran Bey'e iletildi.")

    with tab3:
        m = st.text_input("Sistemdeki E-postanız")
        if st.button("Şifremi Gönder"):
            df_u = get_data("Kullanicilar")
            if not df_u.empty and m in df_u["E-posta"].values:
                s = df_u[df_u["E-posta"]==m]["Sifre"].values[0]
                mail_gonder(m, "Portal Şifreniz", f"Şifreniz: {s}")
                st.success("Şifre gönderildi.")
    st.markdown("</div>", unsafe_allow_html=True)

# --- 4. ANA PROGRAM (GİRİŞ YAPILDI) ---
else:
    st.sidebar.title(f"Hoş Geldin, {st.session_state['kullanici_adi']}")
    menu = st.sidebar.radio("İŞLEMLER", ["📥 Poliçe Girişi", "💰 Cari & Finans", "📅 Yenileme Takvimi", "🔎 Genel Arama", "🔍 Tüm Arşiv"], on_change=ekran_temizle)
    
    if menu == "📥 Poliçe Girişi":
        st.header("📥 Poliçe ve Finans Girişi")
        df_urun = get_data("Ayarlar_Urunler"); df_acente = get_data("Ayarlar_Acenteler")
        file = st.file_uploader("PDF Poliçe Seçin", type="pdf")
        p_data = {"tanzim":"","baslangic":"","bitis":"","musteri":"","plaka":"","net_prim":0.0,"brut_prim":0.0}
        f_bytes = None
        if file:
            f_bytes = file.getvalue()
            with pdfplumber.open(io.BytesIO(f_bytes)) as pdf:
                txt = pdf.pages[0].extract_text()
                if txt: p_data.update(klasik_analiz(txt))

        with st.form("p_form", clear_on_submit=True):
            c1,c2,c3 = st.columns(3)
            tan = c1.text_input("Tanzim Tarihi", p_data["tanzim"]); bas = c2.text_input("Başlangıç", p_data["baslangic"]); bit = c3.text_input("Bitiş", p_data["bitis"])
            mus = st.text_input("Müşteri Ad Soyad", p_data["musteri"])
            c4,c5,c6 = st.columns(3)
            sir = c4.text_input("Şirket"); pno = c5.text_input("Poliçe No"); plk = c6.text_input("Plaka", p_data["plaka"])
            c7,c8,c9 = st.columns(3)
            urn = c7.selectbox("Tür", df_urun["Urun_Adi"].tolist() if not df_urun.empty else ["Trafik"])
            net = c8.number_input("Net Prim", value=float(p_data["net_prim"])); brut = c9.number_input("Brüt Prim", value=float(p_data["brut_prim"]))
            acn = st.selectbox("Acente", df_acente["Acente_Adi"].tolist() if not df_acente.empty else ["Merkez"])
            
            if st.form_submit_button("✅ KAYDET"):
                link = drive_pdf_yukle(f_bytes, f"{mus}_{plk}.pdf") if f_bytes else "Yok"
                client.open_by_key(SHEET_ID).worksheet("Policeler").append_row([tan, bas, bit, temiz_isim(mus), "", sir, urn, pno, plk, net, brut, acn, "", "", link])
                st.success("Kayıt Başarılı!")
                st.cache_data.clear()

    elif menu == "💰 Cari & Finans":
        st.header("💰 Gelişmiş Cari Yönetimi")
        t1, t2 = st.tabs(["🏢 Tali Acente", "👤 Müşteri"])
        df_pol = get_data("Policeler"); df_cari = get_data("Cari_Islemler")
        
        with t1:
            acenteler = df_pol[df_pol["Acente"] != "Merkez"]["Acente"].unique().tolist()
            secilen = st.selectbox("Acente Seç", ["Seçiniz..."] + acenteler)
            if secilen != "Seçiniz...":
                f_pol = df_pol[df_pol["Acente"] == secilen].copy()
                st.dataframe(f_pol, use_container_width=True)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer: f_pol.to_excel(writer, index=False)
                st.download_button("📥 Excel İndir", output.getvalue(), f"{secilen}_Ekstre.xlsx")

    elif menu == "📅 Yenileme Takvimi":
        st.header("📅 Yenileme Takvimi (30 Gün)")
        df_pol = get_data("Policeler")
        if not df_pol.empty:
            df_pol["Kalan"] = (pd.to_datetime(df_pol["Bitiş Tarihi"], format="%d.%m.%Y", errors='coerce') - pd.Timestamp.now()).dt.days
            takvim = df_pol[df_pol["Kalan"] <= 30].sort_values("Kalan")
            st.dataframe(takvim[["Bitiş Tarihi", "Kalan", "Müşteri Adı Soyadı", "Plaka"]], use_container_width=True)

    elif menu == "🔍 Tüm Arşiv":
        st.header("🔍 Poliçe Arşivi")
        df_pol = get_data("Policeler")
        st.dataframe(df_pol, column_config=STIL_AYARLARI, use_container_width=True)

    if st.sidebar.button("🚪 Çıkış"):
        st.session_state["giris_yapildi"] = False; st.rerun()
