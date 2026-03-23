import streamlit as st
import pdfplumber
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import pandas as pd
import re
from datetime import datetime, timedelta
import io 
import json
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import difflib

# ==========================================
# 1. TEMEL AYARLAR VE SABİTLER
# ==========================================
VERSIYON = "Aktürk CRM v6.34 - Streamlit Secrets Tam Entegrasyon"
SHEET_ID = "19zBeYZMLjpMe5rx1d6p6TNwQjHGFfqAx-qVKVxDxh24"
DRIVE_KLASOR_ID = "17wXJilHVDuHhDWS-POS4nr_RjUZnN7eL" 
JSON_FILE = "anahtar.json" # Lokalde çalışırken yedek olarak kalabilir, GitHub'a atılmamalı (.gitignore)

# 🔐 ŞİFRELER SECRETS KASASINDAN ÇEKİLİYOR
try:
    PORTAL_KULLANICI = st.secrets["PORTAL_KULLANICI"]
    PORTAL_SIFRE = st.secrets["PORTAL_SIFRE"]
    GONDEREN_MAIL = st.secrets["GONDEREN_MAIL"]
    MAIL_SIFRE = st.secrets["MAIL_SIFRE"]
except KeyError:
    st.error("🚨 Güvenlik Anahtarları (Secrets) bulunamadı! Lütfen '.streamlit/secrets.toml' dosyasını veya Streamlit Cloud ayarlarını kontrol edin.")
    st.stop() # Şifreler yoksa güvenliği sağlamak için uygulamayı anında durdur

st.set_page_config(page_title="Aktürk Sigorta Portal", layout="wide", initial_sidebar_state="auto")

# ==========================================
# 2. YARDIMCI FONKSİYONLAR
# ==========================================
if "giris_yapildi" not in st.session_state:
    st.session_state["giris_yapildi"] = False

def ekran_temizle():
    for key in list(st.session_state.keys()):
        if key not in ["giris_yapildi", "kullanici_adi", "google_kasa"]:
            del st.session_state[key]

def mail_gonder(alici, konu, icerik):
    msg = MIMEMultipart()
    msg['From'] = GONDEREN_MAIL
    msg['To'] = alici
    msg['Subject'] = konu
    msg.attach(MIMEText(icerik, 'plain'))
    try:
        if not MAIL_SIFRE:
            st.error("Mail şifresi Secrets içine tanımlanmamış!")
            return False
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(GONDEREN_MAIL, MAIL_SIFRE)
        server.send_message(msg)
        server.quit()
        return True
    except: return False

def temiz_isim(metin):
    if pd.isna(metin) or not metin: return ""
    metin = str(metin).strip()
    metin = metin.replace('i', 'İ').replace('ı', 'I').replace('ğ', 'Ğ').replace('ü', 'Ü').replace('ş', 'Ş').replace('ö', 'Ö').replace('ç', 'Ç')
    return metin.upper()

def sayiya_cevir(deger):
    if pd.isna(deger) or str(deger).strip() == "": return 0.0
    if isinstance(deger, (int, float)): return float(deger)
    deger_str = str(deger).strip()
    deger_str = re.sub(r'[^\d.,-]', '', deger_str) 
    deger_str = deger_str.rstrip('.,')
    if not deger_str: return 0.0
    if '.' in deger_str and ',' in deger_str:
        if deger_str.rfind(',') > deger_str.rfind('.'):
            deger_str = deger_str.replace('.', '').replace(',', '.')
        else: deger_str = deger_str.replace(',', '')
    elif ',' in deger_str:
        parts = deger_str.split(',')
        if len(parts) > 2 or (len(parts) == 2 and len(parts[1]) == 3): deger_str = deger_str.replace(',', '')
        else: deger_str = deger_str.replace(',', '.')
    elif '.' in deger_str:
        parts = deger_str.split('.')
        if len(parts) > 2 or (len(parts) == 2 and len(parts[1]) == 3): deger_str = deger_str.replace('.', '')
    try: return float(deger_str)
    except: return 0.0

def para_format(deger):
    try:
        temiz_sayi = sayiya_cevir(deger)
        formatted = "{:,.2f}".format(temiz_sayi)
        return formatted.replace(",", "X").replace(".", ",").replace("X", ".") + " TL"
    except: return "0,00 TL"

def df_gorsel_yap(df, para_sutunlari):
    df_gorsel = df.copy()
    for col in para_sutunlari:
        if col in df_gorsel.columns: df_gorsel[col] = df_gorsel[col].apply(para_format)
    return df_gorsel

def tarih_formatla(tarih_degeri):
    if pd.isna(tarih_degeri) or str(tarih_degeri).strip() == "": return datetime.now().strftime("%d.%m.%Y")
    try: return pd.to_datetime(str(tarih_degeri).strip(), dayfirst=True).strftime("%d.%m.%Y")
    except: return datetime.now().strftime("%d.%m.%Y")

def excel_indir(df, buton_metni, dosya_adi):
    df_export = df.copy()
    if not df_export.empty:
        toplamlar = {}
        for col in df_export.columns:
            if col in ["Net Prim", "Brüt Prim", "Şirket Komisyonu", "Şirket Komisyonu (TL)", "Aktürk Sigorta Kazancı", "Borc", "Alacak"]:
                toplamlar[col] = df_export[col].sum()
            elif col == df_export.columns[0]: toplamlar[col] = "GENEL TOPLAM"
            else: toplamlar[col] = ""
        df_export = pd.concat([df_export, pd.DataFrame([toplamlar])], ignore_index=True)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer: df_export.to_excel(writer, index=False)
    st.download_button(f"📥 {buton_metni}", out.getvalue(), f"{dosya_adi}.xlsx")

@st.cache_resource(show_spinner="Bağlantı Kuruluyor...")
def get_credentials():
    try:
        # Önce Streamlit Cloud Secrets veya lokal secrets.toml içindeki google_kasa'yı arar
        if "google_kasa" in st.secrets:
            creds_dict = json.loads(st.secrets["google_kasa"])
            return Credentials.from_service_account_info(creds_dict, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
    except: pass
    
    # Bulamazsa bilgisayarınızdaki anahtar.json dosyasına bakar (Lokal kullanım için yedek)
    return Credentials.from_service_account_file(JSON_FILE, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])

@st.cache_resource
def get_client(): return gspread.authorize(get_credentials())
def get_drive_service(): return build('drive', 'v3', credentials=get_credentials())
client = get_client()

@st.cache_data(ttl=5, show_spinner=False)
def get_data(sheet_name):
    try:
        ws = client.open_by_key(SHEET_ID).worksheet(sheet_name)
        raw_data = ws.get_all_values() 
        if not raw_data: return pd.DataFrame()
        headers = raw_data[0]
        max_len = len(headers)
        rows = []
        for r in raw_data[1:]:
            if len(r) < max_len: r = r + [""] * (max_len - len(r))
            elif len(r) > max_len: r = r[:max_len]
            rows.append(r)
        df = pd.DataFrame(rows, columns=headers)
        if not df.empty:
            df['Sheet_Row'] = df.index + 2 
            for col in ["Müşteri Adı Soyadı", "Kisi_Kurum", "Musteri_Adi"]:
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
    except: return "Yok"

def drive_pdf_sil(link):
    if pd.isna(link) or link == "Yok" or not link: return
    try:
        match = re.search(r'/d/([a-zA-Z0-9_-]+)', str(link))
        if match:
            file_id = match.group(1)
            drive_service = get_drive_service()
            drive_service.files().delete(fileId=file_id, supportsAllDrives=True).execute()
    except: pass 

def klasik_analiz(metin):
    data = {"tanzim": "", "baslangic": "", "bitis": "", "musteri": "", "tc_vkn": "", "sirket": "", "urun": "", "p_no": "", "plaka": "", "net_prim": 0.0, "brut_prim": 0.0}
    metin_upper = metin.upper()
    
    sirketler = {
        "ANKARA SİGORTA": "Ankara Sigorta", "DOĞA SİGORTA": "Doğa Sigorta", "ALLIANZ": "Allianz Sigorta",
        "HDI SİGORTA": "HDI Sigorta", "HDİ": "HDI Sigorta", "HEPİYİ": "Hepiyi Sigorta", "RAY SİGORTA": "Ray Sigorta",
        "SOMPO": "Sompo Sigorta", "TÜRKİYE SİGORTA": "Türkiye Sigorta", "AK SİGORTA": "Ak Sigorta", "ETHICA": "Ethica Sigorta"
    }
    for anahtar, deger in sirketler.items():
        if anahtar in metin_upper:
            data["sirket"] = deger; break
            
    urun_tipleri = {
        "TRAFİK": "Trafik Sigortası", "KASKO": "Kasko", "SAĞLIK": "Sağlık Sigortası", "TSS": "Sağlık Sigortası",
        "DASK": "Dask", "DOĞAL AFET": "Dask", "KONUT": "Konut Sigortası", "İŞYERİ": "İşyeri Sigortası",
        "İMM": "İmm", "ALLRİSK": "İnşaat Allrisk"
    }
    for anahtar, deger in urun_tipleri.items():
        if anahtar in metin_upper:
            data["urun"] = deger; break

    tarihler = re.findall(r'\b\d{2}[\./-]\d{2}[\./-]\d{4}\b', metin)
    if len(tarihler) >= 3: data["tanzim"], data["baslangic"], data["bitis"] = [t.replace("/", ".") for t in tarihler[:3]]
    tc = re.search(r'\b[0-9]{10,11}\b', metin)
    if tc: data["tc_vkn"] = tc.group()
    
    plaka = re.search(r'\b(?:[0-8][0-9]|9[0-8])\s*[A-Z]{1,3}\s*[0-9]{2,4}\b', metin_upper)
    if plaka: data["plaka"] = plaka.group().replace(" ", "")
    return data

STIL_AYARLARI = {"PDF Linki": st.column_config.LinkColumn("📄 Belge", display_text="📥 PDF'İ AÇ")}

# ==========================================
# 3. KURUMSAL LİGHT TEMA VE MOBİL CSS
# ==========================================
st.markdown("""
    <style>
    .stApp { background-color: #F4F7F6 !important; color: #2C3E50 !important; }
    h1, h2, h3 { color: #1A252F !important; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-weight: 700; }
    p, span, label { color: #2C3E50; }
    
    div[data-baseweb="select"] > div, 
    div[data-baseweb="input"] > div,
    textarea {
        background-color: #FFFFFF !important; border: 1px solid #BDC3C7 !important; border-radius: 6px !important;
    }
    
    input, textarea, div[data-baseweb="select"] *, div[data-baseweb="input"] * {
        color: #1A252F !important; font-weight: 500 !important;
    }

    ul[data-baseweb="menu"] { background-color: #FFFFFF !important; }
    li[data-baseweb="menu-item"] { color: #1A252F !important; }
    li[data-baseweb="menu-item"]:hover { background-color: #F4F7F6 !important; }

    [data-testid="stDataEditor"] textarea,
    [data-testid="stDataEditor"] input,
    .glide-data-grid textarea,
    .gdg-input {
        background-color: #FFFFFF !important;
        color: #1A252F !important;
        caret-color: #1A252F !important;
        font-weight: 600 !important;
    }

    .stButton>button { 
        background: linear-gradient(135deg, #2980B9 0%, #2471A3 100%) !important; 
        color: #FFFFFF !important; font-weight: 600 !important; 
        border-radius: 6px !important; border: none !important;
        transition: all 0.3s ease !important; box-shadow: 0 4px 6px rgba(41, 128, 185, 0.2) !important; 
    }
    .stButton>button:hover { transform: translateY(-2px); box-shadow: 0 6px 12px rgba(41, 128, 185, 0.4) !important; color: #FFFFFF !important; }
    
    .login-box { background-color: #FFFFFF; padding: 40px; border-radius: 12px; border: 1px solid #EAECEE; text-align: center; max-width: 450px; margin: auto; margin-top: 5vh; box-shadow: 0 10px 30px rgba(0,0,0,0.05); }
    div[data-testid="stForm"] { background-color: #FFFFFF; padding: 25px; border-radius: 12px; border: 1px solid #EAECEE; box-shadow: 0 4px 12px rgba(0,0,0,0.03); }
    div[data-testid="stMetric"] { background-color: #FFFFFF !important; border-radius: 8px !important; padding: 15px 20px !important; border: 1px solid #EAECEE !important; border-left: 4px solid #2980B9 !important; box-shadow: 0 2px 8px rgba(0,0,0,0.03) !important; }
    div[data-testid="stMetricValue"] { color: #2C3E50 !important; font-weight: bold !important; }
    div[data-testid="stMetricLabel"] { color: #7F8C8D !important; font-size: 14px !important; }
    
    [data-testid="stSidebar"] { background-color: #F0F4F9 !important; border-right: none !important; }
    [data-testid="stSidebar"] hr { border-color: #DADCE0 !important; }
    [data-testid="stSidebar"] div[data-baseweb="radio"] > div:first-child { display: none !important; }
    
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label {
        padding: 10px 16px !important; margin-bottom: 4px !important; border-radius: 24px !important;
        transition: background-color 0.2s ease !important; cursor: pointer !important; width: 100% !important;
    }
    
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label p { color: #202124 !important; margin: 0 !important; }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label:hover { background-color: #E1E5EA !important; }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label[data-checked="true"],
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label:has(input:checked) { background-color: #D3E3FD !important; }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label[data-checked="true"] p,
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label:has(input:checked) p { color: #041E49 !important; font-weight: 700 !important; }
    
    @media (max-width: 768px) {
        .login-box { padding: 25px 15px !important; margin-top: 2vh !important; width: 92% !important; }
        div[data-testid="stMetric"] { padding: 10px !important; margin-bottom: 10px !important; }
        div[data-testid="stMetricValue"] { font-size: 1.4rem !important; }
        div[data-testid="stMetricLabel"] { font-size: 12px !important; }
        h1 { font-size: 1.6rem !important; }
        h2 { font-size: 1.4rem !important; }
        h3 { font-size: 1.2rem !important; }
        .stButton>button { padding: 0.6rem !important; width: 100% !important; font-size: 16px !important; }
        div[data-testid="stForm"] { padding: 15px !important; }
    }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 4. GİRİŞ VE ÜYELİK SİSTEMİ 
# ==========================================
if not st.session_state["giris_yapildi"]:
    st.markdown("<div class='login-box'>", unsafe_allow_html=True)
    st.markdown("""
        <div style="background: linear-gradient(135deg, #2980B9 0%, #2471A3 100%); width: 80px; height: 80px; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin: 0 auto 15px auto; box-shadow: 0 4px 15px rgba(41, 128, 185, 0.3);">
            <span style="font-size: 40px; font-weight: 900; color: #FFFFFF;">A</span>
        </div>
        <h2 style='margin-top:0px; margin-bottom: 5px; color:#1A252F;'>Aktürk Sigorta</h2>
        <p style='color: #7F8C8D; font-size: 15px; margin-bottom: 25px;'>Kurumsal Yönetim Portalı</p>
    """, unsafe_allow_html=True)
    
    tab1, tab2, tab3 = st.tabs(["🔐 Giriş Yap", "📝 Başvuru", "🔑 Şifremi Unuttum"])
    
    with tab1:
        u = st.text_input("Kullanıcı Adı")
        p = st.text_input("Şifre", type="password")
        if st.button("Sisteme Güvenli Giriş Yap"):
            if u == PORTAL_KULLANICI and p == PORTAL_SIFRE:
                st.session_state["giris_yapildi"] = True
                st.session_state["kullanici_adi"] = u
                st.rerun()
            else:
                with st.spinner("Kullanıcı doğrulanıyor..."):
                    df_u = get_data("Kullanicilar")
                    giris_onay = False
                    if not df_u.empty:
                        k_kol = "Kullanıcı Adı" if "Kullanıcı Adı" in df_u.columns else "Kullanici_Adi"
                        s_kol = "Şifre" if "Şifre" in df_u.columns else "Sifre"
                        if k_kol in df_u.columns and s_kol in df_u.columns:
                            eslesen_kullanici = df_u[df_u[k_kol] == u]
                            if not eslesen_kullanici.empty:
                                if str(eslesen_kullanici[s_kol].values[0]) == str(p):
                                    giris_onay = True
                if giris_onay:
                    st.session_state["giris_yapildi"] = True
                    st.session_state["kullanici_adi"] = u
                    st.rerun()
                else: 
                    st.error("⚠️ Hatalı kullanıcı adı veya şifre!")
    
    with tab2:
        b_ad = st.text_input("Ad Soyad")
        b_mail = st.text_input("E-posta Adresiniz")
        if st.button("Başvuruyu Gönder"):
            if mail_gonder("baran@akturksigorta.net", "Yeni Acente/Personel Başvurusu", f"Başvuran: {b_ad}\nMail: {b_mail}"): st.success("Talebiniz yönetime iletildi.")
            else: st.error("Mail gönderilemedi. Ayarları kontrol edin.")

    with tab3:
        m = st.text_input("Sistemdeki Kayıtlı E-postanız")
        if st.button("Şifremi Mail At"):
            with st.spinner("Sistem kontrol ediliyor..."):
                df_u = get_data("Kullanicilar")
                mail_bulundu = False
                if not df_u.empty:
                    e_kol = "E-posta" if "E-posta" in df_u.columns else "E_Posta"
                    s_kol = "Şifre" if "Şifre" in df_u.columns else "Sifre"
                    if e_kol in df_u.columns and s_kol in df_u.columns:
                        eslesen_mail = df_u[df_u[e_kol] == m]
                        if not eslesen_mail.empty:
                            mail_bulundu = True
                            s = str(eslesen_mail[s_kol].values[0])
            if mail_bulundu:
                if mail_gonder(m, "Aktürk Portal Şifreniz", f"Güvenli giriş şifreniz: {s}"): st.success("Şifreniz e-posta adresinize gönderildi.")
                else: st.error("Mail gönderilemedi. Ayarları kontrol edin.")
            else: st.error("⚠️ Sistemde böyle bir e-posta bulunamadı.")
                
    st.markdown("</div>", unsafe_allow_html=True)

# ==========================================
# 5. ANA UYGULAMA
# ==========================================
else:
    st.sidebar.markdown("""
        <div style='display: flex; align-items: center; margin-bottom: 25px; margin-top: 10px;'>
            <div style='background: linear-gradient(135deg, #2980B9 0%, #2471A3 100%); width: 38px; height: 38px; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin-right: 12px; box-shadow: 0 4px 10px rgba(41, 128, 185, 0.3);'>
                <span style='font-size: 18px; font-weight: 900; color: #FFFFFF;'>A</span>
            </div>
            <div>
                <h2 style='margin: 0; color: #202124; font-size: 20px;'>Aktürk CRM</h2>
                <span style='color: #5F6368; font-size: 13px; font-weight: 500;'>Kullanıcı: {}</span>
            </div>
        </div>
    """.format(st.session_state['kullanici_adi'].upper()), unsafe_allow_html=True)
    
    menu = st.sidebar.radio("", [
        "📥 Poliçe Girişi", 
        "💰 Cari & Finans", 
        "📅 Yenileme Takvimi", 
        "🔎 Genel Arama", 
        "🛠️ Düzeltme & Silme",
        "⚙️ Ayarlar", 
        "🔍 Tüm Arşiv"
    ], on_change=ekran_temizle, label_visibility="collapsed")

    # ------------------------------------------
    # 5.1 POLİÇE GİRİŞİ 
    # ------------------------------------------
    if menu == "📥 Poliçe Girişi":
        st.header("📥 Yeni İşlem Kaydı")
        
        df_urun = get_data("Ayarlar_Urunler")
        urun_listesi = df_urun["Urun_Adi"].tolist() if not df_urun.empty else ["Trafik Sigortası", "Kasko", "Sağlık Sigortası"]
        dict_urun = dict(zip(df_urun['Urun_Adi'], df_urun['Komisyon_Orani'])) if not df_urun.empty else {}
        
        df_acente = get_data("Ayarlar_Acenteler")
        acente_listesi = df_acente["Acente_Adi"].tolist() if not df_acente.empty else ["Aktürk Sigorta (Merkez)"]
        dict_acente = dict(zip(df_acente['Acente_Adi'], df_acente['Tali_Oran'])) if not df_acente.empty else {}
        acente_listesi.append("➕ YENİ TALİ ACENTE EKLE")

        file = st.file_uploader("PDF Poliçe Seçin (Otomatik Okuma)", type="pdf")
        p_data = {"tanzim":"","baslangic":"","bitis":"","musteri":"","plaka":"","tc_vkn":"","sirket":"","urun":"","p_no":"","net_prim":0.0,"brut_prim":0.0}
        f_bytes = None
        
        if file:
            f_bytes = file.getvalue()
            with st.spinner("📄 PDF Okunuyor..."):
                with pdfplumber.open(io.BytesIO(f_bytes)) as pdf:
                    txt = pdf.pages[0].extract_text()
                    if txt: 
                        p_data.update(klasik_analiz(txt))
                        st.success("PDF Tarandı! Lütfen boşlukları kontrol edip kaydedin.")

        with st.form("police_formu", clear_on_submit=True):
            st.markdown("### 📝 İşlem Türünü Seçin")
            islem_turu = st.radio("Bu kaydın amacı nedir?", ["🟢 Normal Poliçe / Yenileme", "🔴 Tam İptal / Satış (Seneye Yenilenmeyecek)", "🟡 Kısmi İptal / Teminat Düşürme (Seneye Yenilenecek)"])
            st.divider()

            st.subheader("1. Müşteri ve İşlem Detayları")
            c1, c2, c3 = st.columns(3)
            tan = c1.text_input("Tanzim Tarihi", p_data["tanzim"])
            bas = c2.text_input("Başlangıç", p_data["baslangic"])
            bit = c3.text_input("Bitiş", p_data["bitis"])
            
            c4, c5, c6 = st.columns(3)
            mus_girdi = c4.text_input("Müşteri Ad Soyad", p_data["musteri"])
            tc = c5.text_input("TC / VKN", p_data["tc_vkn"])
            ilet = c6.text_input("Telefon / E-mail", "")
            
            c7, c8, c9 = st.columns(3)
            sir = c7.text_input("Sigorta Şirketi", p_data["sirket"])
            pno = c8.text_input("Poliçe No", p_data["p_no"])
            plk = c9.text_input("Plaka", p_data["plaka"])
            
            st.subheader("2. Finans, Komisyon ve Ödeme")
            c10, c11, c12, c12_kom = st.columns(4)
            
            urun_index = urun_listesi.index(p_data["urun"]) if p_data["urun"] in urun_listesi else 0
            urn = c10.selectbox("Ürün Türü", urun_listesi, index=urun_index)
            
            net_val = "" if p_data["net_prim"] == 0.0 else str(p_data["net_prim"])
            brut_val = "" if p_data["brut_prim"] == 0.0 else str(p_data["brut_prim"])
            
            net_girdi = c11.text_input("Net Prim", value=net_val)
            brut_girdi = c12.text_input("Brüt Prim", value=brut_val)
            kom_girdi = c12_kom.text_input("Net Komisyon (Opsiyonel)", value="", help="Boş bırakırsanız ayarlardaki oranı kullanır. Buraya değer girerseniz o oran ezilir.")
            
            c13, c14, c15 = st.columns(3)
            acn = c13.selectbox("İşlemi Yapan Acente", acente_listesi)
            
            yeni_acente_adi = ""
            yeni_acente_orani = 0.0
            if acn == "➕ YENİ TALİ ACENTE EKLE":
                yeni_acente_adi = c13.text_input("Yeni Acente Adını Yazın:")
                yeni_acente_orani_girdi = c13.text_input("Komisyon Oranı (Örn: 70)", value="")
                yeni_acente_orani = float(sayiya_cevir(yeni_acente_orani_girdi))
                
            odm = c14.selectbox("Tahsilat Durumu", ["Nakit Alındı", "Müşteri Kredi Kartı", "Havale", "Acenteye Borçlanıldı (Cari)"])
            taksit = c15.number_input("Taksit", min_value=1, value=1)
            adr = st.text_area("Adres", "")
            
            if st.form_submit_button("✅ POLİÇEYİ, CARİYİ VE PDF'İ KAYDET"):
                with st.spinner("Sisteme İşleniyor..."):
                    doc = client.open_by_key(SHEET_ID)
                    aktif_acente = acn
                    mus = temiz_isim(mus_girdi)
                    
                    net = float(sayiya_cevir(net_girdi))
                    brut = float(sayiya_cevir(brut_girdi))
                    kom_manuel = float(sayiya_cevir(kom_girdi))
                    
                    islem_notu = ""
                    if islem_turu != "🟢 Normal Poliçe / Yenileme":
                        net = -abs(net)
                        brut = -abs(brut)
                        kom_manuel = -abs(kom_manuel) if kom_manuel > 0 else 0.0
                        if "Tam İptal" in islem_turu:
                            sir = f"{sir} (İPTAL-SATIŞ)"
                            islem_notu = "SATIŞ/TAM İPTAL"
                        else:
                            sir = f"{sir} (İPTAL-ZEYL)"
                            islem_notu = "KISMİ İPTAL/ZEYL"
                    
                    link = drive_pdf_yukle(f_bytes, f"{mus}_{plk}_{sir}.pdf") if f_bytes else "Yok"
                    
                    u_oran = float(sayiya_cevir(dict_urun.get(urn, 0.0)))
                    if u_oran > 1: u_oran /= 100 
                    
                    if str(kom_girdi).strip() != "" and kom_manuel != 0.0:
                        sirket_komisyonu = kom_manuel
                    else:
                        sirket_komisyonu = net * u_oran
                    
                    if acn == "➕ YENİ TALİ ACENTE EKLE" and yeni_acente_adi != "":
                        aktif_acente = yeni_acente_adi
                        doc.worksheet("Ayarlar_Acenteler").append_row([yeni_acente_adi, yeni_acente_orani], value_input_option='USER_ENTERED')
                        t_oran = yeni_acente_orani
                    else:
                        t_oran = float(sayiya_cevir(dict_acente.get(aktif_acente, 0.0)))
                    if t_oran > 1: t_oran /= 100 
                    
                    akturk_kazanci = float(sirket_komisyonu * t_oran)
                    islem_tarihi = tarih_formatla(tan)
                    
                    ws_pol = doc.worksheet("Policeler")
                    headers = ws_pol.row_values(1)
                    if "Şirket Komisyonu" not in headers:
                        ws_pol.update_cell(1, len(headers)+1, "Şirket Komisyonu")
                        headers.append("Şirket Komisyonu")
                    
                    row_dict = {
                        "Tanzim Tarihi": islem_tarihi, "Başlangıç Tarihi": bas, "Bitiş Tarihi": bit,
                        "Müşteri Adı Soyadı": mus, "TC / VKN": tc, "Sigorta Şirketi": sir,
                        "Sigorta Türü": urn, "Poliçe No": pno, "Plaka": plk,
                        "Net Prim": net, "Brüt Prim": brut, "Şirket Komisyonu": sirket_komisyonu,
                        "Acente": aktif_acente, "Adres": adr, "Telefon / E-mail": ilet, "PDF Linki": link
                    }
                    yeni_satir = [row_dict.get(h, "") for h in headers]
                    ws_pol.append_row(yeni_satir, value_input_option='USER_ENTERED')
                    
                    aciklama = f"{sir} - {urn} - Plaka: {plk}"
                    yeni_satirlar = []
                    if islem_notu: yeni_satirlar.append([islem_tarihi, "Müşteri Carisi", mus, f"{islem_notu} İADESİ - {aciklama}", brut, 0.0, odm, taksit])
                    else: yeni_satirlar.append([islem_tarihi, "Müşteri Carisi", mus, aciklama, brut, 0.0, odm, taksit])
                    
                    if aktif_acente != "Aktürk Sigorta (Merkez)":
                        yeni_satirlar.append([islem_tarihi, "Tali Acente Carisi", aktif_acente, f"Acente Payı İptal/Kesintisi - {aciklama}", akturk_kazanci, 0.0, "Aktürk Sigorta Kazancı", 1])
                    
                    doc.worksheet("Cari_Islemler").append_rows(yeni_satirlar, value_input_option='USER_ENTERED')
                    doc.worksheet("Musteriler").append_row([mus, tc, ilet, brut], value_input_option='USER_ENTERED')
                    st.cache_data.clear()
                st.success("Harika! İşlem başarıyla kaydedildi.")

    # ------------------------------------------
    # 5.2 CARİ & FİNANS 
    # ------------------------------------------
    elif menu == "💰 Cari & Finans":
        st.header("💰 Gelişmiş Finans Yönetimi")
        t1, t2 = st.tabs(["🏢 Tali Acente & Merkez Poliçeleri", "👤 Müşteri Hesapları"])
        df_pol = get_data("Policeler"); df_cari = get_data("Cari_Islemler")
        df_urunler = get_data("Ayarlar_Urunler"); df_acenteler = get_data("Ayarlar_Acenteler")
        
        urun_oranlari = dict(zip(df_urunler['Urun_Adi'], df_urunler['Komisyon_Orani'])) if not df_urunler.empty else {}
        acente_oranlari = dict(zip(df_acenteler['Acente_Adi'], df_acenteler['Tali_Oran'])) if not df_acenteler.empty else {}

        with t1:
            if not df_pol.empty and "Acente" in df_pol.columns:
                acenteler = df_pol[df_pol["Acente"] != "Aktürk Sigorta (Merkez)"]["Acente"].dropna().unique().tolist()
                acenteler = ["Aktürk Sigorta (Merkez)"] + acenteler 
                secilen_acente = st.selectbox("Hesaplarını Görmek İstediğiniz Acente:", ["Seçiniz..."] + acenteler)
                
                if secilen_acente != "Seçiniz...":
                    c_tarih1, c_tarih2 = st.columns(2)
                    ilk_tarih = c_tarih1.date_input("Başlangıç Tarihi", datetime.today().replace(month=1, day=1))
                    son_tarih = c_tarih2.date_input("Bitiş Tarihi", datetime.today())
                    st.divider()

                    acente_policeleri = df_pol[df_pol["Acente"] == secilen_acente].copy()
                    acente_policeleri['Tarih_Obj'] = pd.to_datetime(acente_policeleri['Tanzim Tarihi'], dayfirst=True, errors='coerce')
                    
                    mask_pol = (acente_policeleri['Tarih_Obj'].dt.date >= ilk_tarih) & (acente_policeleri['Tarih_Obj'].dt.date <= son_tarih)
                    filtrelenmis_policeler = acente_policeleri[mask_pol].copy()

                    tali_orani = float(sayiya_cevir(acente_oranlari.get(secilen_acente, 100.0 if secilen_acente == "Aktürk Sigorta (Merkez)" else 0.0)))
                    if tali_orani > 1: tali_orani /= 100 
                    
                    def komisyon_belirle(row):
                        k_manuel = sayiya_cevir(row.get("Şirket Komisyonu", 0))
                        if k_manuel != 0.0: return k_manuel
                        else:
                            u_or = float(sayiya_cevir(urun_oranlari.get(row.get("Sigorta Türü",""), 0.0)))
                            if u_or > 1: u_or /= 100
                            return float(sayiya_cevir(row.get("Net Prim", 0))) * u_or
                            
                    filtrelenmis_policeler["Net Prim"] = filtrelenmis_policeler["Net Prim"].apply(sayiya_cevir)
                    filtrelenmis_policeler["Brüt Prim"] = filtrelenmis_policeler["Brüt Prim"].apply(sayiya_cevir)
                    filtrelenmis_policeler["Şirket Komisyonu"] = filtrelenmis_policeler.apply(komisyon_belirle, axis=1)
                    filtrelenmis_policeler["Aktürk Sigorta Kazancı"] = filtrelenmis_policeler["Şirket Komisyonu"] * tali_orani
                    
                    m1, m2, m3 = st.columns(3)
                    m1.metric("Toplam Net Üretim", para_format(filtrelenmis_policeler['Net Prim'].sum()))
                    m2.metric("Şirket Komisyonu", para_format(filtrelenmis_policeler['Şirket Komisyonu'].sum()))
                    m3.metric("Aktürk Sigorta Kazancı", para_format(filtrelenmis_policeler['Aktürk Sigorta Kazancı'].sum()))
                    
                    st.markdown("#### ✏️ Hızlı Poliçe Düzenleme (Prim ve Komisyon Değişimi)")
                    st.info("💡 **Aşağıdaki tablodan hücrelere çift tıklayarak 'Net Prim', 'Brüt Prim' ve 'Şirket Komisyonu' rakamlarını anında değiştirebilirsiniz.**")
                    
                    edit_cols = ["Tanzim Tarihi", "Müşteri Adı Soyadı", "Plaka", "Sigorta Türü", "Net Prim", "Brüt Prim", "Şirket Komisyonu"]
                    df_to_edit = filtrelenmis_policeler[edit_cols].copy()
                    
                    edited_df = st.data_editor(
                        df_to_edit, 
                        use_container_width=True, 
                        key=f"edit_pol_{secilen_acente}",
                        disabled=["Tanzim Tarihi", "Müşteri Adı Soyadı", "Plaka", "Sigorta Türü"]
                    )
                    
                    if st.button("💾 Tablodaki Değişiklikleri Kaydet", type="primary"):
                        with st.spinner("Güncellemeler Google Sheets'e tekil paket olarak gönderiliyor..."):
                            ws_pol = client.open_by_key(SHEET_ID).worksheet("Policeler")
                            headers = ws_pol.row_values(1)
                            
                            if "Şirket Komisyonu" not in headers:
                                ws_pol.update_cell(1, len(headers)+1, "Şirket Komisyonu")
                                headers.append("Şirket Komisyonu")
                                
                            cells_to_update = []
                            for idx in df_to_edit.index:
                                old_row = df_to_edit.loc[idx]
                                new_row = edited_df.loc[idx]
                                if not old_row.equals(new_row):
                                    sheet_row = int(filtrelenmis_policeler.loc[idx, "Sheet_Row"])
                                    for col in ["Net Prim", "Brüt Prim", "Şirket Komisyonu"]:
                                        if old_row[col] != new_row[col]:
                                            col_idx = headers.index(col) + 1
                                            val = float(sayiya_cevir(new_row[col]))
                                            cells_to_update.append(gspread.Cell(row=sheet_row, col=col_idx, value=val))
                            
                            if cells_to_update:
                                ws_pol.update_cells(cells_to_update, value_input_option='USER_ENTERED')
                                st.success(f"🎉 Harika! Toplam {len(cells_to_update)} adet veri güncellendi.")
                                st.cache_data.clear(); st.rerun()
                            else: st.info("Herhangi bir değişiklik algılanmadı.")
                    
                    excel_indir(filtrelenmis_policeler[edit_cols + ["Aktürk Sigorta Kazancı"]], "Poliçe ve Kazanç Raporunu Excel İndir", f"{secilen_acente}_Raporu")
                    
                    if secilen_acente != "Aktürk Sigorta (Merkez)":
                        st.divider()
                        st.markdown(f"#### 💳 {secilen_acente} - Cari Hesap Ekstresi")
                        if not df_cari.empty and "Kisi_Kurum" in df_cari.columns:
                            a_cari = df_cari[(df_cari["Kisi_Kurum"] == secilen_acente) & (df_cari["Islem_Turu"] == "Tali Acente Carisi")].copy()
                        else: a_cari = pd.DataFrame()
                            
                        if not a_cari.empty:
                            a_cari["Borc"] = a_cari["Borc"].apply(sayiya_cevir); a_cari["Alacak"] = a_cari["Alacak"].apply(sayiya_cevir)
                            a_cari['Tarih_Obj'] = pd.to_datetime(a_cari['Tarih'], dayfirst=True, errors='coerce')
                            
                            genel_bakiye = a_cari["Borc"].sum() - a_cari["Alacak"].sum()
                            if genel_bakiye > 0: st.error(f"🚨 Acentenin Size Borcu (Alacağınız): {para_format(genel_bakiye)}")
                            elif genel_bakiye < 0: st.success(f"✅ Sizin Acenteye Borcunuz: {para_format(abs(genel_bakiye))}")
                            else: st.info("Hesaplar Dengede (0,00 TL)")
                            
                            mask_cari = (a_cari['Tarih_Obj'].dt.date >= ilk_tarih) & (a_cari['Tarih_Obj'].dt.date <= son_tarih)
                            filtrelenmis_cari = a_cari[mask_cari]
                            
                            df_ui_cari = df_gorsel_yap(filtrelenmis_cari[["Tarih", "Islem_Detayi", "Borc", "Alacak", "Odeme_Tipi"]], ["Borc", "Alacak"])
                            st.dataframe(df_ui_cari, use_container_width=True)
                            excel_indir(filtrelenmis_cari[["Tarih", "Islem_Detayi", "Borc", "Alacak", "Odeme_Tipi"]], "Cari Ekstreyi Excel İndir", f"{secilen_acente}_Ekstre")

                        st.divider()
                        with st.form("acente_odeme", clear_on_submit=True):
                            st.markdown(f"💸 **{secilen_acente} - Tahsilat / Ödeme Girişi**")
                            c1, c2, c3, c4 = st.columns(4)
                            o_tarih = c1.date_input("İşlem Tarihi").strftime("%d.%m.%Y")
                            islem_yonu = c2.selectbox("İşlem Türü", ["Acenteden Para Geldi (Tahsilat)", "Acenteye Para Gönderdim (Ödeme)"])
                            o_tutar = float(sayiya_cevir(c3.text_input("Tutar (Örn: 1500,50)", value="")))
                            o_detay = c4.text_input("Açıklama")
                            
                            if st.form_submit_button("İşlemi Kaydet"):
                                borc_yaz = o_tutar if islem_yonu == "Acenteye Para Gönderdim (Ödeme)" else 0.0
                                alacak_yaz = o_tutar if islem_yonu == "Acenteden Para Geldi (Tahsilat)" else 0.0
                                client.open_by_key(SHEET_ID).worksheet("Cari_Islemler").append_row([o_tarih, "Tali Acente Carisi", secilen_acente, o_detay, borc_yaz, alacak_yaz, "Nakit/Havale", 1], value_input_option='USER_ENTERED')
                                st.cache_data.clear(); st.rerun()

        with t2:
            if not df_cari.empty and "Kisi_Kurum" in df_cari.columns:
                musteriler = df_cari[df_cari["Islem_Turu"] == "Müşteri Carisi"]["Kisi_Kurum"].dropna().unique().tolist()
                secilen_musteri = st.selectbox("Müşteri Seçin:", ["Seçiniz..."] + sorted(musteriler))
                if secilen_musteri != "Seçiniz...":
                    m_cari = df_cari[(df_cari["Kisi_Kurum"] == secilen_musteri) & (df_cari["Islem_Turu"] == "Müşteri Carisi")].copy()
                    
                    c_mtarih1, c_mtarih2 = st.columns(2)
                    m_ilk = c_mtarih1.date_input("Müşteri Başlangıç", datetime.today().replace(month=1, day=1))
                    m_son = c_mtarih2.date_input("Müşteri Bitiş", datetime.today())
                    
                    m_cari["Borc"] = m_cari["Borc"].apply(sayiya_cevir); m_cari["Alacak"] = m_cari["Alacak"].apply(sayiya_cevir)
                    m_cari['Tarih_Obj'] = pd.to_datetime(m_cari['Tarih'], dayfirst=True, errors='coerce')
                    
                    bakiye = m_cari["Borc"].sum() - m_cari["Alacak"].sum()
                    if bakiye > 0: st.error(f"🚨 Müşterinin Toplam Borcu: {para_format(bakiye)}")
                    elif bakiye < 0: st.success(f"✅ Fazla Ödeme (Alacaklı): {para_format(abs(bakiye))}")
                    else: st.info("Borç Yok (0,00 TL)")
                    
                    mask_m = (m_cari['Tarih_Obj'].dt.date >= m_ilk) & (m_cari['Tarih_Obj'].dt.date <= m_son)
                    df_ui_mus = df_gorsel_yap(m_cari[mask_m][["Tarih", "Islem_Detayi", "Borc", "Alacak", "Odeme_Tipi"]], ["Borc", "Alacak"])
                    st.dataframe(df_ui_mus, use_container_width=True)
                    excel_indir(m_cari[mask_m][["Tarih", "Islem_Detayi", "Borc", "Alacak", "Odeme_Tipi"]], "Müşteri Ekstresini İndir", f"{secilen_musteri}_Ekstre")

                    st.divider()
                    with st.form("musteri_odeme", clear_on_submit=True):
                        st.markdown(f"💳 **{secilen_musteri} - Tahsilat / İade Girişi**")
                        c1, c2, c3, c4 = st.columns(4)
                        m_tarih = c1.date_input("Tarih").strftime("%d.%m.%Y")
                        m_yon = c2.selectbox("Tür", ["Müşteriden Para Geldi (Tahsilat)", "Müşteriye Para İade Edildi"])
                        m_tutar = float(sayiya_cevir(c3.text_input("Tutar (Örn: 1500,50)", value="")))
                        m_detay = c4.text_input("Açıklama")
                        if st.form_submit_button("İşlemi Kaydet"):
                            m_borc = m_tutar if m_yon == "Müşteriye Para İade Edildi" else 0.0
                            m_alc = m_tutar if m_yon == "Müşteriden Para Geldi (Tahsilat)" else 0.0
                            client.open_by_key(SHEET_ID).worksheet("Cari_Islemler").append_row([m_tarih, "Müşteri Carisi", secilen_musteri, m_detay, m_borc, m_alc, "Nakit/Havale", 1], value_input_option='USER_ENTERED')
                            st.cache_data.clear(); st.rerun()

    # ------------------------------------------
    # 5.3 YENİLEME TAKVİMİ 
    # ------------------------------------------
    elif menu == "📅 Yenileme Takvimi":
        st.header("📅 Özgür Yenileme Takibi")
        df_pol = get_data("Policeler")
        
        t_col1, t_col2 = st.columns(2)
        takvim_basla = t_col1.date_input("Takvim Başlangıç", datetime.today())
        takvim_bitis = t_col2.date_input("Takvim Bitiş", datetime.today() + timedelta(days=30))
        st.divider()

        if not df_pol.empty and "Bitiş Tarihi" in df_pol.columns:
            df_pol['Bit_Obj'] = pd.to_datetime(df_pol['Bitiş Tarihi'], dayfirst=True, errors='coerce')
            mask_takvim = (df_pol['Bit_Obj'].dt.date >= takvim_basla) & (df_pol['Bit_Obj'].dt.date <= takvim_bitis)
            takvim_temel = df_pol[mask_takvim].copy()

            takvim_temel = takvim_temel[~takvim_temel["Sigorta Şirketi"].str.contains("İPTAL", na=False)]
            takvim_temel = takvim_temel[takvim_temel["Başlangıç Tarihi"] != takvim_temel["Bitiş Tarihi"]]

            iptal_df = df_pol[df_pol["Sigorta Şirketi"].str.contains("İPTAL-SATIŞ", na=False)]
            iptal_plakalar = iptal_df[iptal_df["Plaka"] != ""]["Plaka"].unique()
            iptal_musteriler = iptal_df[iptal_df["Plaka"] == ""]["Müşteri Adı Soyadı"].unique()
            
            takvim_temel = takvim_temel[~takvim_temel["Plaka"].isin(iptal_plakalar)]
            takvim_temel = takvim_temel[~((takvim_temel["Plaka"] == "") & (takvim_temel["Müşteri Adı Soyadı"].isin(iptal_musteriler)))]

            takvim_temel["Kalan Gün"] = (takvim_temel['Bit_Obj'].dt.normalize() - pd.Timestamp.today().normalize()).dt.days
            takvim = takvim_temel.sort_values("Kalan Gün")
            
            if not takvim.empty:
                def uyari(gun):
                    if gun < 0: return "🚨 SÜRESİ DOLDU"
                    if gun <= 10: return "🔴 ACİL YENİLEME"
                    if gun <= 30: return "🟡 YAKLAŞIYOR"
                    return "🟢 SÜRESİ VAR"
                takvim["DURUM"] = takvim["Kalan Gün"].apply(uyari)
                
                gosterim = ["DURUM", "Kalan Gün", "Bitiş Tarihi", "Müşteri Adı Soyadı", "Plaka", "Sigorta Türü", "Sigorta Şirketi", "Telefon / E-mail"]
                if "PDF Linki" in takvim.columns: gosterim.append("PDF Linki")
                
                st.dataframe(takvim[gosterim], column_config=STIL_AYARLARI, use_container_width=True)
                excel_indir(takvim[gosterim], "Bu Takvimi Excel İndir", "Ozel_Yenileme_Takvimi")
            else: st.info("Bu tarihler arasında süresi dolacak poliçe bulunmuyor.")

    # ------------------------------------------
    # 5.4 ARAMA VE AYARLAR 
    # ------------------------------------------
    elif menu == "🔎 Genel Arama":
        st.header("🔎 Gelişmiş Arama")
        df_pol = get_data("Policeler")
        if not df_pol.empty:
            ara = st.text_input("🔍 Müşteri Adı veya Plaka Yazın:")
            if ara:
                ara_temiz = temiz_isim(ara)
                mask = df_pol['Müşteri Adı Soyadı'].str.contains(ara_temiz, na=False) | df_pol['Plaka'].str.contains(ara_temiz, na=False)
                sonuc = df_pol[mask]
                
                if not sonuc.empty:
                    df_ui_arama = df_gorsel_yap(sonuc, ["Net Prim", "Brüt Prim", "Şirket Komisyonu"])
                    goster_sutunlar = [c for c in df_ui_arama.columns if c != "Sheet_Row"]
                    st.dataframe(df_ui_arama[goster_sutunlar], column_config=STIL_AYARLARI, use_container_width=True)
                    excel_indir(sonuc, "Arama Sonuçlarını Excel İndir", f"Arama_Sonucu_{ara_temiz}")
                else: 
                    st.warning("Eşleşen tam kayıt bulunamadı.")
                    tum_isimler = df_pol['Müşteri Adı Soyadı'].dropna().unique().tolist()
                    tum_plakalar = df_pol['Plaka'].dropna().unique().tolist()
                    aranacak_liste = [str(x) for x in tum_isimler + tum_plakalar if str(x).strip() != ""]
                    
                    oneriler = difflib.get_close_matches(ara_temiz, aranacak_liste, n=3, cutoff=0.6)
                    if oneriler:
                        oneri_metni = ", ".join(oneriler)
                        st.info(f"💡 **Bunu mu demek istediniz:** {oneri_metni}")

    # ------------------------------------------
    # 🛠️ DÜZELTME VE SİLME MERKEZİ 
    # ------------------------------------------
    elif menu == "🛠️ Düzeltme & Silme":
        st.header("🛠️ Kayıt Düzeltme ve Silme Merkezi")
        t1, t2, t3 = st.tabs(["👤 Müşteri Bilgisi Düzelt", "🗑️ Poliçe Sil (Ve Carisini)", "🗑️ Serbest Cari Kaydı Sil"])
        
        with t1:
            st.markdown("### Tüm Sistemde Müşteri Bilgilerini Güncelle")
            st.info("💡 Müşterinin adını değiştirirseniz, sistem o kişinin eski adıyla kaydedilmiş **tüm poliçelerini ve cari hesaplarını** otomatik olarak yeni adına günceller.")
            df_mus = get_data("Musteriler")
            if not df_mus.empty:
                musteri_listesi = df_mus["Musteri_Adi"].dropna().unique().tolist()
                secilen_mus = st.selectbox("Düzeltilecek Müşteriyi Seçin:", ["Seçiniz..."] + sorted(musteri_listesi))
                
                if secilen_mus != "Seçiniz...":
                    mus_bilgi = df_mus[df_mus["Musteri_Adi"] == secilen_mus].iloc[0]
                    with st.form("mus_duzelt"):
                        yeni_isim = st.text_input("Müşteri Adı Soyadı", value=secilen_mus)
                        yeni_tc = st.text_input("TC / VKN", value=str(mus_bilgi.get("TC_VKN", "")))
                        yeni_tel = st.text_input("Telefon", value=str(mus_bilgi.get("Telefon", "")))
                        
                        if st.form_submit_button("💾 Güncelle ve Tüm Sisteme Uygula"):
                            with st.spinner("Müşteri bilgileri Poliçe, Cari ve Müşteri listelerinde güncelleniyor..."):
                                doc = client.open_by_key(SHEET_ID)
                                y_isim_temiz = temiz_isim(yeni_isim)
                                
                                ws_mus = doc.worksheet("Musteriler")
                                m_headers = ws_mus.row_values(1)
                                mus_satir = int(mus_bilgi["Sheet_Row"])
                                updates = []
                                if "Musteri_Adi" in m_headers: updates.append(gspread.Cell(row=mus_satir, col=m_headers.index("Musteri_Adi")+1, value=y_isim_temiz))
                                if "TC_VKN" in m_headers: updates.append(gspread.Cell(row=mus_satir, col=m_headers.index("TC_VKN")+1, value=yeni_tc))
                                if "Telefon" in m_headers: updates.append(gspread.Cell(row=mus_satir, col=m_headers.index("Telefon")+1, value=yeni_tel))
                                if updates: ws_mus.update_cells(updates)
                                
                                if y_isim_temiz != secilen_mus:
                                    df_pol = get_data("Policeler")
                                    ws_pol = doc.worksheet("Policeler")
                                    p_headers = ws_pol.row_values(1)
                                    p_updates = []
                                    if "Müşteri Adı Soyadı" in p_headers:
                                        p_col = p_headers.index("Müşteri Adı Soyadı") + 1
                                        for idx, row in df_pol[df_pol["Müşteri Adı Soyadı"] == secilen_mus].iterrows():
                                            p_updates.append(gspread.Cell(row=int(row["Sheet_Row"]), col=p_col, value=y_isim_temiz))
                                    if p_updates: ws_pol.update_cells(p_updates)
                                    
                                    df_cari = get_data("Cari_Islemler")
                                    ws_cari = doc.worksheet("Cari_Islemler")
                                    c_headers = ws_cari.row_values(1)
                                    c_updates = []
                                    if "Kisi_Kurum" in c_headers:
                                        c_col = c_headers.index("Kisi_Kurum") + 1
                                        for idx, row in df_cari[(df_cari["Kisi_Kurum"] == secilen_mus) & (df_cari["Islem_Turu"] == "Müşteri Carisi")].iterrows():
                                            c_updates.append(gspread.Cell(row=int(row["Sheet_Row"]), col=c_col, value=y_isim_temiz))
                                    if c_updates: ws_cari.update_cells(c_updates)
                            
                            st.success(f"Başarılı! Müşteri '{y_isim_temiz}' olarak tüm sistemde güncellendi.")
                            st.cache_data.clear()
                            st.rerun()

        with t2:
            st.markdown("### Hatalı Poliçeyi Sil")
            st.error("🚨 DİKKAT: Buradan bir poliçe sildiğinizde, o poliçenin müşteriye yansıyan borcu, (varsa) tali acenteye yansıyan hakediş kayıtları ve Google Drive'daki PDF dosyası kalıcı olarak SİLİNİR.")
            df_pol = get_data("Policeler")
            if not df_pol.empty:
                ara_pol = st.text_input("Silmek istediğiniz poliçeyi arayın (Plaka, Poliçe No, Müşteri):").upper()
                if ara_pol:
                    ara_pol_t = temiz_isim(ara_pol)
                    mask = df_pol['Müşteri Adı Soyadı'].str.contains(ara_pol_t, na=False) | df_pol['Plaka'].str.contains(ara_pol_t, na=False) | df_pol['Poliçe No'].str.contains(ara_pol_t, na=False)
                    sonuc = df_pol[mask]
                    if not sonuc.empty:
                        silinecek_secim = st.selectbox(
                            "Silinecek Poliçeyi Seçin:", 
                            sonuc.apply(lambda x: f"{x['Tanzim Tarihi']} | {x['Müşteri Adı Soyadı']} | {x['Plaka']} | {x['Sigorta Türü']} | Brüt: {x['Brüt Prim']} (Satır:{x['Sheet_Row']})", axis=1).tolist()
                        )
                        
                        if st.button("🚨 Seçili Poliçeyi TAMAMEN SİL", type="primary"):
                            with st.spinner("Poliçe, Cari Kayıtlar ve PDF Drive'dan siliniyor..."):
                                satir_no = int(re.search(r'\(Satır:(\d+)\)', silinecek_secim).group(1))
                                s_row = sonuc[sonuc["Sheet_Row"] == satir_no].iloc[0]
                                
                                pdf_linki = s_row.get("PDF Linki", "Yok")
                                drive_pdf_sil(pdf_linki)
                                
                                doc = client.open_by_key(SHEET_ID)
                                doc.worksheet("Policeler").delete_rows(satir_no)
                                
                                aciklama_koku = f"{s_row['Sigorta Şirketi']} - {s_row['Sigorta Türü']} - Plaka: {s_row['Plaka']}".replace(" (İPTAL-SATIŞ)","").replace(" (İPTAL-ZEYL)","")
                                df_cari = get_data("Cari_Islemler")
                                if not df_cari.empty:
                                    bagli_cariler = df_cari[df_cari["Islem_Detayi"].str.contains(aciklama_koku, regex=False, na=False)]
                                    if not bagli_cariler.empty:
                                        sil_satirlar = sorted(bagli_cariler["Sheet_Row"].astype(int).tolist(), reverse=True)
                                        ws_cari = doc.worksheet("Cari_Islemler")
                                        for c_satir in sil_satirlar:
                                            ws_cari.delete_rows(c_satir)
                                            
                                st.success("Poliçe, PDF dosyası ve bağlı cari işlemler kalıcı olarak silindi!")
                                st.cache_data.clear()
                                st.rerun()
                    else: st.warning("Eşleşen poliçe bulunamadı.")
                        
        with t3:
            st.markdown("### Serbest Cari Kaydı Sil")
            st.info("Manuel eklediğiniz Tahsilat, Ödeme veya İade gibi cari işlemlerini silmek içindir.")
            df_cari = get_data("Cari_Islemler")
            if not df_cari.empty:
                ara_cari = st.text_input("Silinecek Cari Kaydını Arayın (Müşteri, Tutar, Açıklama vb.):").upper()
                if ara_cari:
                    ara_c_t = temiz_isim(ara_cari)
                    mask_c = df_cari['Kisi_Kurum'].str.contains(ara_c_t, na=False) | df_cari['Islem_Detayi'].str.contains(ara_c_t, na=False) | df_cari['Borc'].astype(str).str.contains(ara_c_t, na=False) | df_cari['Alacak'].astype(str).str.contains(ara_c_t, na=False)
                    sonuc_c = df_cari[mask_c]
                    if not sonuc_c.empty:
                        sil_c_sec = st.selectbox(
                            "Silinecek İşlemi Seçin:", 
                            sonuc_c.apply(lambda x: f"{x['Tarih']} | {x['Kisi_Kurum']} | {x['Islem_Detayi']} | Borç: {x['Borc']} / Alacak: {x['Alacak']} (Satır:{x['Sheet_Row']})", axis=1).tolist()
                        )
                        if st.button("🚨 Bu Cari Kaydını SİL", type="primary"):
                            c_satir_no = int(re.search(r'\(Satır:(\d+)\)', sil_c_sec).group(1))
                            client.open_by_key(SHEET_ID).worksheet("Cari_Islemler").delete_rows(c_satir_no)
                            st.success("Cari kayıt başarıyla silindi!")
                            st.cache_data.clear()
                            st.rerun()
                    else: st.warning("Cari kayıt bulunamadı.")

    elif menu == "⚙️ Ayarlar":
        st.header("⚙️ Sistem Ayarları")
        t1, t2, t3 = st.tabs(["📊 Ürün Oranları", "🏢 Tali Acente Oranları", "🔄 Veri Senkronizasyonu"])
        
        with t1:
            st.info("Tablo üzerinde çift tıklayarak oranları güncelleyebilirsiniz.")
            df_urun = get_data("Ayarlar_Urunler")
            if df_urun.empty: df_urun = pd.DataFrame(columns=["Urun_Adi", "Komisyon_Orani"])
            edited_urun = st.data_editor(df_urun.drop(columns=["Sheet_Row"], errors='ignore'), num_rows="dynamic", use_container_width=True)
            if st.button("💾 Ürün Değişikliklerini Kaydet"):
                ws_urun = client.open_by_key(SHEET_ID).worksheet("Ayarlar_Urunler")
                ws_urun.clear(); edited_urun = edited_urun.fillna("")
                data_to_save = [edited_urun.columns.values.tolist()]
                for r in edited_urun.values.tolist(): data_to_save.append([float(sayiya_cevir(v)) if isinstance(v, (int, float, str)) and sayiya_cevir(v) != 0 else v for v in r])
                ws_urun.append_rows(data_to_save, value_input_option='USER_ENTERED')
                st.cache_data.clear(); st.rerun()

        with t2:
            st.info("Tablo üzerinden acente oranlarını güncelleyebilirsiniz.")
            df_acente = get_data("Ayarlar_Acenteler")
            if df_acente.empty: df_acente = pd.DataFrame(columns=["Acente_Adi", "Tali_Oran"])
            edited_acente = st.data_editor(df_acente.drop(columns=["Sheet_Row"], errors='ignore'), num_rows="dynamic", use_container_width=True)
            if st.button("💾 Acente Değişikliklerini Kaydet"):
                ws_acente = client.open_by_key(SHEET_ID).worksheet("Ayarlar_Acenteler")
                ws_acente.clear(); edited_acente = edited_acente.fillna("")
                data_to_save = [edited_acente.columns.values.tolist()]
                for r in edited_acente.values.tolist(): data_to_save.append([float(sayiya_cevir(v)) if isinstance(v, (int, float, str)) and sayiya_cevir(v) != 0 else v for v in r])
                ws_acente.append_rows(data_to_save, value_input_option='USER_ENTERED')
                st.cache_data.clear(); st.rerun()

        with t3:
            st.warning("Eksik kalan geçmiş poliçeleri Cari İşlemler tablosuna otomatik işler.")
            if st.button("🚀 Senkronize Et"):
                st.cache_data.clear() 
                with st.spinner("Geçmiş verileriniz taranıyor..."):
                    df_pol = get_data("Policeler"); df_cari = get_data("Cari_Islemler")
                    df_urun = get_data("Ayarlar_Urunler"); df_acente = get_data("Ayarlar_Acenteler"); df_mus = get_data("Musteriler")
                    urun_oranlari = dict(zip(df_urun['Urun_Adi'], df_urun['Komisyon_Orani'])) if not df_urun.empty else {}
                    acente_oranlari = dict(zip(df_acente['Acente_Adi'], df_acente['Tali_Oran'])) if not df_acente.empty else {}
                    doc = client.open_by_key(SHEET_ID); ws_cari = doc.worksheet("Cari_Islemler")
                    
                    mevcut_cariler = df_cari["Islem_Detayi"].tolist() if not df_cari.empty and "Islem_Detayi" in df_cari.columns else []
                    mevcut_musteriler = [temiz_isim(m) for m in (df_mus["Musteri_Adi"].tolist() if not df_mus.empty and "Musteri_Adi" in df_mus.columns else [])]

                    yeni_satirlar_cari = []; yeni_satirlar_mus = []
                    for index, row in df_pol.iterrows():
                        mus = temiz_isim(str(row.get("Müşteri Adı Soyadı", ""))); tc = str(row.get("TC / VKN", "")); ilet = str(row.get("Telefon / E-mail", ""))
                        sir = str(row.get("Sigorta Şirketi", "")); urn = str(row.get("Sigorta Türü", "")); plk = str(row.get("Plaka", ""))
                        net = float(sayiya_cevir(row.get("Net Prim", 0))); brut = float(sayiya_cevir(row.get("Brüt Prim", 0)))
                        kom = float(sayiya_cevir(row.get("Şirket Komisyonu", 0))); acn = str(row.get("Acente", "Aktürk Sigorta (Merkez)"))
                        islem_tarihi = tarih_formatla(row.get("Tanzim Tarihi", ""))

                        islem_notu = ""
                        if "İPTAL-SATIŞ" in sir: islem_notu = "SATIŞ/TAM İPTAL İADESİ - "
                        elif "İPTAL-ZEYL" in sir: islem_notu = "KISMİ İPTAL/ZEYL İADESİ - "
                        elif net < 0 or brut < 0: islem_notu = "İPTAL/İADE - "

                        aciklama = f"{islem_notu}{sir.replace(' (İPTAL-SATIŞ)','').replace(' (İPTAL-ZEYL)','')} - {urn} - Plaka: {plk}"

                        if aciklama not in mevcut_cariler and f"Acente Payı Kesintisi - {aciklama}" not in mevcut_cariler:
                            u_oran = float(sayiya_cevir(urun_oranlari.get(urn, 0.0)))
                            if u_oran > 1: u_oran /= 100
                            sirket_komisyonu = kom if kom != 0.0 else net * u_oran
                                
                            t_oran = float(sayiya_cevir(acente_oranlari.get(acn, 0.0)))
                            if t_oran > 1: t_oran /= 100
                            akturk_kazanci = float(sirket_komisyonu * t_oran)

                            yeni_satirlar_cari.append([islem_tarihi, "Müşteri Carisi", mus, aciklama, brut, 0.0, "Aktarılmış Kayıt", 1])
                            if acn != "Aktürk Sigorta (Merkez)":
                                yeni_satirlar_cari.append([islem_tarihi, "Tali Acente Carisi", acn, f"Acente Payı Kesintisi - {aciklama}", akturk_kazanci, 0.0, "Aktürk Sigorta Kazancı", 1])
                            mevcut_cariler.append(aciklama)

                        if mus and mus not in mevcut_musteriler:
                            yeni_satirlar_mus.append([mus, tc, ilet, brut]); mevcut_musteriler.append(mus) 

                    if yeni_satirlar_cari: ws_cari.append_rows(yeni_satirlar_cari, value_input_option='USER_ENTERED') 
                    if yeni_satirlar_mus: doc.worksheet("Musteriler").append_rows(yeni_satirlar_mus, value_input_option='USER_ENTERED')
                    
                    if yeni_satirlar_cari or yeni_satirlar_mus: st.success("🎉 Senkronizasyon başarıyla tamamlandı."); st.cache_data.clear()
                    else: st.info("Eksik kayıt bulunamadı.")

    elif menu == "🔍 Tüm Arşiv":
        st.header("📂 Tüm Poliçe Arşivi")
        df_pol = get_data("Policeler")
        if not df_pol.empty:
            df_ui_arsiv = df_gorsel_yap(df_pol, ["Net Prim", "Brüt Prim", "Şirket Komisyonu"])
            goster_sutunlar = [c for c in df_ui_arsiv.columns if c != "Sheet_Row"]
            st.dataframe(df_ui_arsiv[goster_sutunlar], column_config=STIL_AYARLARI, use_container_width=True)
            
            st.divider()
            excel_indir(df_pol, "Tüm Arşivi Excel Olarak İndir", "Tum_Police_Arsivi")

    st.sidebar.divider()
    if st.sidebar.button("🚪 Güvenli Çıkış Yap"):
        st.session_state["giris_yapildi"] = False; st.rerun()
