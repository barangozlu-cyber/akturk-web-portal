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

# ==========================================
# 1. TEMEL AYARLAR VE SABİTLER
# ==========================================
VERSIYON = "Aktürk CRM v6.22 - Manuel Komisyon ve Canlı Editör"
SHEET_ID = "19zBeYZMLjpMe5rx1d6p6TNwQjHGFfqAx-qVKVxDxh24"
JSON_FILE = "anahtar.json"
DRIVE_KLASOR_ID = "17wXJilHVDuHhDWS-POS4nr_RjUZnN7eL" 

PORTAL_KULLANICI = "admin"
PORTAL_SIFRE = "akturk2026"
GONDEREN_MAIL = "sistem@akturksigorta.net"
MAIL_SIFRE = "1994686baran" 

st.set_page_config(page_title="Aktürk Sigorta Portal", layout="wide", initial_sidebar_state="expanded")

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
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(GONDEREN_MAIL, MAIL_SIFRE)
        server.send_message(msg)
        server.quit()
        return True
    except: 
        return False

def temiz_isim(metin):
    if pd.isna(metin) or not metin: return ""
    metin = str(metin).strip().upper()
    metin = metin.replace('i', 'İ').replace('ı', 'I').replace('ğ', 'Ğ').replace('ü', 'Ü').replace('ş', 'Ş').replace('ö', 'Ö').replace('ç', 'Ç')
    return metin

# 🌟 TİTANYUM TL MOTORU 
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
        else:
            deger_str = deger_str.replace(',', '')
    elif ',' in deger_str:
        parts = deger_str.split(',')
        if len(parts) > 2 or (len(parts) == 2 and len(parts[1]) == 3):
            deger_str = deger_str.replace(',', '')
        else:
            deger_str = deger_str.replace(',', '.')
    elif '.' in deger_str:
        parts = deger_str.split('.')
        if len(parts) > 2 or (len(parts) == 2 and len(parts[1]) == 3):
            deger_str = deger_str.replace('.', '')
        
    try: return float(deger_str)
    except: return 0.0

# GÖRSEL TL FORMATI 
def para_format(deger):
    try:
        temiz_sayi = sayiya_cevir(deger)
        formatted = "{:,.2f}".format(temiz_sayi)
        return formatted.replace(",", "X").replace(".", ",").replace("X", ".") + " TL"
    except:
        return "0,00 TL"

def df_gorsel_yap(df, para_sutunlari):
    df_gorsel = df.copy()
    for col in para_sutunlari:
        if col in df_gorsel.columns:
            df_gorsel[col] = df_gorsel[col].apply(para_format)
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
            elif col == df_export.columns[0]:
                toplamlar[col] = "GENEL TOPLAM"
            else:
                toplamlar[col] = ""
        df_export = pd.concat([df_export, pd.DataFrame([toplamlar])], ignore_index=True)

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False)
    st.download_button(f"📥 {buton_metni}", out.getvalue(), f"{dosya_adi}.xlsx")

@st.cache_resource(show_spinner="Bağlantı Kuruluyor...")
def get_credentials():
    try:
        if "google_kasa" in st.secrets:
            creds_dict = json.loads(st.secrets["google_kasa"])
            return Credentials.from_service_account_info(creds_dict, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
    except Exception: pass
    return Credentials.from_service_account_file(JSON_FILE, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])

@st.cache_resource
def get_client(): return gspread.authorize(get_credentials())
def get_drive_service(): return build('drive', 'v3', credentials=get_credentials())

client = get_client()

# 🌟 HAM VERİ OKUYUCUSU (GSPREAD'İN VİRGÜL SİLMESİNİ ENGELLEYEN METOT)
@st.cache_data(ttl=5, show_spinner=False)
def get_data(sheet_name):
    try:
        ws = client.open_by_key(SHEET_ID).worksheet(sheet_name)
        raw_data = ws.get_all_values() 
        if not raw_data:
            return pd.DataFrame()
        
        headers = raw_data[0]
        max_len = len(headers)
        rows = []
        for r in raw_data[1:]:
            if len(r) < max_len:
                r = r + [""] * (max_len - len(r))
            elif len(r) > max_len:
                r = r[:max_len]
            rows.append(r)
            
        df = pd.DataFrame(rows, columns=headers)
        if not df.empty:
            # Satır ID'sini ekliyoruz ki Editör hangi satırı güncelleyeceğini bilsin
            df['Sheet_Row'] = df.index + 2 
            for col in ["Müşteri Adı Soyadı", "Kisi_Kurum", "Musteri_Adi"]:
                if col in df.columns: df[col] = df[col].apply(temiz_isim)
        return df
    except Exception as e: 
        return pd.DataFrame()

def drive_pdf_yukle(file_bytes, file_name):
    try:
        drive_service = get_drive_service()
        file_metadata = {'name': file_name, 'parents': [DRIVE_KLASOR_ID]}
        media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype='application/pdf', resumable=True)
        file = drive_service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink', supportsAllDrives=True).execute()
        drive_service.permissions().create(fileId=file.get('id'), body={'type': 'anyone', 'role': 'reader'}, supportsAllDrives=True).execute()
        return file.get('webViewLink')
    except Exception: return "Yok"

def klasik_analiz(metin):
    data = {"tanzim": "", "baslangic": "", "bitis": "", "musteri": "", "tc_vkn": "", "sirket": "", "urun": "", "p_no": "", "plaka": "", "net_prim": 0.0, "brut_prim": 0.0}
    metin_upper = metin.upper()
    if "ANKARA SİGORTA" in metin_upper: data["sirket"] = "Ankara Sigorta"
    elif "DOĞA SİGORTA" in metin_upper: data["sirket"] = "Doğa Sigorta"
    elif "ALLIANZ" in metin_upper: data["sirket"] = "Allianz Sigorta"
    tarihler = re.findall(r'\b\d{2}[\./-]\d{2}[\./-]\d{4}\b', metin)
    if len(tarihler) >= 3: data["tanzim"], data["baslangic"], data["bitis"] = [t.replace("/", ".") for t in tarihler[:3]]
    tc = re.search(r'\b[0-9]{10,11}\b', metin)
    if tc: data["tc_vkn"] = tc.group()
    plaka = re.search(r'\b[0-9]{2}\s*[A-Z]{1,3}\s*[0-9]{2,4}\b', metin)
    if plaka: data["plaka"] = plaka.group()
    return data

STIL_AYARLARI = {"PDF Linki": st.column_config.LinkColumn("📄 Belge", display_text="📥 PDF'İ AÇ")}

# ==========================================
# 3. GÖRSEL TASARIM (CSS)
# ==========================================
st.markdown("""
    <style>
    .stApp { background-color: #0E1117 !important; }
    h1, h2, h3 { color: #F5C518 !important; }
    .stButton>button { background: linear-gradient(135deg, #F5C518 0%, #D4A002 100%) !important; color: #161A22 !important; font-weight: 800; border-radius: 8px; transition: all 0.3s ease; border: none; }
    .stButton>button:hover { transform: translateY(-2px); box-shadow: 0 5px 15px rgba(245, 197, 24, 0.4); }
    .login-box { background-color: #161A22; padding: 50px; border-radius: 20px; border: 1px solid rgba(245, 197, 24, 0.2); text-align: center; max-width: 450px; margin: auto; margin-top: 5vh; box-shadow: 0 15px 40px rgba(0,0,0,0.6); }
    [data-testid="stSidebar"] { background-color: #161A22 !important; border-right: 1px solid #2D303E !important; }
    div[data-testid="stMetric"] { background-color: #1E232F; border-radius: 12px; padding: 15px 20px; border: 1px solid #2D303E; box-shadow: 0 4px 10px rgba(0,0,0,0.2); }
    div[data-testid="stMetricValue"] { color: #F5C518 !important; font-weight: bold; }
    div[data-testid="stMetricLabel"] { color: #A0AEC0 !important; font-size: 15px; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 4. GİRİŞ VE ÜYELİK SİSTEMİ 
# ==========================================
if not st.session_state["giris_yapildi"]:
    st.markdown("<div class='login-box'>", unsafe_allow_html=True)
    st.markdown("""
        <div style="background: linear-gradient(135deg, #F5C518 0%, #D4A002 100%); width: 90px; height: 90px; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin: 0 auto 20px auto; box-shadow: 0 4px 20px rgba(245, 197, 24, 0.4);">
            <span style="font-size: 50px; font-weight: 900; color: #161A22;">A</span>
        </div>
        <h2 style='margin-top:0px; margin-bottom: 5px;'>Aktürk Sigorta</h2>
        <p style='color: #A0AEC0; font-size: 16px; margin-bottom: 30px;'>Poliçe & Finans Takip Sistemi</p>
    """, unsafe_allow_html=True)
    
    tab1, tab2, tab3 = st.tabs(["🔐 Giriş Yap", "📝 Başvuru", "🔑 Şifremi Unuttum"])
    
    with tab1:
        u = st.text_input("Kullanıcı Adı")
        p = st.text_input("Şifre", type="password")
        if st.button("Sisteme Gir"):
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
                    st.error("⚠️ Hatalı kullanıcı adı veya şifre! Lütfen tekrar deneyin.")
    
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
                if mail_gonder(m, "Aktürk Portal Şifreniz", f"Güvenli giriş şifreniz: {s}"): 
                    st.success("Şifreniz e-posta adresinize gönderildi.")
                else: 
                    st.error("Mail gönderilemedi. Ayarları kontrol edin.")
            else:
                st.error("⚠️ Sistemde böyle bir e-posta bulunamadı.")
                
    st.markdown("</div>", unsafe_allow_html=True)

# ==========================================
# 5. ANA UYGULAMA
# ==========================================
else:
    st.sidebar.title(f"🛡️ Aktürk Sigorta")
    st.sidebar.caption(f"Kullanıcı: {st.session_state['kullanici_adi'].upper()}")
    st.sidebar.divider()
    
    menu = st.sidebar.radio("MENÜ", [
        "📥 Poliçe Girişi", 
        "💰 Cari & Finans", 
        "📅 Yenileme Takvimi", 
        "🔎 Genel Arama", 
        "⚙️ Ayarlar", 
        "🔍 Tüm Arşiv"
    ], on_change=ekran_temizle)

    # ------------------------------------------
    # 5.1 POLİÇE GİRİŞİ 
    # ------------------------------------------
    if menu == "📥 Poliçe Girişi":
        st.header("📥 Yeni İşlem Kaydı")
        
        df_urun = get_data("Ayarlar_Urunler")
        urun_listesi = df_urun["Urun_Adi"].tolist() if not df_urun.empty else ["Trafik", "Kasko", "Sağlık"]
        dict_urun = dict(zip(df_urun['Urun_Adi'], df_urun['Komisyon_Orani'])) if not df_urun.empty else {}
        
        df_acente = get_data("Ayarlar_Acenteler")
        acente_listesi = df_acente["Acente_Adi"].tolist() if not df_acente.empty else ["Aktürk Sigorta (Merkez)"]
        dict_acente = dict(zip(df_acente['Acente_Adi'], df_acente['Tali_Oran'])) if not df_acente.empty else {}
        acente_listesi.append("➕ YENİ TALİ ACENTE EKLE")

        file = st.file_uploader("PDF Poliçe Seçin (Otomatik Okuma)", type="pdf")
        p_data = {"tanzim":"","baslangic":"","bitis":"","musteri":"","plaka":"","tc_vkn":"","sirket":"","p_no":"","net_prim":0.0,"brut_prim":0.0}
        f_bytes = None
        
        if file:
            f_bytes = file.getvalue()
            with st.spinner("📄 PDF Okunuyor..."):
                with pdfplumber.open(io.BytesIO(f_bytes)) as pdf:
                    txt = pdf.pages[0].extract_text()
                    if txt: 
                        p_data.update(klasik_analiz(txt))
                        st.success("PDF Tarandı! Lütfen boşlukları doldurup kaydedin.")

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
            urn = c10.selectbox("Ürün Türü", urun_listesi)
            
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
                    
                    # 🌟 KOMİSYON KARARI: Kullanıcı değer girdiyse onu kullan, yoksa net primden hesapla
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
                    
                    # 🌟 DİNAMİK SÜTUN MİMARİSİ (Şirket Komisyonu Sütununu Otomatik Yaratır)
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
                    
                    bugun = datetime.now().strftime("%d.%m.%Y")
                    aciklama = f"{sir} - {urn} - Plaka: {plk}"
                    
                    yeni_satirlar = []
                    if islem_notu:
                        yeni_satirlar.append([islem_tarihi, "Müşteri Carisi", mus, f"{islem_notu} İADESİ - {aciklama}", brut, 0.0, odm, taksit])
                    else:
                        yeni_satirlar.append([islem_tarihi, "Müşteri Carisi", mus, aciklama, brut, 0.0, odm, taksit])
                    
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
                secilen_acente = st.selectbox("Hesaplarını/Poliçelerini Görmek İstediğiniz Acente:", ["Seçiniz..."] + acenteler)
                
                if secilen_acente != "Seçiniz...":
                    st.markdown("### 📅 Tarih Aralığı Seçimi")
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
                    
                    # 🌟 KOMİSYON OKUMA MANTIĞI: Manuel yazılmışsa onu, yoksa otomatik hesaplamayı alır
                    def komisyon_belirle(row):
                        k_manuel = sayiya_cevir(row.get("Şirket Komisyonu", 0))
                        if k_manuel != 0.0:
                            return k_manuel
                        else:
                            u_or = float(sayiya_cevir(urun_oranlari.get(row.get("Sigorta Türü",""), 0.0)))
                            if u_or > 1: u_or /= 100
                            return float(sayiya_cevir(row.get("Net Prim", 0))) * u_or
                            
                    filtrelenmis_policeler["Net Prim"] = filtrelenmis_policeler["Net Prim"].apply(sayiya_cevir)
                    filtrelenmis_policeler["Brüt Prim"] = filtrelenmis_policeler["Brüt Prim"].apply(sayiya_cevir)
                    filtrelenmis_policeler["Şirket Komisyonu"] = filtrelenmis_policeler.apply(komisyon_belirle, axis=1)
                    filtrelenmis_policeler["Aktürk Sigorta Kazancı"] = filtrelenmis_policeler["Şirket Komisyonu"] * tali_orani
                    
                    st.markdown(f"#### 📊 {secilen_acente} - Kazanç Özeti ({ilk_tarih.strftime('%d.%m.%Y')} - {son_tarih.strftime('%d.%m.%Y')})")
                    m1, m2, m3 = st.columns(3)
                    m1.metric("Toplam Net Üretim", para_format(filtrelenmis_policeler['Net Prim'].sum()))
                    m2.metric("Toplam Şirket Komisyonu", para_format(filtrelenmis_policeler['Şirket Komisyonu'].sum()))
                    m3.metric("Aktürk Sigorta Kazancı", para_format(filtrelenmis_policeler['Aktürk Sigorta Kazancı'].sum()))
                    
                    # 🌟 CANLI VERİ EDİTÖRÜ (GEÇMİŞ POLİÇELERİ DÜZENLEME)
                    st.markdown("#### ✏️ Hızlı Poliçe Düzenleme (Prim ve Komisyon Değişimi)")
                    st.info("💡 **Aşağıdaki tablodan hücrelere çift tıklayarak 'Net Prim', 'Brüt Prim' ve 'Şirket Komisyonu' rakamlarını anında değiştirebilirsiniz. İşiniz bitince aşağıdaki 'Kaydet' butonuna basmayı unutmayın.** *(Not: Cari yansımaları güncellemek isterseniz Ayarlar > Senkronize Et kullanın)*")
                    
                    edit_cols = ["Tanzim Tarihi", "Müşteri Adı Soyadı", "Plaka", "Sigorta Türü", "Net Prim", "Brüt Prim", "Şirket Komisyonu"]
                    # Dropdown vs bozmamak için sadece numeric kolonları düzenlenebilir yapalım
                    df_to_edit = filtrelenmis_policeler[edit_cols].copy()
                    
                    edited_df = st.data_editor(
                        df_to_edit, 
                        use_container_width=True, 
                        key=f"edit_pol_{secilen_acente}",
                        disabled=["Tanzim Tarihi", "Müşteri Adı Soyadı", "Plaka", "Sigorta Türü"] # Sadece prim/komisyon değiştirilebilsin
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
                                st.cache_data.clear()
                                st.rerun()
                            else:
                                st.info("Herhangi bir değişiklik algılanmadı.")
                    
                    excel_indir(filtrelenmis_policeler[edit_cols + ["Aktürk Sigorta Kazancı"]], "Poliçe ve Kazanç Raporunu Excel İndir", f"{secilen_acente}_Raporu")
                    
                    if secilen_acente != "Aktürk Sigorta (Merkez)":
                        st.divider()
                        st.markdown(f"#### 💳 {secilen_acente} - Cari Hesap Ekstresi")
                        if not df_cari.empty and "Kisi_Kurum" in df_cari.columns:
                            a_cari = df_cari[(df_cari["Kisi_Kurum"] == secilen_acente) & (df_cari["Islem_Turu"] == "Tali Acente Carisi")].copy()
                        else:
                            a_cari = pd.DataFrame()
                            
                        if not a_cari.empty:
                            a_cari["Borc"] = a_cari["Borc"].apply(sayiya_cevir)
                            a_cari["Alacak"] = a_cari["Alacak"].apply(sayiya_cevir)
                            a_cari['Tarih_Obj'] = pd.to_datetime(a_cari['Tarih'], dayfirst=True, errors='coerce')
                            
                            genel_bakiye = a_cari["Borc"].sum() - a_cari["Alacak"].sum()
                            if genel_bakiye > 0: st.error(f"🚨 Acentenin Size Borcu (Sizin Alacağınız): {para_format(genel_bakiye)}")
                            elif genel_bakiye < 0: st.success(f"✅ Sizin Acenteye Borcunuz: {para_format(abs(genel_bakiye))}")
                            else: st.info("Hesaplar Dengede (0,00 TL)")
                            
                            mask_cari = (a_cari['Tarih_Obj'].dt.date >= ilk_tarih) & (a_cari['Tarih_Obj'].dt.date <= son_tarih)
                            filtrelenmis_cari = a_cari[mask_cari]
                            
                            df_ui_cari = df_gorsel_yap(filtrelenmis_cari[["Tarih", "Islem_Detayi", "Borc", "Alacak", "Odeme_Tipi"]], ["Borc", "Alacak"])
                            st.dataframe(df_ui_cari, use_container_width=True)
                            
                            excel_indir(filtrelenmis_cari[["Tarih", "Islem_Detayi", "Borc", "Alacak", "Odeme_Tipi"]], "Cari Ekstreyi Excel İndir (Dip Toplamlı)", f"{secilen_acente}_Ekstre")

                        st.divider()
                        with st.form("acente_odeme", clear_on_submit=True):
                            st.markdown(f"💸 **{secilen_acente} - Tahsilat / Ödeme Girişi**")
                            c1, c2, c3, c4 = st.columns(4)
                            o_tarih = c1.date_input("İşlem Tarihi").strftime("%d.%m.%Y")
                            islem_yonu = c2.selectbox("İşlem Türü", ["Acenteden Para Geldi (Tahsilat)", "Acenteye Para Gönderdim (Ödeme)"])
                            o_tutar_girdi = c3.text_input("Tutar (Örn: 1500,50)", value="")
                            o_tutar = float(sayiya_cevir(o_tutar_girdi))
                            o_detay = c4.text_input("Açıklama")
                            
                            if st.form_submit_button("İşlemi Kaydet"):
                                borc_yaz = o_tutar if islem_yonu == "Acenteye Para Gönderdim (Ödeme)" else 0.0
                                alacak_yaz = o_tutar if islem_yonu == "Acenteden Para Geldi (Tahsilat)" else 0.0
                                client.open_by_key(SHEET_ID).worksheet("Cari_Islemler").append_row([o_tarih, "Tali Acente Carisi", secilen_acente, o_detay, borc_yaz, alacak_yaz, "Nakit/Havale", 1], value_input_option='USER_ENTERED')
                                st.cache_data.clear()
                                st.rerun()

        with t2:
            if not df_cari.empty and "Kisi_Kurum" in df_cari.columns:
                musteriler = df_cari[df_cari["Islem_Turu"] == "Müşteri Carisi"]["Kisi_Kurum"].dropna().unique().tolist()
                secilen_musteri = st.selectbox("Detayını Görmek İstediğiniz Müşteriyi Seçin:", ["Seçiniz..."] + sorted(musteriler))
                if secilen_musteri != "Seçiniz...":
                    m_cari = df_cari[(df_cari["Kisi_Kurum"] == secilen_musteri) & (df_cari["Islem_Turu"] == "Müşteri Carisi")].copy()
                    
                    c_mtarih1, c_mtarih2 = st.columns(2)
                    m_ilk = c_mtarih1.date_input("Müşteri Başlangıç Tarihi", datetime.today().replace(month=1, day=1))
                    m_son = c_mtarih2.date_input("Müşteri Bitiş Tarihi", datetime.today())
                    
                    m_cari["Borc"] = m_cari["Borc"].apply(sayiya_cevir)
                    m_cari["Alacak"] = m_cari["Alacak"].apply(sayiya_cevir)
                    m_cari['Tarih_Obj'] = pd.to_datetime(m_cari['Tarih'], dayfirst=True, errors='coerce')
                    
                    bakiye = m_cari["Borc"].sum() - m_cari["Alacak"].sum()
                    if bakiye > 0: st.error(f"🚨 Müşterinin Toplam Kalan Borcu: {para_format(bakiye)}")
                    elif bakiye < 0: st.success(f"✅ Fazla Ödeme (Müşteri Alacaklı): {para_format(abs(bakiye))}")
                    else: st.info("Borç Yok (0,00 TL)")
                    
                    mask_m = (m_cari['Tarih_Obj'].dt.date >= m_ilk) & (m_cari['Tarih_Obj'].dt.date <= m_son)
                    filtrelenmis_m = m_cari[mask_m]
                    
                    df_ui_mus = df_gorsel_yap(filtrelenmis_m[["Tarih", "Islem_Detayi", "Borc", "Alacak", "Odeme_Tipi"]], ["Borc", "Alacak"])
                    st.dataframe(df_ui_mus, use_container_width=True)
                    
                    excel_indir(filtrelenmis_m[["Tarih", "Islem_Detayi", "Borc", "Alacak", "Odeme_Tipi"]], "Müşteri Ekstresini Excel İndir (Dip Toplamlı)", f"{secilen_musteri}_Ekstre")

                    st.divider()
                    with st.form("musteri_odeme", clear_on_submit=True):
                        st.markdown(f"💳 **{secilen_musteri} - Tahsilat / İade Girişi**")
                        c1, c2, c3, c4 = st.columns(4)
                        m_tarih = c1.date_input("İşlem Tarihi").strftime("%d.%m.%Y")
                        m_yon = c2.selectbox("İşlem Türü", ["Müşteriden Para Geldi (Tahsilat)", "Müşteriye Para İade Edildi"])
                        m_tutar_girdi = c3.text_input("Tutar (Örn: 1500,50)", value="")
                        m_tutar = float(sayiya_cevir(m_tutar_girdi))
                        m_detay = c4.text_input("Açıklama")
                        if st.form_submit_button("İşlemi Kaydet"):
                            m_borc_yaz = m_tutar if m_yon == "Müşteriye Para İade Edildi" else 0.0
                            m_alacak_yaz = m_tutar if m_yon == "Müşteriden Para Geldi (Tahsilat)" else 0.0
                            client.open_by_key(SHEET_ID).worksheet("Cari_Islemler").append_row([m_tarih, "Müşteri Carisi", secilen_musteri, m_detay, m_borc_yaz, m_alacak_yaz, "Nakit/Havale", 1], value_input_option='USER_ENTERED')
                            st.cache_data.clear()
                            st.rerun()

    # ------------------------------------------
    # 5.3 YENİLEME TAKVİMİ 
    # ------------------------------------------
    elif menu == "📅 Yenileme Takvimi":
        st.header("📅 Özgür Yenileme ve Takip Merkezi")
        df_pol = get_data("Policeler")
        
        st.markdown("Hangi tarihler arasındaki poliçe bitişlerini görmek istiyorsunuz?")
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
            else:
                st.info(f"{takvim_basla.strftime('%d.%m.%Y')} ile {takvim_bitis.strftime('%d.%m.%Y')} tarihleri arasında süresi dolacak poliçe bulunmuyor.")
        else:
            st.info("Poliçe verisi bulunamadı.")

    # ------------------------------------------
    # 5.4 ARAMA VE AYARLAR 
    # ------------------------------------------
    elif menu == "🔎 Genel Arama":
        st.header("🔎 Gelişmiş Arama Motoru")
        df_pol = get_data("Policeler")
        if not df_pol.empty:
            ara = st.text_input("🔍 Müşteri Adı veya Plaka Yazın:").upper()
            if ara:
                mask = df_pol['Müşteri Adı Soyadı'].str.upper().str.contains(ara, na=False) | df_pol['Plaka'].str.upper().str.contains(ara, na=False)
                sonuc = df_pol[mask]
                if not sonuc.empty:
                    df_ui_arama = df_gorsel_yap(sonuc, ["Net Prim", "Brüt Prim", "Şirket Komisyonu"])
                    goster_sutunlar = [c for c in df_ui_arama.columns if c != "Sheet_Row"]
                    st.dataframe(df_ui_arama[goster_sutunlar], column_config=STIL_AYARLARI, use_container_width=True)
                else: st.warning("Eşleşen kayıt bulunamadı.")

    elif menu == "⚙️ Ayarlar":
        st.header("⚙️ Sistem Ayarları")
        t1, t2, t3 = st.tabs(["📊 Ürün Oranları", "🏢 Tali Acente Oranları", "🔄 Veri Senkronizasyonu"])
        
        with t1:
            st.info("💡 Tablo üzerinde çift tıklayarak isimleri ve oranları güncelleyebilir, en alta yeni satır ekleyebilir veya yandaki kutucuğu işaretleyip klavyeden 'Delete' tuşuyla silebilirsiniz.")
            df_urun = get_data("Ayarlar_Urunler")
            if df_urun.empty:
                df_urun = pd.DataFrame(columns=["Urun_Adi", "Komisyon_Orani"])
                
            edited_urun = st.data_editor(df_urun.drop(columns=["Sheet_Row"], errors='ignore'), num_rows="dynamic", use_container_width=True, key="urun_editor")
            if st.button("💾 Ürün Değişikliklerini Kaydet"):
                with st.spinner("Kaydediliyor..."):
                    ws_urun = client.open_by_key(SHEET_ID).worksheet("Ayarlar_Urunler")
                    ws_urun.clear()
                    edited_urun = edited_urun.fillna("")
                    data_to_save = [edited_urun.columns.values.tolist()]
                    for r in edited_urun.values.tolist(): data_to_save.append([float(sayiya_cevir(v)) if isinstance(v, (int, float, str)) and sayiya_cevir(v) != 0 else v for v in r])
                    ws_urun.append_rows(data_to_save, value_input_option='USER_ENTERED')
                    st.cache_data.clear()
                    st.success("Ürün oranları başarıyla güncellendi!")
                    st.rerun()

        with t2:
            st.info("💡 Tablo üzerinde çift tıklayarak acente isimlerini/oranlarını güncelleyebilir, en alta yeni acente ekleyebilir veya silebilirsiniz.")
            df_acente = get_data("Ayarlar_Acenteler")
            if df_acente.empty:
                df_acente = pd.DataFrame(columns=["Acente_Adi", "Tali_Oran"])
                
            edited_acente = st.data_editor(df_acente.drop(columns=["Sheet_Row"], errors='ignore'), num_rows="dynamic", use_container_width=True, key="acente_editor")
            if st.button("💾 Acente Değişikliklerini Kaydet"):
                with st.spinner("Kaydediliyor..."):
                    ws_acente = client.open_by_key(SHEET_ID).worksheet("Ayarlar_Acenteler")
                    ws_acente.clear()
                    edited_acente = edited_acente.fillna("")
                    data_to_save = [edited_acente.columns.values.tolist()]
                    for r in edited_acente.values.tolist(): data_to_save.append([float(sayiya_cevir(v)) if isinstance(v, (int, float, str)) and sayiya_cevir(v) != 0 else v for v in r])
                    ws_acente.append_rows(data_to_save, value_input_option='USER_ENTERED')
                    st.cache_data.clear()
                    st.success("Acente oranları başarıyla güncellendi!")
                    st.rerun()

        with t3:
            st.warning("Bu işlem 'Poliçeler' sayfanızdaki verileri okuyup, hem 'Cari İşlemler' hem de 'Müşteriler' tablosunda eksik olan kayıtları ilgili oldukları tarihe göre otomatik hesaplar ve ekler.")
            if st.button("🚀 Geçmiş Poliçeleri Cari Tabloya Senkronize Et"):
                st.cache_data.clear() 
                
                with st.spinner("Geçmiş verileriniz taranıyor, tarihler düzenleniyor... Lütfen bekleyin."):
                    df_pol = get_data("Policeler")
                    df_cari = get_data("Cari_Islemler")
                    df_urun = get_data("Ayarlar_Urunler")
                    df_acente = get_data("Ayarlar_Acenteler")
                    df_mus = get_data("Musteriler")

                    urun_oranlari = dict(zip(df_urun['Urun_Adi'], df_urun['Komisyon_Orani'])) if not df_urun.empty else {}
                    acente_oranlari = dict(zip(df_acente['Acente_Adi'], df_acente['Tali_Oran'])) if not df_acente.empty else {}

                    doc = client.open_by_key(SHEET_ID)
                    ws_cari = doc.worksheet("Cari_Islemler")
                    
                    mevcut_cariler = df_cari["Islem_Detayi"].tolist() if not df_cari.empty and "Islem_Detayi" in df_cari.columns else []
                    mevcut_musteriler = df_mus["Musteri_Adi"].tolist() if not df_mus.empty and "Musteri_Adi" in df_mus.columns else []
                    mevcut_musteriler = [temiz_isim(m) for m in mevcut_musteriler]

                    yeni_satirlar_cari = [] 
                    yeni_satirlar_mus = []
                    ornek_brut = 0.0

                    for index, row in df_pol.iterrows():
                        mus = temiz_isim(str(row.get("Müşteri Adı Soyadı", "")))
                        tc = str(row.get("TC / VKN", ""))
                        ilet = str(row.get("Telefon / E-mail", ""))
                        sir = str(row.get("Sigorta Şirketi", ""))
                        urn = str(row.get("Sigorta Türü", ""))
                        plk = str(row.get("Plaka", ""))
                        
                        net = float(sayiya_cevir(row.get("Net Prim", 0)))
                        brut = float(sayiya_cevir(row.get("Brüt Prim", 0)))
                        kom = float(sayiya_cevir(row.get("Şirket Komisyonu", 0)))
                        ornek_brut = brut 
                        acn = str(row.get("Acente", "Aktürk Sigorta (Merkez)"))

                        islem_tarihi = tarih_formatla(row.get("Tanzim Tarihi", ""))

                        islem_notu = ""
                        if "İPTAL-SATIŞ" in sir: islem_notu = "SATIŞ/TAM İPTAL İADESİ - "
                        elif "İPTAL-ZEYL" in sir: islem_notu = "KISMİ İPTAL/ZEYL İADESİ - "
                        elif net < 0 or brut < 0: islem_notu = "İPTAL/İADE - "

                        aciklama = f"{islem_notu}{sir.replace(' (İPTAL-SATIŞ)','').replace(' (İPTAL-ZEYL)','')} - {urn} - Plaka: {plk}"

                        if aciklama not in mevcut_cariler and f"Acente Payı Kesintisi - {aciklama}" not in mevcut_cariler:
                            u_oran = float(sayiya_cevir(urun_oranlari.get(urn, 0.0)))
                            if u_oran > 1: u_oran /= 100
                            
                            if kom != 0.0:
                                sirket_komisyonu = kom
                            else:
                                sirket_komisyonu = net * u_oran
                                
                            t_oran = float(sayiya_cevir(acente_oranlari.get(acn, 0.0)))
                            if t_oran > 1: t_oran /= 100
                            akturk_kazanci = float(sirket_komisyonu * t_oran)

                            yeni_satirlar_cari.append([islem_tarihi, "Müşteri Carisi", mus, aciklama, brut, 0.0, "Aktarılmış Kayıt", 1])
                            
                            if acn != "Aktürk Sigorta (Merkez)":
                                yeni_satirlar_cari.append([islem_tarihi, "Tali Acente Carisi", acn, f"Acente Payı Kesintisi - {aciklama}", akturk_kazanci, 0.0, "Aktürk Sigorta Kazancı", 1])
                            
                            mevcut_cariler.append(aciklama)

                        if mus and mus not in mevcut_musteriler:
                            yeni_satirlar_mus.append([mus, tc, ilet, brut])
                            mevcut_musteriler.append(mus) 

                    if yeni_satirlar_cari:
                        ws_cari.append_rows(yeni_satirlar_cari, value_input_option='USER_ENTERED') 
                    
                    if yeni_satirlar_mus:
                        doc.worksheet("Musteriler").append_rows(yeni_satirlar_mus, value_input_option='USER_ENTERED')

                    if yeni_satirlar_cari or yeni_satirlar_mus:
                        st.success(f"🎉 Harika! {len(yeni_satirlar_cari)} Cari işlem ve {len(yeni_satirlar_mus)} Yeni Müşteri başarıyla senkronize edildi. (Son okunan brüt: {ornek_brut})")
                        st.cache_data.clear()
                    else:
                        st.info("Tüm poliçeleriniz zaten Cari ve Müşteri sayfanızda işlenmiş durumda. Eksik kayıt bulunamadı.")

    elif menu == "🔍 Tüm Arşiv":
        st.header("📂 Tüm Poliçe Arşivi")
        df_pol = get_data("Policeler")
        if not df_pol.empty:
            df_ui_arsiv = df_gorsel_yap(df_pol, ["Net Prim", "Brüt Prim", "Şirket Komisyonu"])
            goster_sutunlar = [c for c in df_ui_arsiv.columns if c != "Sheet_Row"]
            st.dataframe(df_ui_arsiv[goster_sutunlar], column_config=STIL_AYARLARI, use_container_width=True)

    st.sidebar.divider()
    if st.sidebar.button("🚪 Güvenli Çıkış Yap"):
        st.session_state["giris_yapildi"] = False; st.rerun()
