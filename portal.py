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

# ==========================================
# 1. TEMEL AYARLAR VE SABİTLER
# ==========================================
VERSIYON = "Aktürk CRM v6.2 - İptal Modülü Entegreli"
SHEET_ID = "19zBeYZMLjpMe5rx1d6p6TNwQjHGFfqAx-qVKVxDxh24"
JSON_FILE = "anahtar.json"
DRIVE_KLASOR_ID = "17wXJilHVDuHhDWS-POS4nr_RjUZnN7eL" 

# Mail & Giriş Ayarları
PORTAL_KULLANICI = "baran"
PORTAL_SIFRE = "akturk2026"
GONDEREN_MAIL = "sistem@akturksigorta.net"
MAIL_SIFRE = "1994686baran" # Mutlaka doldurun

st.set_page_config(page_title="Aktürk Sigorta Portal", layout="wide", initial_sidebar_state="expanded")

# ==========================================
# 2. YARDIMCI FONKSİYONLAR
# ==========================================
if "giris_yapildi" not in st.session_state:
    st.session_state["giris_yapildi"] = False

def ekran_temizle():
    """Sayfa geçişlerinde hayalet kutuları tertemiz yapar."""
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

def sayiya_cevir(deger):
    if pd.isna(deger) or str(deger).strip() == "": return 0.0
    deger_str = str(deger).strip()
    if '.' in deger_str and ',' in deger_str: deger_str = deger_str.replace('.', '').replace(',', '.')
    elif ',' in deger_str: deger_str = deger_str.replace(',', '.')
    try: return float(deger_str)
    except:
        num = re.sub(r'[^\d.]', '', deger_str)
        try: return float(num) if num else 0.0
        except: return 0.0

@st.cache_resource(show_spinner="Bağlantı Kuruluyor...")
def get_credentials():
    try:
        # Önce buluttaki gizli kasaya (Streamlit Cloud) bakmayı dener
        if "google_kasa" in st.secrets:
            creds_dict = json.loads(st.secrets["google_kasa"])
            return Credentials.from_service_account_info(creds_dict, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
    except Exception:
        pass
        
    # Masaüstü (Yerel) kullanım için anahtar.json dosyasından okur
    return Credentials.from_service_account_file(JSON_FILE, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])

@st.cache_resource
def get_client(): return gspread.authorize(get_credentials())
def get_drive_service(): return build('drive', 'v3', credentials=get_credentials())

client = get_client()

@st.cache_data(ttl=5, show_spinner="Veriler Güncelleniyor...")
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
    except Exception as e: return "Yok"

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
    .stButton>button { background: linear-gradient(135deg, #F5C518 0%, #D4A002 100%) !important; color: black !important; font-weight: 800; border-radius: 8px; }
    .login-box { background-color: #161A22; padding: 40px; border-radius: 15px; border: 1px solid #2D303E; text-align: center; max-width: 450px; margin: auto; margin-top: 5vh; box-shadow: 0 10px 30px rgba(0,0,0,0.5); }
    [data-testid="stSidebar"] { background-color: #161A22 !important; border-right: 1px solid #2D303E !important; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 4. GİRİŞ VE ÜYELİK SİSTEMİ
# ==========================================
if not st.session_state["giris_yapildi"]:
    st.markdown("<div class='login-box'>", unsafe_allow_html=True)
    st.image("https://cdn-icons-png.flaticon.com/512/2600/2600102.png", width=70)
    st.markdown("<h2>🛡️ Aktürk Portal</h2>", unsafe_allow_html=True)
    
    tab1, tab2, tab3 = st.tabs(["🔐 Giriş Yap", "📝 Başvuru", "🔑 Şifremi Unuttum"])
    
    with tab1:
        u = st.text_input("Kullanıcı Adı")
        p = st.text_input("Şifre", type="password")
        if st.button("Sisteme Gir"):
            df_u = get_data("Kullanicilar")
            if (u == PORTAL_KULLANICI and p == PORTAL_SIFRE) or (not df_u.empty and u in df_u["Kullanici_Adi"].values and p in df_u[df_u["Kullanici_Adi"]==u]["Sifre"].values[0]):
                st.session_state["giris_yapildi"] = True
                st.session_state["kullanici_adi"] = u
                st.rerun()
            else: 
                st.error("Hatalı kullanıcı adı veya şifre!")
    
    with tab2:
        b_ad = st.text_input("Ad Soyad")
        b_mail = st.text_input("E-posta Adresiniz")
        if st.button("Başvuruyu Gönder"):
            if mail_gonder("baran@akturksigorta.net", "Yeni Acente/Personel Başvurusu", f"Başvuran: {b_ad}\nMail: {b_mail}"):
                st.success("Talebiniz yönetime iletildi.")
            else:
                st.error("Mail gönderilemedi. Ayarları kontrol edin.")

    with tab3:
        m = st.text_input("Sistemdeki Kayıtlı E-postanız")
        if st.button("Şifremi Mail At"):
            df_u = get_data("Kullanicilar")
            if not df_u.empty and m in df_u["E-posta"].values:
                s = df_u[df_u["E-posta"]==m]["Sifre"].values[0]
                if mail_gonder(m, "Aktürk Portal Şifreniz", f"Güvenli giriş şifreniz: {s}"):
                    st.success("Şifreniz e-posta adresinize gönderildi.")
            else:
                st.error("Sistemde böyle bir e-posta bulunamadı.")
    st.markdown("</div>", unsafe_allow_html=True)

# ==========================================
# 5. ANA UYGULAMA (GİRİŞ BAŞARILI)
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
    # 5.1 POLİÇE GİRİŞİ (Akıllı PDF Tarama)
    # ------------------------------------------
    if menu == "📥 Poliçe Girişi":
        st.header("📥 Yeni Poliçe Kaydı")
        
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
            st.subheader("1. Müşteri ve Poliçe Detayları")
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
            
            # --- YENİ İPTAL MODÜLÜ ---
            iptal_mi = st.checkbox("🚨 BU BİR İPTAL (İADE) İŞLEMİDİR (İşaretlerseniz hesaplara otomatik eksi bakiyeyle geçer)")
            
            c10, c11, c12 = st.columns(3)
            urn = c10.selectbox("Ürün Türü", urun_listesi)
            net = c11.number_input("Net Prim", value=float(p_data["net_prim"]))
            brut = c12.number_input("Brüt Prim", value=float(p_data["brut_prim"]))
            
            c13, c14 = st.columns(2)
            acn = c13.selectbox("İşlemi Yapan Acente", acente_listesi)
            
            yeni_acente_adi = ""
            yeni_acente_orani = 0.0
            if acn == "➕ YENİ TALİ ACENTE EKLE":
                yeni_acente_adi = c13.text_input("Yeni Acente Adını Yazın:")
                yeni_acente_orani = c13.number_input("Komisyon Oranı (Örn: 70)", value=0.0)
                
            odm = c14.selectbox("Tahsilat Durumu", ["Nakit Alındı", "Müşteri Kredi Kartı", "Havale", "Acenteye Borçlanıldı (Cari)"])
            taksit = c14.number_input("Taksit", min_value=1, value=1)
            adr = st.text_area("Adres", "")
            
            if st.form_submit_button("✅ POLİÇEYİ, CARİYİ VE PDF'İ KAYDET"):
                with st.spinner("Sisteme İşleniyor ve Buluta Yükleniyor..."):
                    doc = client.open_by_key(SHEET_ID)
                    aktif_acente = acn
                    mus = temiz_isim(mus_girdi)
                    
                    # İptal Durumu Kontrolü ve Hesaplamaları
                    if iptal_mi:
                        net = -abs(net)
                        brut = -abs(brut)
                        sir = f"{sir} (İPTAL)"
                    
                    link = drive_pdf_yukle(f_bytes, f"{mus}_{plk}_{sir}.pdf") if f_bytes else "Yok"
                    
                    u_oran = sayiya_cevir(dict_urun.get(urn, 0.0))
                    if u_oran > 1: u_oran /= 100 
                    
                    if acn == "➕ YENİ TALİ ACENTE EKLE" and yeni_acente_adi != "":
                        aktif_acente = yeni_acente_adi
                        doc.worksheet("Ayarlar_Acenteler").append_row([yeni_acente_adi, yeni_acente_orani])
                        t_oran = sayiya_cevir(yeni_acente_orani)
                    else:
                        t_oran = sayiya_cevir(dict_acente.get(aktif_acente, 0.0))
                    if t_oran > 1: t_oran /= 100 
                    
                    akturk_kazanci = net * u_oran * t_oran
                    
                    # 1. Poliçe Kaydı
                    doc.worksheet("Policeler").append_row([tan, bas, bit, mus, tc, sir, urn, pno, plk, net, brut, aktif_acente, adr, ilet, link])
                    
                    # 2. Cari Kayıtları
                    bugun = datetime.now().strftime("%d.%m.%Y")
                    aciklama = f"{sir} - {urn} - Plaka: {plk}"
                    
                    if iptal_mi:
                        doc.worksheet("Cari_Islemler").append_row([bugun, "Müşteri Carisi", mus, f"İPTAL İADESİ - {aciklama}", brut, 0, odm, taksit])
                    else:
                        doc.worksheet("Cari_Islemler").append_row([bugun, "Müşteri Carisi", mus, aciklama, brut, 0, odm, taksit])
                    
                    if aktif_acente != "Aktürk Sigorta (Merkez)":
                        doc.worksheet("Cari_Islemler").append_row([bugun, "Tali Acente Carisi", aktif_acente, f"Acente Payı Kesintisi - {aciklama}", akturk_kazanci, 0, "Aktürk Sigorta Kazancı", 1])
                    
                    # 3. Müşteri Veritabanı
                    doc.worksheet("Musteriler").append_row([mus, tc, ilet, brut])
                    st.cache_data.clear()
                st.success("Harika! Kayıtlar başarıyla oluşturuldu.")

    # ------------------------------------------
    # 5.2 CARİ & FİNANS (Gelişmiş Detay ve Filtre)
    # ------------------------------------------
    elif menu == "💰 Cari & Finans":
        st.header("💰 Gelişmiş Finans Yönetimi")
        t1, t2 = st.tabs(["🏢 Tali Acente Hesapları", "👤 Müşteri Hesapları"])
        df_pol = get_data("Policeler")
        df_cari = get_data("Cari_Islemler")
        df_urunler = get_data("Ayarlar_Urunler")
        df_acenteler = get_data("Ayarlar_Acenteler")
        
        urun_oranlari = dict(zip(df_urunler['Urun_Adi'], df_urunler['Komisyon_Orani'])) if not df_urunler.empty else {}
        acente_oranlari = dict(zip(df_acenteler['Acente_Adi'], df_acenteler['Tali_Oran'])) if not df_acenteler.empty else {}

        # --- TALİ ACENTE SEKME ---
        with t1:
            if not df_pol.empty and "Acente" in df_pol.columns:
                acenteler = df_pol[df_pol["Acente"] != "Aktürk Sigorta (Merkez)"]["Acente"].dropna().unique().tolist()
                secilen_acente = st.selectbox("Hesabını Görmek İstediğiniz Acenteyi Seçin:", ["Seçiniz..."] + acenteler)
                
                if secilen_acente != "Seçiniz...":
                    acente_policeleri = df_pol[df_pol["Acente"] == secilen_acente].copy()
                    tali_orani = sayiya_cevir(acente_oranlari.get(secilen_acente, 0.0))
                    if tali_orani > 1: tali_orani /= 100 
                    
                    def urun_orani_getir(urun_adi):
                        u_or = sayiya_cevir(urun_oranlari.get(urun_adi, 0.0))
                        return u_or / 100 if u_or > 1 else u_or
                        
                    acente_policeleri["Net Prim"] = acente_policeleri["Net Prim"].apply(sayiya_cevir)
                    acente_policeleri["Ürün Komisyonu"] = acente_policeleri["Sigorta Türü"].apply(urun_orani_getir)
                    acente_policeleri["Aktürk Sigorta Kazancı"] = acente_policeleri["Net Prim"] * acente_policeleri["Ürün Komisyonu"] * tali_orani
                    
                    st.write(f"📊 **{secilen_acente} - Poliçe Üretimi ve Kazanç Tablosu**")
                    st.dataframe(acente_policeleri[["Tanzim Tarihi", "Müşteri Adı Soyadı", "Plaka", "Net Prim", "Aktürk Sigorta Kazancı", "Sigorta Şirketi"]], use_container_width=True)
                    
                    a_cari = df_cari[(df_cari["Kisi_Kurum"] == secilen_acente) & (df_cari["Islem_Turu"] == "Tali Acente Carisi")].copy()
                    if not a_cari.empty:
                        c_tarih1, c_tarih2 = st.columns(2)
                        ilk_tarih = c_tarih1.date_input("Acente Başlangıç Tarihi", datetime.today().replace(day=1))
                        son_tarih = c_tarih2.date_input("Acente Bitiş Tarihi", datetime.today())
                        
                        a_cari["Borc"] = a_cari["Borc"].apply(sayiya_cevir)
                        a_cari["Alacak"] = a_cari["Alacak"].apply(sayiya_cevir)
                        a_cari['Tarih_Obj'] = pd.to_datetime(a_cari['Tarih'], format='%d.%m.%Y', errors='coerce')
                        
                        genel_bakiye = a_cari["Borc"].sum() - a_cari["Alacak"].sum()
                        if genel_bakiye > 0: st.error(f"🚨 Acentenin Size Borcu (Alacağınız): {genel_bakiye:,.2f} TL")
                        elif genel_bakiye < 0: st.success(f"✅ Sizin Acenteye Borcunuz: {abs(genel_bakiye):,.2f} TL")
                        else: st.info("Hesaplar Dengede (0.00 TL)")
                        
                        mask = (a_cari['Tarih_Obj'].dt.date >= ilk_tarih) & (a_cari['Tarih_Obj'].dt.date <= son_tarih)
                        filtrelenmis = a_cari[mask]
                        st.dataframe(filtrelenmis[["Tarih", "Islem_Detayi", "Borc", "Alacak", "Odeme_Tipi"]], use_container_width=True)
                        
                        buffer_acente = io.BytesIO()
                        with pd.ExcelWriter(buffer_acente, engine='openpyxl') as writer:
                            filtrelenmis[["Tarih", "Islem_Detayi", "Borc", "Alacak", "Odeme_Tipi"]].to_excel(writer, index=False, sheet_name='Ekstre')
                        st.download_button("📥 Tali Acente Ekstresini Excel İndir", buffer_acente.getvalue(), f"{secilen_acente}_Ekstre.xlsx")
                    
                    st.divider()
                    with st.form("acente_odeme", clear_on_submit=True):
                        st.markdown("💸 **Acenteden Tahsilat / Ödeme Girişi**")
                        c1, c2, c3 = st.columns(3)
                        o_tarih = c1.date_input("Tarih").strftime("%d.%m.%Y")
                        o_tutar = c2.number_input("Tutar (TL)", min_value=0.0)
                        o_detay = c3.text_input("Açıklama")
                        if st.form_submit_button("Tahsilatı İşle"):
                            client.open_by_key(SHEET_ID).worksheet("Cari_Islemler").append_row([o_tarih, "Tali Acente Carisi", secilen_acente, o_detay, 0, o_tutar, "Nakit/Havale", 1])
                            st.cache_data.clear()
                            st.rerun()

        # --- MÜŞTERİ SEKME ---
        with t2:
            if not df_cari.empty and "Kisi_Kurum" in df_cari.columns:
                musteriler = df_cari[df_cari["Islem_Turu"] == "Müşteri Carisi"]["Kisi_Kurum"].dropna().unique().tolist()
                secilen_musteri = st.selectbox("Detayını Görmek İstediğiniz Müşteriyi Seçin:", ["Seçiniz..."] + sorted(musteriler))
                
                if secilen_musteri != "Seçiniz...":
                    m_cari = df_cari[(df_cari["Kisi_Kurum"] == secilen_musteri) & (df_cari["Islem_Turu"] == "Müşteri Carisi")].copy()
                    
                    c_mtarih1, c_mtarih2 = st.columns(2)
                    m_ilk = c_mtarih1.date_input("Müşteri Başlangıç Tarihi", datetime.today().replace(day=1))
                    m_son = c_mtarih2.date_input("Müşteri Bitiş Tarihi", datetime.today())
                    
                    m_cari["Borc"] = m_cari["Borc"].apply(sayiya_cevir)
                    m_cari["Alacak"] = m_cari["Alacak"].apply(sayiya_cevir)
                    m_cari['Tarih_Obj'] = pd.to_datetime(m_cari['Tarih'], format='%d.%m.%Y', errors='coerce')
                    
                    bakiye = m_cari["Borc"].sum() - m_cari["Alacak"].sum()
                    if bakiye > 0: st.error(f"🚨 Müşterinin Kalan Borcu: {bakiye:,.2f} TL")
                    elif bakiye < 0: st.success(f"✅ Fazla Ödeme (Alacaklı): {abs(bakiye):,.2f} TL")
                    else: st.info("Borç Yok (0.00 TL)")
                    
                    mask_m = (m_cari['Tarih_Obj'].dt.date >= m_ilk) & (m_cari['Tarih_Obj'].dt.date <= m_son)
                    filtrelenmis_m = m_cari[mask_m]
                    st.dataframe(filtrelenmis_m[["Tarih", "Islem_Detayi", "Borc", "Alacak", "Odeme_Tipi"]], use_container_width=True)
                    
                    buffer_musteri = io.BytesIO()
                    with pd.ExcelWriter(buffer_musteri, engine='openpyxl') as writer:
                        filtrelenmis_m[["Tarih", "Islem_Detayi", "Borc", "Alacak", "Odeme_Tipi"]].to_excel(writer, index=False, sheet_name='Musteri_Ekstre')
                    st.download_button("📥 Müşteri Ekstresini Excel İndir", buffer_musteri.getvalue(), f"{secilen_musteri}_Ekstre.xlsx")
                    
                    st.divider()
                    with st.form("musteri_odeme", clear_on_submit=True):
                        st.markdown("💳 **Müşteriden Tahsilat Girişi**")
                        c1, c2, c3 = st.columns(3)
                        m_tarih = c1.date_input("Tahsilat Tarihi").strftime("%d.%m.%Y")
                        m_tutar = c2.number_input("Tutar (TL)", min_value=0.0)
                        m_detay = c3.text_input("Açıklama (Örn: Havale Geldi)")
                        if st.form_submit_button("Tahsilatı Kaydet"):
                            client.open_by_key(SHEET_ID).worksheet("Cari_Islemler").append_row([m_tarih, "Müşteri Carisi", secilen_musteri, m_detay, 0, m_tutar, "Nakit/Havale", 1])
                            st.cache_data.clear()
                            st.rerun()

    # ------------------------------------------
    # 5.3 YENİLEME TAKVİMİ
    # ------------------------------------------
    elif menu == "📅 Yenileme Takvimi":
        st.header("📅 Yenileme ve Takip Merkezi (Son 30 Gün)")
        df_pol = get_data("Policeler")
        if not df_pol.empty and "Bitiş Tarihi" in df_pol.columns:
            def kalan_gun(bitis_str):
                try:
                    match = re.search(r'\d{2}[\./-]\d{2}[\./-]\d{4}', str(bitis_str))
                    if match:
                        bit_date = pd.to_datetime(match.group().replace('/', '.').replace('-', '.'), format='%d.%m.%Y')
                        return (bit_date - pd.to_datetime(datetime.today().strftime('%d.%m.%Y'), format='%d.%m.%Y')).days
                    return None
                except: return None

            df_pol["Kalan Gün"] = df_pol["Bitiş Tarihi"].apply(kalan_gun)
            takvim = df_pol[(df_pol["Kalan Gün"].notna()) & (df_pol["Kalan Gün"] <= 30)].sort_values("Kalan Gün")
            
            if not takvim.empty:
                def uyari(gun):
                    if gun < 0: return "🚨 SÜRESİ DOLDU"
                    if gun <= 10: return "🔴 ACİL YENİLEME"
                    return "🟡 YAKLAŞIYOR"
                takvim["DURUM"] = takvim["Kalan Gün"].apply(uyari)
                
                gosterim = ["DURUM", "Kalan Gün", "Bitiş Tarihi", "Müşteri Adı Soyadı", "Plaka", "Sigorta Türü", "Sigorta Şirketi"]
                if "PDF Linki" in takvim.columns: gosterim.append("PDF Linki")
                
                st.dataframe(takvim[gosterim], column_config=STIL_AYARLARI, use_container_width=True)
                
                out_takvim = io.BytesIO()
                with pd.ExcelWriter(out_takvim, engine='openpyxl') as writer: takvim[gosterim].to_excel(writer, index=False)
                st.download_button("📥 Takvimi Excel İndir", out_takvim.getvalue(), "Yenileme_Takvimi.xlsx")
            else:
                st.info("Yakın tarihte süresi dolacak poliçe bulunmuyor.")

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
                    st.success(f"{len(sonuc)} kayıt bulundu.")
                    st.dataframe(sonuc, column_config=STIL_AYARLARI, use_container_width=True)
                else: st.warning("Eşleşen kayıt bulunamadı.")

    elif menu == "⚙️ Ayarlar":
        st.header("⚙️ Sistem Ayarları")
        t1, t2 = st.tabs(["📊 Ürün Komisyon Oranları", "🏢 Tali Acente Oranları"])
        with t1:
            st.dataframe(get_data("Ayarlar_Urunler"), use_container_width=True)
            with st.form("u_form", clear_on_submit=True):
                u_ad = st.text_input("Yeni Ürün Adı")
                u_oran = st.number_input("Şirket Komisyonu (%)", value=0.0)
                if st.form_submit_button("Ekle"):
                    client.open_by_key(SHEET_ID).worksheet("Ayarlar_Urunler").append_row([u_ad, u_oran]); st.cache_data.clear(); st.rerun()
        with t2:
            st.dataframe(get_data("Ayarlar_Acenteler"), use_container_width=True)
            with st.form("a_form", clear_on_submit=True):
                a_ad = st.text_input("Yeni Acente Adı")
                a_oran = st.number_input("Tali Oranı (%)", value=0.0)
                if st.form_submit_button("Ekle"):
                    client.open_by_key(SHEET_ID).worksheet("Ayarlar_Acenteler").append_row([a_ad, a_oran]); st.cache_data.clear(); st.rerun()

    # ------------------------------------------
    # 5.5 TÜM ARŞİV
    # ------------------------------------------
    elif menu == "🔍 Tüm Arşiv":
        st.header("📂 Tüm Poliçe Arşivi")
        df_pol = get_data("Policeler")
        if not df_pol.empty:
            st.dataframe(df_pol, column_config=STIL_AYARLARI, use_container_width=True)
            out_arsiv = io.BytesIO()
            with pd.ExcelWriter(out_arsiv, engine='openpyxl') as writer: df_pol.to_excel(writer, index=False)
            st.download_button("📥 Tüm Arşivi Excel İndir", out_arsiv.getvalue(), "Tum_Arsiv.xlsx")

    # ------------------------------------------
    # 5.6 GÜVENLİ ÇIKIŞ
    # ------------------------------------------
    st.sidebar.divider()
    if st.sidebar.button("🚪 Güvenli Çıkış Yap"):
        st.session_state["giris_yapildi"] = False
        st.rerun()
