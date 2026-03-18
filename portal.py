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
import json # Gizli kasa (Secrets) için eklendi

# --- 1. AYARLAR ---
VERSIYON = "CRM v4.6 - Güvenli Bulut & Web Entegre"
SHEET_ID = "19zBeYZMLjpMe5rx1d6p6TNwQjHGFfqAx-qVKVxDxh24"
JSON_FILE = "anahtar.json"
DRIVE_KLASOR_ID = "17wXJilHVDuHhDWS-POS4nr_RjUZnN7eL" 

# GİRİŞ BİLGİLERİ
PORTAL_KULLANICI = "baran"
PORTAL_SIFRE = "akturk2026"

st.set_page_config(page_title="Aktürk Sigorta Portal", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    .stApp { background-color: #0E1117 !important; }
    h1, h2, h3 { color: #F5C518 !important; font-weight: 800 !important; letter-spacing: -0.5px; }
    label, p, span { color: #E0E0E0 !important; font-weight: 500 !important; }
    input, select, textarea { 
        color: #FFFFFF !important; background-color: #1A1C23 !important; 
        border: 1px solid #333333 !important; border-radius: 6px !important;
        padding: 10px !important; font-weight: 600 !important;
    }
    input:focus, select:focus, textarea:focus { border: 1px solid #F5C518 !important; box-shadow: 0 0 5px rgba(245, 197, 24, 0.4) !important; }
    .stButton>button { 
        background: linear-gradient(135deg, #F5C518 0%, #D4A002 100%) !important; color: #000000 !important; 
        font-weight: 800 !important; height: 3.2em; width: 100%; border-radius: 8px; border: none !important;
        box-shadow: 0 4px 6px rgba(0,0,0,0.2) !important; transition: all 0.3s ease !important;
    }
    .stButton>button:hover { transform: translateY(-2px); box-shadow: 0 6px 12px rgba(245, 197, 24, 0.4) !important; background: linear-gradient(135deg, #FFD000 0%, #E6B000 100%) !important; }
    .login-box { background-color: #161A22; padding: 40px; border-radius: 15px; border: 1px solid #2D303E; box-shadow: 0 10px 30px rgba(0,0,0,0.5); max-width: 400px; margin: auto; margin-top: 10vh; text-align: center; }
    [data-testid="stSidebar"] { background-color: #161A22 !important; border-right: 1px solid #2D303E !important; }
    .stTabs [data-baseweb="tab-list"] { background-color: #1A1C23; border-radius: 8px; padding: 5px; box-shadow: inset 0 2px 4px rgba(0,0,0,0.3); }
    .stTabs [data-baseweb="tab"] { color: #A0A0A0 !important; font-weight: 700; border-radius: 6px; padding: 10px 20px; transition: 0.3s; }
    .stTabs [aria-selected="true"] { background-color: #2D303E !important; color: #F5C518 !important; box-shadow: 0 2px 5px rgba(0,0,0,0.4); }
    [data-testid="stStatusWidget"], [data-testid="stHeader"] {visibility: hidden !important; display: none !important;}
    </style>
    """, unsafe_allow_html=True)

if "giris_yapildi" not in st.session_state:
    st.session_state["giris_yapildi"] = False

def temiz_isim(metin):
    if pd.isna(metin) or not metin: return ""
    metin = str(metin).strip()
    metin = metin.replace('i', 'İ').replace('ı', 'I').replace('ğ', 'Ğ').replace('ü', 'Ü').replace('ş', 'Ş').replace('ö', 'Ö').replace('ç', 'Ç')
    return metin.upper()

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

# YENİ GÜÇLENDİRİLMİŞ KİMLİK DOĞRULAMA (SECRETS)
@st.cache_resource(show_spinner="🔐 Sunucuya Güvenli Bağlantı Kuruluyor...")
def get_credentials():
    # 1. Önce Streamlit Cloud'daki Gizli Kasaya (Secrets) bakar
    if "google_kasa" in st.secrets:
        creds_dict = json.loads(st.secrets["google_kasa"])
        return Credentials.from_service_account_info(creds_dict, scopes=[
            "https://www.googleapis.com/auth/spreadsheets", 
            "https://www.googleapis.com/auth/drive"
        ])
    # 2. Eğer bilgisayarda test ediliyorsa yerel dosyaya bakar
    else:
        return Credentials.from_service_account_file(JSON_FILE, scopes=[
            "https://www.googleapis.com/auth/spreadsheets", 
            "https://www.googleapis.com/auth/drive"
        ])

@st.cache_resource
def get_client(): return gspread.authorize(get_credentials())
def get_drive_service(): return build('drive', 'v3', credentials=get_credentials())

client = get_client()

@st.cache_data(ttl=5, show_spinner="🛡️ Veritabanı Güncelleniyor...")
def get_data(sheet_name):
    try:
        ws = client.open_by_key(SHEET_ID).worksheet(sheet_name)
        data = ws.get_all_records()
        df = pd.DataFrame(data) if data else pd.DataFrame()
        if not df.empty:
            if "Müşteri Adı Soyadı" in df.columns: df["Müşteri Adı Soyadı"] = df["Müşteri Adı Soyadı"].apply(temiz_isim)
            if "Kisi_Kurum" in df.columns: df["Kisi_Kurum"] = df["Kisi_Kurum"].apply(temiz_isim)
        return df
    except: return pd.DataFrame()

def drive_pdf_yukle(file_bytes, file_name):
    try:
        drive_service = get_drive_service()
        file_metadata = {'name': file_name, 'parents': [DRIVE_KLASOR_ID]}
        media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype='application/pdf', resumable=True)
        file = drive_service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink', supportsAllDrives=True).execute()
        try: drive_service.permissions().create(fileId=file.get('id'), body={'type': 'anyone', 'role': 'reader'}, supportsAllDrives=True).execute()
        except: pass
        return file.get('webViewLink')
    except Exception as e: return f"Hata: {e}"

def klasik_analiz(metin):
    data = {"tanzim": "", "baslangic": "", "bitis": "", "musteri": "", "tc_vkn": "", "sirket": "", "urun": "", "p_no": "", "plaka": "", "net_prim": 0.0, "brut_prim": 0.0}
    metin_upper = metin.upper()
    if "ANKARA SİGORTA" in metin_upper or "ANKARA ANONİM" in metin_upper: data["sirket"] = "Ankara Sigorta"
    elif "DOĞA SİGORTA" in metin_upper: data["sirket"] = "Doğa Sigorta"
    elif "ALLIANZ" in metin_upper or "ALLİANZ" in metin_upper: data["sirket"] = "Allianz Sigorta"
    tarihler = re.findall(r'\b\d{2}[\./-]\d{2}[\./-]\d{4}\b', metin)
    if len(tarihler) >= 1: data["tanzim"] = tarihler[0].replace("/", ".")
    if len(tarihler) >= 2: data["baslangic"] = tarihler[1].replace("/", ".")
    if len(tarihler) >= 3: data["bitis"] = tarihler[2].replace("/", ".")
    tc = re.search(r'\b[0-9]{10,11}\b', metin)
    if tc: data["tc_vkn"] = tc.group()
    plaka = re.search(r'\b[0-9]{2}\s*[A-Z]{1,3}\s*[0-9]{2,4}\b', metin)
    if plaka: data["plaka"] = plaka.group()
    data["musteri"] = temiz_isim(data.get("musteri", ""))
    return data

STIL_AYARLARI = {
    "PDF Linki": st.column_config.LinkColumn("📄 PDF Belgesi", help="Müşterinin poliçesini görüntülemek için tıklayın", display_text="📥 PDF'İ AÇ / İNDİR")
}

# ==========================================
# 🔐 MODÜL 0: GÜVENLİ GİRİŞ (LOGIN) EKRANI
# ==========================================
if not st.session_state["giris_yapildi"]:
    st.markdown("<div class='login-box'>", unsafe_allow_html=True)
    st.image("https://cdn-icons-png.flaticon.com/512/2600/2600102.png", width=80) 
    st.markdown("<h2>🛡️ Aktürk Portal</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color:#888; font-size:14px;'>Sisteme erişmek için yetki doğrulayın.</p>", unsafe_allow_html=True)
    
    with st.form("login_form"):
        k_adi = st.text_input("Kullanıcı Adı")
        sif = st.text_input("Şifre", type="password")
        if st.form_submit_button("Sisteme Giriş Yap"):
            if k_adi == PORTAL_KULLANICI and sif == PORTAL_SIFRE:
                st.session_state["giris_yapildi"] = True
                st.rerun()
            else:
                st.error("❌ Hatalı kullanıcı adı veya şifre!")
    st.markdown("</div>", unsafe_allow_html=True)

# ==========================================
# 🚀 ANA UYGULAMA
# ==========================================
else:
    st.sidebar.title("🛡️ AKTÜRK SİGORTA")
    st.sidebar.caption(VERSIYON)
    menu = st.sidebar.radio("İŞLEMLER", ["📥 Poliçe Girişi", "💰 Cari & Finans", "📅 Yenileme Takvimi", "🔎 Genel Arama", "⚙️ Ayarlar", "🔍 Tüm Arşiv"])
    
    st.sidebar.divider()
    if st.sidebar.button("🚪 Güvenli Çıkış Yap"):
        st.session_state["giris_yapildi"] = False
        st.rerun()

    if menu == "📥 Poliçe Girişi":
        st.header("📥 Poliçe ve Finans Girişi")
        
        df_urun = get_data("Ayarlar_Urunler")
        urun_listesi = df_urun["Urun_Adi"].tolist() if not df_urun.empty else ["Trafik", "Kasko", "Sağlık", "DASK"]
        dict_urun = dict(zip(df_urun['Urun_Adi'], df_urun['Komisyon_Orani'])) if not df_urun.empty else {}
        
        df_acente = get_data("Ayarlar_Acenteler")
        acente_listesi = df_acente["Acente_Adi"].tolist() if not df_acente.empty else ["Aktürk Sigorta (Merkez)"]
        dict_acente = dict(zip(df_acente['Acente_Adi'], df_acente['Tali_Oran'])) if not df_acente.empty else {}
        acente_listesi.append("➕ YENİ TALİ ACENTE EKLE")

        file = st.file_uploader("PDF Poliçe Seçin", type="pdf")
        data = {"net_prim": 0.0, "brut_prim": 0.0}
        file_bytes = None
        
        if file:
            file_bytes = file.getvalue()
            with st.spinner("📄 PDF Taranıyor..."):
                with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                    txt = pdf.pages[0].extract_text()
                    if txt:
                        data.update(klasik_analiz(txt))
                        st.success("PDF Tarandı! Lütfen eksikleri tamamlayıp kaydedin.")

        with st.form("police_formu", clear_on_submit=True): 
            st.subheader("1. Poliçe Bilgileri")
            c1, c2, c3 = st.columns(3)
            tan = c1.text_input("Tanzim Tarihi (A)", data.get("tanzim", ""))
            bas = c1.text_input("Başlangıç (B)", data.get("baslangic", ""))
            bit = c1.text_input("Bitiş (C)", data.get("bitis", ""))
            
            c4, c5, c6 = st.columns(3)
            mus_girdi = c4.text_input("Müşteri Adı Soyadı (D)", data.get("musteri", ""))
            tc = c5.text_input("TC / VKN (E)", data.get("tc_vkn", ""))
            ilet = c6.text_input("Telefon / E-mail (N)", "")
            
            c7, c8, c9 = st.columns(3)
            sir = c7.text_input("Sigorta Şirketi (F)", data.get("sirket", ""))
            pno = c8.text_input("Poliçe No (H)", data.get("p_no", ""))
            plk = c9.text_input("Plaka (I)", data.get("plaka", ""))
            
            st.subheader("2. Finans ve Komisyon")
            c10, c11, c12 = st.columns(3)
            urn = c10.selectbox("Sigorta Türü (G)", urun_listesi)
            net = c11.number_input("Net Prim (J)", value=float(data.get("net_prim", 0.0)))
            brut = c12.number_input("Brüt Prim (K)", value=float(data.get("brut_prim", 0.0)))
            
            st.subheader("3. Acente ve Tahsilat")
            c13, c14 = st.columns(2)
            acn = c13.selectbox("Acente (L)", acente_listesi)
            
            yeni_acente_adi = ""
            yeni_acente_orani = 0.0
            if acn == "➕ YENİ TALİ ACENTE EKLE":
                yeni_acente_adi = c13.text_input("Yeni Acente Adını Yazın:")
                yeni_acente_orani = c13.number_input("Bu acentenin komisyon oranı nedir?", value=0.00)
                
            odm = c14.selectbox("Tahsilat / Ödeme Durumu", [
                "Nakit Alındı", "Müşteri Kredi Kartı (Tek Çekim)", "Müşteri Kredi Kartı (Taksitli)", 
                "Müşteri Havale Gönderdi", "KENDİ KARTIMDAN ÇEKTİM - Nakit Alındı",
                "KENDİ KARTIMDAN ÇEKTİM - Havale Bekleniyor", "Acenteye Borçlanıldı (Cari)"
            ])
            taksit = c14.number_input("Taksit Sayısı (Yoksa 1)", min_value=1, value=1)
            adr = st.text_area("Adres (M)", "")
            
            if st.form_submit_button("✅ POLİÇEYİ, CARİYİ VE PDF'İ KAYDET"):
                with st.spinner("💾 PDF Drive'a Yükleniyor ve Veriler İşleniyor..."):
                    doc = client.open_by_key(SHEET_ID)
                    aktif_acente = acn
                    mus = temiz_isim(mus_girdi)
                    
                    pdf_linki = "Yok"
                    if file_bytes:
                        dosya_adi = f"{mus}_{plk}_{sir}.pdf".replace(" ", "_")
                        pdf_linki = drive_pdf_yukle(file_bytes, dosya_adi)
                    
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
                    
                    doc.worksheet("Policeler").append_row([tan, bas, bit, mus, tc, sir, urn, pno, plk, net, brut, aktif_acente, adr, ilet, pdf_linki])
                    bugun = datetime.now().strftime("%d.%m.%Y")
                    aciklama = f"{sir} - {urn} - Plaka: {plk}"
                    doc.worksheet("Cari_Islemler").append_row([bugun, "Müşteri Carisi", mus, aciklama + " - Poliçe Kesimi", brut, 0, odm, taksit])
                    
                    if aktif_acente != "Aktürk Sigorta (Merkez)":
                        doc.worksheet("Cari_Islemler").append_row([bugun, "Tali Acente Carisi", aktif_acente, aciklama + " - Acente Payı Kesintisi", akturk_kazanci, 0, "Aktürk Sigorta Kazancı", 1])
                    
                    doc.worksheet("Musteriler").append_row([mus, tc, ilet, brut])
                    st.cache_data.clear()
                st.success(f"Harika! Poliçe kaydedildi ve PDF Drive'a yedeklendi.")

    elif menu == "💰 Cari & Finans":
        st.header("💰 Gelişmiş Cari Hesap Yönetimi")
        t1, t2 = st.tabs(["🏢 Tali Acente Carisi", "👤 Müşteri Carisi"])
        
        df_pol = get_data("Policeler")
        df_cari = get_data("Cari_Islemler")
        df_urunler = get_data("Ayarlar_Urunler")
        df_acenteler = get_data("Ayarlar_Acenteler")
        
        urun_oranlari = dict(zip(df_urunler['Urun_Adi'], df_urunler['Komisyon_Orani'])) if not df_urunler.empty else {}
        acente_oranlari = dict(zip(df_acenteler['Acente_Adi'], df_acenteler['Tali_Oran'])) if not df_acenteler.empty else {}

        with t1:
            if not df_pol.empty and "Acente" in df_pol.columns:
                acenteler = df_pol[df_pol["Acente"] != "Aktürk Sigorta (Merkez)"]["Acente"].dropna().unique().tolist()
                if acenteler:
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
                        st.dataframe(acente_policeleri[["Tanzim Tarihi", "Müşteri Adı Soyadı", "Plaka", "Net Prim", "Aktürk Sigorta Kazancı"]], use_container_width=True)
                        
                        a_cari = df_cari[(df_cari["Kisi_Kurum"] == secilen_acente) & (df_cari["Islem_Turu"] == "Tali Acente Carisi")].copy()
                        if not a_cari.empty:
                            c_tarih1, c_tarih2 = st.columns(2)
                            ilk_tarih = c_tarih1.date_input("Başlangıç Tarihi Seçin", datetime.today().replace(day=1))
                            son_tarih = c_tarih2.date_input("Bitiş Tarihi Seçin", datetime.today())
                            a_cari["Borc"] = a_cari["Borc"].apply(sayiya_cevir)
                            a_cari["Alacak"] = a_cari["Alacak"].apply(sayiya_cevir)
                            a_cari['Tarih_Obj'] = pd.to_datetime(a_cari['Tarih'], format='%d.%m.%Y', errors='coerce')
                            genel_bakiye = a_cari["Borc"].sum() - a_cari["Alacak"].sum()
                            if genel_bakiye > 0: st.error(f"🚨 Toplam Borç: {genel_bakiye:,.2f} TL")
                            elif genel_bakiye < 0: st.success(f"✅ Sizin Borcunuz: {abs(genel_bakiye):,.2f} TL")
                            mask = (a_cari['Tarih_Obj'].dt.date >= ilk_tarih) & (a_cari['Tarih_Obj'].dt.date <= son_tarih)
                            filtrelenmis = a_cari[mask]
                            st.dataframe(filtrelenmis[["Tarih", "Islem_Detayi", "Borc", "Alacak", "Odeme_Tipi"]], use_container_width=True)
                            
                            buffer_acente = io.BytesIO()
                            with pd.ExcelWriter(buffer_acente, engine='openpyxl') as writer:
                                filtrelenmis[["Tarih", "Islem_Detayi", "Borc", "Alacak", "Odeme_Tipi"]].to_excel(writer, index=False, sheet_name='Acente_Ekstre')
                            st.download_button(label="📥 Tali Acente Ekstresini Excel İndir (.xlsx)", data=buffer_acente.getvalue(), file_name=f"{secilen_acente}_Ekstre.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        st.divider()
                        with st.form("acente_odeme", clear_on_submit=True):
                            c1, c2, c3 = st.columns(3)
                            o_tarih = c1.date_input("Tarih").strftime("%d.%m.%Y")
                            o_tutar = c2.number_input("Tutar (TL)", min_value=0.0)
                            o_detay = c3.text_input("Açıklama")
                            if st.form_submit_button("Tahsilatı İşle"):
                                with st.spinner("💸 Ödeme Hesaba İşleniyor..."):
                                    client.open_by_key(SHEET_ID).worksheet("Cari_Islemler").append_row([o_tarih, "Tali Acente Carisi", secilen_acente, o_detay, 0, o_tutar, "Nakit/Havale", 1])
                                    st.cache_data.clear()
                                    st.success("Tahsilat eklendi!")

        with t2:
            if not df_cari.empty and "Kisi_Kurum" in df_cari.columns:
                musteriler = df_cari[df_cari["Islem_Turu"] == "Müşteri Carisi"]["Kisi_Kurum"].dropna().unique().tolist()
                if musteriler:
                    secilen_musteri = st.selectbox("Müşteri Seçin:", ["Seçiniz..."] + sorted(musteriler))
                    if secilen_musteri != "Seçiniz...":
                        m_cari = df_cari[(df_cari["Kisi_Kurum"] == secilen_musteri) & (df_cari["Islem_Turu"] == "Müşteri Carisi")].copy()
                        c_mtarih1, c_mtarih2 = st.columns(2)
                        m_ilk = c_mtarih1.date_input("Başlangıç", datetime.today().replace(day=1))
                        m_son = c_mtarih2.date_input("Bitiş", datetime.today())
                        m_cari["Borc"] = m_cari["Borc"].apply(sayiya_cevir)
                        m_cari["Alacak"] = m_cari["Alacak"].apply(sayiya_cevir)
                        m_cari['Tarih_Obj'] = pd.to_datetime(m_cari['Tarih'], format='%d.%m.%Y', errors='coerce')
                        bakiye = m_cari["Borc"].sum() - m_cari["Alacak"].sum()
                        if bakiye > 0: st.error(f"🚨 Borç: {bakiye:,.2f} TL")
                        elif bakiye < 0: st.success(f"✅ Alacaklı: {abs(bakiye):,.2f} TL")
                        mask_m = (m_cari['Tarih_Obj'].dt.date >= m_ilk) & (m_cari['Tarih_Obj'].dt.date <= m_son)
                        filtrelenmis_m = m_cari[mask_m]
                        st.dataframe(filtrelenmis_m[["Tarih", "Islem_Detayi", "Borc", "Alacak", "Odeme_Tipi"]], use_container_width=True)
                        
                        buffer_musteri = io.BytesIO()
                        with pd.ExcelWriter(buffer_musteri, engine='openpyxl') as writer:
                            filtrelenmis_m[["Tarih", "Islem_Detayi", "Borc", "Alacak", "Odeme_Tipi"]].to_excel(writer, index=False, sheet_name='Musteri_Ekstre')
                        st.download_button(label="📥 Müşteri Ekstresini Excel İndir (.xlsx)", data=buffer_musteri.getvalue(), file_name=f"{secilen_musteri}_Ekstre.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        
                        st.divider()
                        with st.form("musteri_odeme", clear_on_submit=True):
                            c1, c2, c3 = st.columns(3)
                            m_tarih = c1.date_input("Tarih").strftime("%d.%m.%Y")
                            m_tutar = c2.number_input("Tutar", min_value=0.0)
                            m_detay = c3.text_input("Açıklama")
                            if st.form_submit_button("Tahsilatı İşle"):
                                with st.spinner("💳 Tahsilat Kaydediliyor..."):
                                    client.open_by_key(SHEET_ID).worksheet("Cari_Islemler").append_row([m_tarih, "Müşteri Carisi", secilen_musteri, m_detay, 0, m_tutar, "Nakit", 1])
                                    st.cache_data.clear()
                                    st.success("Tahsilat işlendi!")

    elif menu == "📅 Yenileme Takvimi":
        st.header("📅 Poliçe Yenileme ve Takip Merkezi")
        df_pol = get_data("Policeler")
        if not df_pol.empty and "Bitiş Tarihi" in df_pol.columns:
            def kalan_gun_hesapla(bitis_str):
                try:
                    if pd.isna(bitis_str) or str(bitis_str).strip() == "": return None
                    match = re.search(r'\d{2}[\./-]\d{2}[\./-]\d{4}', str(bitis_str))
                    if match:
                        clean_date = match.group().replace('/', '.').replace('-', '.')
                        bitis_date = pd.to_datetime(clean_date, format='%d.%m.%Y')
                        bugun = pd.to_datetime(datetime.today().strftime('%d.%m.%Y'), format='%d.%m.%Y')
                        return (bitis_date - bugun).days
                    return None
                except: return None

            df_pol["Kalan Gün"] = df_pol["Bitiş Tarihi"].apply(kalan_gun_hesapla)
            df_yaklasan = df_pol[(df_pol["Kalan Gün"].notna()) & (df_pol["Kalan Gün"] <= 30)].copy()
            
            if not df_yaklasan.empty:
                df_yaklasan = df_yaklasan.sort_values(by="Kalan Gün")
                def durum_belirle(gun):
                    if gun < 0: return "🚨 Süresi Doldu"
                    elif gun <= 10: return "🔴 Çok Yaklaştı"
                    elif gun <= 20: return "🟡 Yaklaşıyor"
                    return "🟢 Güvenli"
                df_yaklasan["Uyarı Durumu"] = df_yaklasan["Kalan Gün"].apply(durum_belirle)
                
                gosterim_sutunlari = ["Bitiş Tarihi", "Kalan Gün", "Uyarı Durumu", "Müşteri Adı Soyadı", "Telefon / E-mail", "Plaka", "Sigorta Türü"]
                if "PDF Linki" in df_yaklasan.columns: gosterim_sutunlari.append("PDF Linki")
                
                st.success(f"🔔 Önümüzdeki 30 gün içinde süresi dolacak {len(df_yaklasan)} adet poliçe bulundu.")
                st.dataframe(df_yaklasan[gosterim_sutunlari], column_config=STIL_AYARLARI, use_container_width=True)
                
                buffer_takvim = io.BytesIO()
                with pd.ExcelWriter(buffer_takvim, engine='openpyxl') as writer: df_yaklasan[gosterim_sutunlari].to_excel(writer, index=False, sheet_name='Yenileme_Takvimi')
                st.download_button("📥 Bu Ayın Yenileme Listesini Excel İndir (.xlsx)", data=buffer_takvim.getvalue(), file_name=f"Yenilemeler_{datetime.now().strftime('%m_%Y')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else: st.info("🎉 Harika! Önümüzdeki 30 gün içinde yenilenecek poliçeniz bulunmuyor.")
        else: st.info("Sistemde tarih hesaplaması yapılacak poliçe bulunamadı.")

    elif menu == "🔎 Genel Arama":
        st.header("🔎 Gelişmiş Poliçe Arama")
        df_pol = get_data("Policeler")
        if not df_pol.empty:
            arama_metni = st.text_input("🔍 İsim veya Plaka yazın:").upper()
            if arama_metni:
                mask = df_pol['Müşteri Adı Soyadı'].str.upper().str.contains(arama_metni, na=False) | \
                       df_pol['Plaka'].str.upper().str.contains(arama_metni, na=False)
                sonuclar = df_pol[mask]
                if not sonuclar.empty:
                    st.success(f"{len(sonuclar)} kayıt bulundu!")
                    sutun_ayarlari = {}
                    if "PDF Linki" in sonuclar.columns: sutun_ayarlari["PDF Linki"] = STIL_AYARLARI["PDF Linki"]
                    st.dataframe(sonuclar, column_config=sutun_ayarlari, use_container_width=True)
                    
                    buffer_arama = io.BytesIO()
                    with pd.ExcelWriter(buffer_arama, engine='openpyxl') as writer: sonuclar.to_excel(writer, index=False)
                    st.download_button("📥 Excel İndir (.xlsx)", data=buffer_arama.getvalue(), file_name="arama.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else: st.warning("Bulunamadı.")

    elif menu == "⚙️ Ayarlar":
        st.header("⚙️ Ayarlar")
        t1, t2 = st.tabs(["Ürün Oranları", "Acente Oranları"])
        with t1:
            st.dataframe(get_data("Ayarlar_Urunler"), use_container_width=True)
            with st.form("u", clear_on_submit=True): 
                u1=st.text_input("Ad"); u2=st.number_input("Oran", format="%.2f")
                if st.form_submit_button("Ekle"): client.open_by_key(SHEET_ID).worksheet("Ayarlar_Urunler").append_row([u1, u2]); st.cache_data.clear()
        with t2:
            st.dataframe(get_data("Ayarlar_Acenteler"), use_container_width=True)
            with st.form("a", clear_on_submit=True): 
                a1=st.text_input("Ad"); a2=st.number_input("Oran", format="%.2f")
                if st.form_submit_button("Ekle"): client.open_by_key(SHEET_ID).worksheet("Ayarlar_Acenteler").append_row([a1, a2]); st.cache_data.clear()

    elif menu == "🔍 Tüm Arşiv":
        st.header("🔍 Tüm Poliçe Arşivi")
        df_pol = get_data("Policeler")
        if not df_pol.empty:
            sutun_ayarlari = {}
            if "PDF Linki" in df_pol.columns: sutun_ayarlari["PDF Linki"] = STIL_AYARLARI["PDF Linki"]
            st.dataframe(df_pol, column_config=sutun_ayarlari, use_container_width=True)
            
            buffer_arsiv = io.BytesIO()
            with pd.ExcelWriter(buffer_arsiv, engine='openpyxl') as writer: df_pol.to_excel(writer, index=False)
            st.download_button("📥 Tüm Arşivi Excel Olarak İndir (.xlsx)", data=buffer_arsiv.getvalue(), file_name="Tum_Arsiv.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")