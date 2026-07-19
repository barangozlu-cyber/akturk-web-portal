import sys
import subprocess
# --- STANDART KÜTÜPHANELER ---
import streamlit as st
import pdfplumber
import pandas as pd
import re
from datetime import datetime
import io
import os
import random
from sqlalchemy import create_engine

# ... kodunun geri kalanı aynı şekilde devam ediyor ...

# ==========================================
# 1. PREMIUM ERP ARAYÜZ (UI/UX) CSS KODLARI
# ==========================================
st.set_page_config(page_title="Aktürk ERP v11.0 (PostgreSQL)", page_icon="🛡️", layout="wide", initial_sidebar_state="auto")

gizleme_kodu = """
<style>
:root { --primary: #0EA5E9; --primary-dark: #0284C7; --bg-color: #F8FAFC; --card-bg: #FFFFFF; --text-main: #1E293B; --text-muted: #64748B; }
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"] { font-family: 'Plus Jakarta Sans', sans-serif !important; }
.stApp { background-color: var(--bg-color) !important; color: var(--text-main) !important; }
h1, h2, h3, h4 { color: #0F172A !important; font-weight: 800 !important; letter-spacing: -0.03em !important; }
#MainMenu {visibility: hidden;} footer {visibility: hidden;} header[data-testid="stHeader"] { background: transparent !important; }
[data-testid="stSidebar"] { background-color: #FFFFFF !important; border-right: 1px solid #E2E8F0 !important; box-shadow: 4px 0 24px rgba(0,0,0,0.02) !important; }
[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label { padding: 12px 16px !important; margin-bottom: 8px !important; border-radius: 12px !important; font-weight: 600 !important; transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important; color: var(--text-muted); border: 1px solid transparent; }
[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label:hover { background-color: #F1F5F9 !important; color: #0F172A !important; transform: translateX(4px); }
[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label:has(input:checked) { background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%) !important; color: #FFFFFF !important; box-shadow: 0 4px 12px rgba(15, 23, 42, 0.2) !important; border: none !important; }
[data-testid="stSidebar"] .stRadio div[role="radiogroup"] label:has(input:checked) p { color: #FFFFFF !important; }
div[data-testid="stMetric"] { background-color: var(--card-bg) !important; border-radius: 16px !important; padding: 24px !important; border: 1px solid #E2E8F0 !important; border-left: 6px solid var(--primary) !important; }
.section-header { color: #0284C7; font-size: 1.1rem; font-weight: 700; margin-bottom: 1rem; text-transform: uppercase; border-bottom: 2px solid #E0F2FE; padding-bottom: 0.5rem; }
</style>
"""
st.markdown(gizleme_kodu, unsafe_allow_html=True)

# ==========================================
# 2. BULUT VERİTABANI (POSTGRESQL) AYARLARI
# ==========================================
VERSIYON = "v11.0 (PostgreSQL Sunucusu)"
PDF_DIR = "uploads"

# Senin CasaOS sunucu bilgilerin koda entegre edildi
DB_USER = "casaos"
DB_PASS = "casaos"
DB_HOST = "192.168.1.14"
DB_PORT = "5432"
DB_NAME = "casaos"

try:
    # Veritabanı motorunu oluştur
    db_url = f'postgresql+psycopg2://{DB_USER}:{DB_PASS}@{DB_HOST}:{DB_PORT}/{DB_NAME}'
    engine = create_engine(db_url)
except Exception as e:
    st.error(f"Veritabanı bağlantı motoru oluşturulamadı: {e}")

if not os.path.exists(PDF_DIR):
    os.makedirs(PDF_DIR)

# ==========================================
# 3. YARDIMCI FONKSİYONLAR (SQL ENTEGRELİ)
# ==========================================
if "giris_yapildi" not in st.session_state: st.session_state["giris_yapildi"] = False

def ekran_temizle():
    for key in list(st.session_state.keys()):
        if key not in ["giris_yapildi", "kullanici_adi"]: del st.session_state[key]

@st.cache_data(ttl=5)
def get_data(table_name):
    """PostgreSQL sunucusundan tabloyu çeker"""
    try:
        df = pd.read_sql_table(table_name, engine)
        df = df.fillna("")
        if not df.empty:
            for col in ["Müşteri Adı Soyadı", "Kisi_Kurum", "Musteri_Adi", "Sirket_Adi", "Sigorta Şirketi", "Acente", "Acente_Adi"]:
                if col in df.columns: df[col] = df[col].apply(temiz_isim)
        return df
    except Exception:
        return pd.DataFrame()

def guvenli_kayit(table_name, veri_listesi_dict):
    """Sözlük listesini PostgreSQL sunucusuna INSERT eder"""
    try:
        yeni_df = pd.DataFrame(veri_listesi_dict)
        yeni_df.to_sql(table_name, engine, if_exists='append', index=False)
        get_data.clear() # Cache temizle
        return True
    except Exception as e:
        st.error(f"🚨 Sunucuya yazarken hata oluştu: {e}")
        return False

def sunucuya_pdf_kaydet(file_bytes, original_name):
    """PDF'i yerel uploads klasörüne kaydeder"""
    try:
        dosya_adi = original_name.replace("/", "_").replace("\\", "_")
        file_path = os.path.join(PDF_DIR, dosya_adi)
        with open(file_path, "wb") as f:
            f.write(file_bytes)
        return file_path
    except Exception: return "Yok"

def temiz_isim(metin):
    if pd.isna(metin) or not metin: return ""
    return str(metin).strip().replace('i', 'İ').replace('ı', 'I').upper()

def sayiya_cevir(deger):
    if pd.isna(deger) or str(deger).strip() == "": return 0.0
    if isinstance(deger, (int, float)): return float(deger)
    deger_str = re.sub(r'[^\d.,-]', '', str(deger).strip()).rstrip('.,')
    if not deger_str: return 0.0
    if '.' in deger_str and ',' in deger_str:
        if deger_str.rfind(',') > deger_str.rfind('.'): deger_str = deger_str.replace('.', '').replace(',', '.')
        else: deger_str = deger_str.replace(',', '')
    elif ',' in deger_str:
        if deger_str.count(',') > 1: deger_str = deger_str.replace(',', '')
        else: deger_str = deger_str.replace(',', '.')
    try: return float(deger_str)
    except: return 0.0

def para_format(deger):
    try: return "{:,.2f}".format(sayiya_cevir(deger)).replace(",", "X").replace(".", ",").replace("X", ".") + " TL"
    except: return "0,00 TL"

def tarih_formatla(tarih_degeri):
    if pd.isna(tarih_degeri) or str(tarih_degeri).strip() == "": return datetime.now().strftime("%d.%m.%Y")
    try: return pd.to_datetime(str(tarih_degeri).strip(), dayfirst=True).strftime("%d.%m.%Y")
    except: return datetime.now().strftime("%d.%m.%Y")

def fis_no_uret(islem_tipi="POL"): return f"{islem_tipi}-{datetime.now().strftime('%y%m%d%H%M')}{random.randint(10, 99)}"

def klasik_analiz(metin):
    data = {"tanzim": "", "baslangic": "", "bitis": "", "musteri": "", "tc_vkn": "", "sirket": "", "urun": "", "p_no": "", "plaka": "", "net_prim": 0.0, "brut_prim": 0.0}
    metin_upper = metin.upper()
    sirketler = {"ANKARA SİGORTA": "Ankara Sigorta", "DOĞA SİGORTA": "Doğa Sigorta", "ALLIANZ": "Allianz Sigorta", "HDI SİGORTA": "HDI Sigorta", "HDİ": "HDI Sigorta"}
    for anahtar, deger in sirketler.items():
        if anahtar in metin_upper: data["sirket"] = deger; break
    
    tarihler = re.findall(r'\b\d{2}[\./-]\d{2}[\./-]\d{4}\b', metin)
    if len(tarihler) >= 3: data["tanzim"], data["baslangic"], data["bitis"] = [t.replace("/", ".") for t in tarihler[:3]]

    plaka = re.search(r'\b(0[1-9]|[1-7][0-9]|8[0-1])\s*[A-Z]{1,3}\s*[0-9]{2,4}\b', metin_upper)
    if plaka: data["plaka"] = plaka.group().replace(" ", "")

    primler = re.findall(r'(\d+[\.,]\d{2})\s*(?:TL|TRY|€|\$|EUR|USD)?', metin)
    if primler:
        prim_listesi = sorted([sayiya_cevir(p) for p in primler], reverse=True)
        if len(prim_listesi) >= 1: data["brut_prim"] = prim_listesi[0]
        if len(prim_listesi) >= 2: data["net_prim"] = prim_listesi[1]
    return data

# ==========================================
# 4. GİRİŞ SİSTEMİ 
# ==========================================
if not st.session_state["giris_yapildi"]:
    st.markdown("<div class='login-box'>", unsafe_allow_html=True)
    st.markdown("""<div style="background: linear-gradient(135deg, #1D4ED8 0%, #1E3A8A 100%); width: 88px; height: 88px; border-radius: 20px; display: flex; align-items: center; justify-content: center; margin: 0 auto 20px auto; box-shadow: 0 10px 25px rgba(29, 78, 216, 0.4); transform: rotate(-5deg);"><span style="font-size: 48px; font-weight: 900; color: #FFFFFF; transform: rotate(5deg);">A</span></div><h2 style='margin-top:0px; margin-bottom: 5px; color:#0F172A; font-weight: 800; text-align: center;'>Aktürk Sigorta</h2><p style='color: #64748B; font-size: 15px; margin-bottom: 30px; font-weight: 600; text-align: center;'>Kurumsal Yönetim Portalı</p>""", unsafe_allow_html=True)
    
    with st.container(border=True):
        u = st.text_input("Kullanıcı Adı")
        p = st.text_input("Şifre", type="password")
        if st.button("Sisteme Giriş Yap", type="primary", use_container_width=True):
            df_u = get_data("Kullanicilar")
            if not df_u.empty and "Kullanıcı Adı" in df_u.columns:
                eslesen = df_u[df_u["Kullanıcı Adı"] == u]
                if not eslesen.empty and str(eslesen["Şifre"].values[0]) == str(p):
                    st.session_state["giris_yapildi"] = True; st.session_state["kullanici_adi"] = u; st.rerun()
                else: st.error("⚠️ Hatalı kullanıcı adı veya şifre!")
            elif u == "admin" and p == "123456": # Veritabanı boşken ilk giriş için acil durum kapısı
                st.session_state["giris_yapildi"] = True; st.session_state["kullanici_adi"] = u; st.rerun()
            else:
                st.error("⚠️ Hatalı kullanıcı adı veya şifre!")
    st.markdown("</div>", unsafe_allow_html=True)

# ==========================================
# ANA UYGULAMA MANTIGI
# ==========================================
else:
    st.sidebar.markdown("""<div style='display: flex; align-items: center; margin-bottom: 30px; margin-top: 10px; padding: 10px; background-color: #F8FAFC; border-radius: 12px; border: 1px solid #E2E8F0;'><div style='background: linear-gradient(135deg, #1D4ED8 0%, #1E3A8A 100%); width: 42px; height: 42px; border-radius: 10px; display: flex; align-items: center; justify-content: center; margin-right: 15px; box-shadow: 0 4px 10px rgba(29, 78, 216, 0.3);'><span style='font-size: 20px; font-weight: 900; color: #FFFFFF;'>A</span></div><div><h2 style='margin: 0; color: #0F172A; font-size: 18px; font-weight: 800; letter-spacing: -0.5px;'>Aktürk ERP</h2><span style='color: #64748B; font-size: 12px; font-weight: 700;'>Kullanıcı: {0}</span></div></div>""".format(st.session_state['kullanici_adi'].upper()), unsafe_allow_html=True)
    menu = st.sidebar.radio("", ["📥 İşlem Merkezi (Poliçe)", "💰 Cari & Mutabakat", "📂 Genel Arşiv", "🚪 Çıkış"], on_change=ekran_temizle, label_visibility="collapsed")
    st.sidebar.markdown(f"<div style='text-align: center; margin-top: 20px;'><span style='color: #94A3B8; font-size: 11px; font-weight: 600;'>{VERSIYON}</span></div>", unsafe_allow_html=True)

    if menu == "📥 İşlem Merkezi (Poliçe)":
        st.header("Yeni Poliçe & İşlem Kaydı")
        
        # Ayarları Veritabanından Çek
        df_urun = get_data("Ayarlar_Urunler")
        urun_listesi = df_urun["Urun_Adi"].tolist() if not df_urun.empty else ["Trafik Sigortası", "Kasko", "Sağlık Sigortası"]
        dict_urun = dict(zip(df_urun['Urun_Adi'], df_urun['Komisyon_Orani'])) if not df_urun.empty else {}
        
        df_acente = get_data("Ayarlar_Acenteler")
        acente_listesi = df_acente["Acente_Adi"].tolist() if not df_acente.empty else ["AKTÜRK SİGORTA (MERKEZ)"]
        if "AKTÜRK SİGORTA (MERKEZ)" not in acente_listesi: acente_listesi.insert(0, "AKTÜRK SİGORTA (MERKEZ)")
        dict_acente = dict(zip(df_acente['Acente_Adi'], df_acente['Tali_Oran'])) if not df_acente.empty else {}
        acente_listesi.append("➕ YENİ TALİ ACENTE EKLE")

        df_sirket_ayar = get_data("Ayarlar_Sirketler")
        sirket_listesi = df_sirket_ayar["Sirket_Adi"].tolist() if not df_sirket_ayar.empty else ["Doğa Sigorta", "Allianz Sigorta"]
        sirket_listesi.append("➕ YENİ ŞİRKET EKLE")

        with st.container(border=True):
            st.markdown("<div class='section-header'>1. PDF TARAMA MOTORU</div>", unsafe_allow_html=True)
            file = st.file_uploader("PDF Poliçe Seçin", type="pdf", label_visibility="collapsed")
            p_data = {"tanzim":"","baslangic":"","bitis":"","musteri":"","plaka":"","tc_vkn":"","sirket":"","urun":"","p_no":"","net_prim":0.0,"brut_prim":0.0}
            f_bytes = None
            
            if file:
                f_bytes = file.getvalue()
                with st.spinner("📄 PDF Okunuyor..."):
                    with pdfplumber.open(io.BytesIO(f_bytes)) as pdf:
                        if txt := pdf.pages[0].extract_text(): 
                            p_data.update(klasik_analiz(txt))
                            st.toast("PDF başarıyla tarandı!", icon="✅")

        with st.form("police_formu", clear_on_submit=True):
            st.markdown("<div class='section-header'>MÜŞTERİ VE POLİÇE DETAYLARI</div>", unsafe_allow_html=True)
            c1, c2, c3 = st.columns(3)
            tan = c1.text_input("Tanzim Tarihi", p_data["tanzim"])
            bas = c2.text_input("Başlangıç", p_data["baslangic"])
            bit = c3.text_input("Bitiş", p_data["bitis"])
            
            c4, c5, c6 = st.columns(3)
            mus_girdi = c4.text_input("Müşteri Ad Soyad", p_data["musteri"])
            tc = c5.text_input("TC / VKN", p_data["tc_vkn"])
            plk = c6.text_input("Plaka", p_data["plaka"])
            
            c7, c8, c9 = st.columns(3)
            sir_girdi = c7.selectbox("Sigorta Şirketi", sirket_listesi)
            yeni_sirket_adi = c7.text_input("Yeni Şirket Adı:") if sir_girdi == "➕ YENİ ŞİRKET EKLE" else ""
            pno = c8.text_input("Poliçe No", p_data["p_no"])
            urn = c9.selectbox("Ürün Türü", urun_listesi)

            st.markdown("<div class='section-header'>FİNANS VE KOMİSYON</div>", unsafe_allow_html=True)
            c10, c11, c12 = st.columns(3)
            net_girdi = c10.text_input("Net Prim", value=str(p_data["net_prim"]))
            brut_girdi = c11.text_input("Brüt Prim", value=str(p_data["brut_prim"]))
            alinan_odeme_str = c12.text_input("Peşinat Alınan (TL)", value="0")
            
            acn = st.selectbox("İşlemi Yapan / Kesilen Ekran", acente_listesi)
            yeni_acente_adi = ""; yeni_acente_orani = 0.0
            if acn == "➕ YENİ TALİ ACENTE EKLE":
                yeni_acente_adi = st.text_input("Yeni Acente Adı:")
                yeni_acente_orani = float(sayiya_cevir(st.text_input("Tali Komisyon Oranı (Örn: 70)", value="0")))
            
            submit_btn = st.form_submit_button("Sisteme İşle ve Kaydet", type="primary", use_container_width=True)
            
            if submit_btn:
                with st.status("Veritabanına Yazılıyor...", expanded=True) as status:
                    try:
                        mus = temiz_isim(mus_girdi)
                        sir = temiz_isim(sir_girdi) if sir_girdi != "➕ YENİ ŞİRKET EKLE" else temiz_isim(yeni_sirket_adi)
                        aktif_acente = temiz_isim(acn) if acn != "➕ YENİ TALİ ACENTE EKLE" else temiz_isim(yeni_acente_adi)
                        plk_temiz = str(plk).replace(" ", "").upper()
                        yeni_fis_no = fis_no_uret("POL")
                        
                        net = float(sayiya_cevir(net_girdi)); brut = float(sayiya_cevir(brut_girdi))
                        alinan_odeme = float(sayiya_cevir(alinan_odeme_str))
                        
                        link = sunucuya_pdf_kaydet(f_bytes, f"{mus}_{plk_temiz}.pdf") if f_bytes else "Yok"
                        
                        u_oran = float(sayiya_cevir(dict_urun.get(urn, 0.0)))
                        sirket_komisyonu = net * (u_oran / 100 if u_oran > 1 else u_oran)
                        
                        if aktif_acente == temiz_isim(yeni_acente_adi):
                            guvenli_kayit("Ayarlar_Acenteler", [{"Acente_Adi": aktif_acente, "Tali_Oran": yeni_acente_orani}])
                            t_oran = yeni_acente_orani
                        else: t_oran = float(sayiya_cevir(dict_acente.get(aktif_acente, 0.0)))
                        
                        akturk_kazanci = sirket_komisyonu * (t_oran / 100 if t_oran > 1 else t_oran)

                        if sir_girdi == "➕ YENİ ŞİRKET EKLE" and yeni_sirket_adi != "":
                            guvenli_kayit("Ayarlar_Sirketler", [{"Sirket_Adi": sir}])

                        # SQL INSERT İŞLEMLERİ (Veri Sözlükleri)
                        pol_veri = {
                            "Tanzim Tarihi": tarih_formatla(tan), "Başlangıç Tarihi": bas, "Bitiş Tarihi": bit, 
                            "Müşteri Adı Soyadı": mus, "TC / VKN": tc, "Sigorta Şirketi": sir, "Sigorta Türü": urn, 
                            "Poliçe No": pno, "Plaka": plk_temiz, "Net Prim": net, "Brüt Prim": brut, 
                            "Şirket Komisyonu": sirket_komisyonu, "Acente": aktif_acente, "PDF Linki": link
                        }
                        
                        cari_veriler = []
                        aciklama = f"{sir} - {urn} - Plaka: {plk_temiz}"
                        
                        # Müşteri Carisi
                        cari_veriler.append({"Tarih": tarih_formatla(tan), "Islem_Turu": "Müşteri Carisi", "Kisi_Kurum": mus, "Islem_Detayi": aciklama, "Borc": brut, "Alacak": 0.0, "Fiş No": yeni_fis_no})
                        if alinan_odeme > 0:
                            cari_veriler.append({"Tarih": tarih_formatla(tan), "Islem_Turu": "Müşteri Carisi", "Kisi_Kurum": mus, "Islem_Detayi": "Anında Tahsilat", "Borc": 0.0, "Alacak": alinan_odeme, "Fiş No": yeni_fis_no})
                        
                        # Acente / Şirket Carisi
                        if aktif_acente == "AKTÜRK SİGORTA (MERKEZ)":
                            cari_veriler.append({"Tarih": tarih_formatla(tan), "Islem_Turu": "Sigorta Şirketi Carisi", "Kisi_Kurum": sir, "Islem_Detayi": f"Şirket Komisyonu Hakediş - {aciklama}", "Borc": sirket_komisyonu, "Alacak": 0.0, "Fiş No": yeni_fis_no})
                        else:
                            cari_veriler.append({"Tarih": tarih_formatla(tan), "Islem_Turu": "Tali Acente Carisi", "Kisi_Kurum": aktif_acente, "Islem_Detayi": f"Acente Payı Hakediş - {aciklama}", "Borc": akturk_kazanci, "Alacak": 0.0, "Fiş No": yeni_fis_no})

                        if guvenli_kayit("Policeler", [pol_veri]) and guvenli_kayit("Cari_Islemler", cari_veriler):
                            status.update(label="Kayıt Başarılı!", state="complete")
                            st.toast("Veriler PostgreSQL sunucusuna işlendi.", icon="🎉")
                            st.balloons()
                        else:
                            status.update(label="Hata!", state="error")
                            
                    except Exception as e:
                        st.error(f"Sistem Hatası: {e}")

    elif menu == "💰 Cari & Mutabakat":
        st.header("Cari & Mutabakat Bilançoları (Bulut)")
        df_cari = get_data("Cari_Islemler")
        
        if not df_cari.empty:
            df_ozet = df_cari[df_cari["Kisi_Kurum"].str.strip() != ""].copy()
            df_ozet["Borc"] = df_ozet["Borc"].apply(sayiya_cevir)
            df_ozet["Alacak"] = df_ozet["Alacak"].apply(sayiya_cevir)
            
            grup = df_ozet.groupby(["Islem_Turu", "Kisi_Kurum"])[["Borc", "Alacak"]].sum().reset_index()
            grup["Bakiye"] = grup["Borc"] - grup["Alacak"]
            grup.rename(columns={"Islem_Turu": "Hesap Türü", "Kisi_Kurum": "Kişi/Kurum", "Borc": "Toplam Borç", "Alacak": "Toplam Alacak", "Bakiye": "Kalan Bakiye"}, inplace=True)
            
            st.dataframe(df_gorsel_yap(grup, ["Toplam Borç", "Toplam Alacak", "Kalan Bakiye"]), use_container_width=True)
        else:
            st.info("Veritabanında henüz cari işlem bulunmuyor.")

    elif menu == "📂 Genel Arşiv":
        st.header("Tüm Poliçe Arşivi (Bulut)")
        df_pol = get_data("Policeler")
        if not df_pol.empty:
            st.dataframe(df_gorsel_yap(df_pol, ["Net Prim", "Brüt Prim", "Şirket Komisyonu"]), use_container_width=True)
        else:
            st.info("Veritabanında poliçe kaydı bulunmuyor.")
            
    elif menu == "🚪 Çıkış":
        st.session_state["giris_yapildi"] = False
        st.rerun()
