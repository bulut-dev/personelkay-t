import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
import sys
import pandas as pd
import shutil
import webbrowser  # Web sayfasını açmak için gerekli kütüphane
from datetime import datetime

def resource_path(relative_path):
    """ 
    PyInstaller ile paketlendiğinde dosyaların geçici dizindeki 
    yollarını bulmasını sağlayan yardımcı fonksiyon.
    """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class PersonelSistemiGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("mustafaozkanbulut.org")
        self.root.geometry("1200x700")
        self.root.minsize(1000, 600)
        
        # Ekranın yeniden boyutlandırılabilir olması için ağırlıkları yapılandırıyoruz
        self.root.columnconfigure(1, weight=1) # Sağ taraf genişleyebilir
        self.root.rowconfigure(0, weight=1)    # Satır dikeyde genişleyebilir

        # İkon dosyası kontrolü
        try:
            icon_path = resource_path("uygulama_ikonu.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except:
            pass

        self.dosya_adi = 'personeller.json'
        
        # Veri Yükleme
        self.personeller = self.verileri_yukle()

        # UI Bileşenleri
        self.arayuz_olustur()
        self.listeyi_guncelle()

    def verileri_yukle(self):
        if os.path.exists(self.dosya_adi):
            try:
                with open(self.dosya_adi, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                return []
        return []

    def verileri_kaydet(self):
        try:
            with open(self.dosya_adi, 'w', encoding='utf-8') as f:
                json.dump(self.personeller, f, ensure_ascii=False, indent=4)
        except Exception as e:
            messagebox.showerror("Hata", f"Veriler kaydedilirken bir hata oluştu: {e}")

    def open_blog(self, event):
        """Web tarayıcısını belirtilen adreste açar."""
        webbrowser.open_new("https://mustafaozkanbulut.org")

    def arayuz_olustur(self):
        # Sol Panel - Ana Taşıyıcı
        left_container = tk.Frame(self.root)
        left_container.grid(row=0, column=0, padx=20, pady=20, sticky="ns")

        # Tıklanabilir Blog Linki
        link_label = tk.Label(
            left_container, 
            text="mustafaozkanbulut.org", 
            font=("Arial", 11, "bold", "underline"), 
            fg="blue", 
            cursor="hand2"
        )
        link_label.pack(side="top", anchor="w", pady=(10, 25))
        link_label.bind("<Button-1>", self.open_blog)

        # Form Alanı (Başlık tekrar "Personel Bilgileri")
        form_frame = tk.LabelFrame(left_container, text="Personel Bilgileri", padx=10, pady=10)
        form_frame.pack(side="top", fill="both", expand=True)

        label_font = ("Arial", 10, "bold")
        fields = [
            ("TC Kimlik No:", "tc"),
            ("Ad Soyad:", "ad_soyad"),
            ("Görevi:", "gorevi"),
            ("Tel No:", "tel"),
            ("İşe Giriş Tarihi:", "ise_giris"),
            ("İşten Ayrılış Tarihi:", "isten_ayrilis"),
            ("IBAN:", "iban"),
            ("Adres:", "adres")
        ]

        self.entries = {}
        for i, (label_text, key) in enumerate(fields):
            tk.Label(form_frame, text=label_text, font=label_font).grid(row=i*2, column=0, sticky="w", pady=(2, 0))
            entry = tk.Entry(form_frame, font=("Arial", 10))
            entry.grid(row=i*2 + 1, column=0, sticky="ew", pady=(0, 5))
            self.entries[key] = entry

        # Butonlar
        btn_frame = tk.Frame(form_frame)
        btn_frame.grid(row=16, column=0, pady=15)

        tk.Button(btn_frame, text="Ekle/Güncelle", command=self.personel_ekle, bg="#4CAF50", fg="white", width=12).pack(side="left", padx=5)
        tk.Button(btn_frame, text="Temizle", command=self.formu_temizle, bg="#2196F3", fg="white", width=12).pack(side="left", padx=5)
        
        tk.Button(form_frame, text="Seçili Kaydı Sil", command=self.personel_sil, bg="#f44336", fg="white").grid(row=17, column=0, sticky="ew", padx=5, pady=5)
        tk.Button(form_frame, text="Sistem Yedeği Al (.json)", command=self.yedek_al, bg="#607D8B", fg="white").grid(row=18, column=0, sticky="ew", padx=5, pady=5)

        # Sağ Panel - Liste ve Arama (Esnek genişlik)
        list_frame = tk.Frame(self.root)
        list_frame.grid(row=0, column=1, padx=(0, 20), pady=20, sticky="nsew")

        # Üst Panel (Arama ve Excel)
        top_panel = tk.Frame(list_frame)
        top_panel.pack(fill="x", pady=(0, 10))

        tk.Label(top_panel, text="Ara (TC/İsim):", font=label_font).pack(side="left")
        self.search_entry = tk.Entry(top_panel)
        self.search_entry.pack(side="left", fill="x", expand=True, padx=10)
        self.search_entry.bind("<KeyRelease>", lambda e: self.listeyi_guncelle())

        tk.Button(top_panel, text="Excel'e Aktar", command=self.excel_aktar, bg="#1D6F42", fg="white", padx=10).pack(side="right")

        # Tablo
        table_container = tk.Frame(list_frame)
        table_container.pack(fill="both", expand=True)

        columns = ("tc", "ad_soyad", "gorevi", "tel", "ise_giris", "isten_ayrilis", "iban")
        self.tree = ttk.Treeview(table_container, columns=columns, show="headings")
        
        for col in columns:
            self.tree.heading(col, text=col.replace("_", " ").upper())
            self.tree.column(col, minwidth=100, width=120, stretch=True)

        # Scrollbarlar
        v_scrollbar = ttk.Scrollbar(table_container, orient="vertical", command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(list_frame, orient="horizontal", command=self.tree.xview)
        
        self.tree.configure(yscroll=v_scrollbar.set, xscroll=h_scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        v_scrollbar.pack(side="right", fill="y")
        h_scrollbar.pack(side="bottom", fill="x")

        self.tree.bind("<<TreeviewSelect>>", self.kayit_sec)

    def yedek_al(self):
        if not os.path.exists(self.dosya_adi):
            messagebox.showwarning("Hata", "Yedeklenecek veri dosyası bulunamadı.")
            return

        tarih = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        varsayilan_ad = f"personel_yedek_{tarih}.json"
        
        dosya_yolu = filedialog.asksaveasfilename(
            defaultextension=".json",
            initialfile=varsayilan_ad,
            title="Yedek Dosyasını Kaydet"
        )

        if dosya_yolu:
            try:
                shutil.copy2(self.dosya_adi, dosya_yolu)
                messagebox.showinfo("Başarılı", "Yedekleme başarıyla tamamlandı.")
            except Exception as e:
                messagebox.showerror("Hata", f"Yedekleme hatası: {e}")

    def excel_aktar(self):
        if not self.personeller:
            messagebox.showwarning("Uyarı", "Aktarılacak kayıt yok.")
            return

        dosya_yolu = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            title="Excel Olarak Kaydet"
        )

        if dosya_yolu:
            try:
                df = pd.DataFrame(self.personeller)
                df.to_excel(dosya_yolu, index=False)
                messagebox.showinfo("Başarılı", "Excel dosyası başarıyla oluşturuldu.")
            except Exception as e:
                messagebox.showerror("Hata", f"Excel hatası: {e}")

    def alfanumerik_ve_bosluk_mu(self, metin):
        """Metnin harf, rakam ve boşluk içerip içermediğini kontrol eder (Türkçe karakterler dahil)."""
        return all(char.isalnum() or char == " " for char in metin)

    def personel_ekle(self):
        data = {key: entry.get().strip() for key, entry in self.entries.items()}
        
        # TC Doğrulaması (11 Hane ve Sadece Rakam)
        if not data["tc"].isdigit() or len(data["tc"]) != 11:
            messagebox.showwarning("Hata", "TC Kimlik No sadece rakamlardan oluşmalı ve 11 hane olmalıdır!")
            return

        # Ad Soyad Doğrulaması (Zorunlu ve Alfanumerik)
        if not data["ad_soyad"]:
            messagebox.showwarning("Hata", "Ad Soyad alanı boş bırakılamaz!")
            return
        
        if not self.alfanumerik_ve_bosluk_mu(data["ad_soyad"]):
            messagebox.showwarning("Hata", "Ad Soyad sadece harf, rakam ve boşluk içerebilir!")
            return

        # Görevi Doğrulaması (Alfanumerik)
        if data["gorevi"] and not self.alfanumerik_ve_bosluk_mu(data["gorevi"]):
            messagebox.showwarning("Hata", "Görevi alanı sadece harf, rakam ve boşluk içerebilir!")
            return

        # Telefon numarası kontrolü: Sadece rakam ve boşluklara izin ver
        if data["tel"] and not all(char.isdigit() or char == " " for char in data["tel"]):
            messagebox.showwarning("Hata", "Telefon sadece rakam ve boşluk içerebilir!")
            return

        # Kayıt İşlemi
        existing_index = next((i for i, p in enumerate(self.personeller) if p["tc"] == data["tc"]), None)
        if existing_index is not None:
            if messagebox.askyesno("Güncelle", "Bu TC ile kayıt mevcut. Güncellensin mi?"):
                self.personeller[existing_index] = data
            else: return
        else:
            self.personeller.append(data)

        self.verileri_kaydet()
        self.listeyi_guncelle()
        # Kayıt başarılıysa formu temizle ve odaklan
        self.formu_temizle()
        messagebox.showinfo("Başarılı", "İşlem başarıyla tamamlandı.")

    def listeyi_guncelle(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        search = self.search_entry.get().lower()
        for p in self.personeller:
            if search in p.get("tc", "").lower() or search in p.get("ad_soyad", "").lower():
                display_values = [
                    p.get("tc", ""),
                    p.get("ad_soyad", ""),
                    p.get("gorevi", ""),
                    p.get("tel", ""),
                    p.get("ise_giris", ""),
                    p.get("isten_ayrilis", ""),
                    p.get("iban", "")
                ]
                self.tree.insert("", "end", values=display_values)

    def kayit_sec(self, event):
        selected = self.tree.selection()
        if selected:
            item = self.tree.item(selected[0])
            tc_val = str(item["values"][0])
            personel = next((p for p in self.personeller if p["tc"] == tc_val), None)
            if personel:
                self.formu_temizle()
                for key, entry in self.entries.items():
                    entry.insert(0, personel.get(key, ""))

    def personel_sil(self):
        selected = self.tree.selection()
        if selected:
            item = self.tree.item(selected[0])
            tc_val = str(item["values"][0])
            if messagebox.askyesno("Onay", f"{tc_val} TC nolu personeli silmek istediğinize emin misiniz?"):
                self.personeller = [p for p in self.personeller if p["tc"] != tc_val]
                self.verileri_kaydet()
                self.listeyi_guncelle()
                self.formu_temizle()

    def formu_temizle(self):
        """Tüm giriş kutularını temizler ve odağı ilk kutuya getirir."""
        for entry in self.entries.values():
            entry.delete(0, tk.END)
        # İmleci tekrar en üstteki TC alanına getirelim
        if "tc" in self.entries:
            self.entries["tc"].focus_set()

if __name__ == "__main__":
    root = tk.Tk()
    app = PersonelSistemiGUI(root)
    root.mainloop()
