# Teams Oto-Katılma ⚡

Microsoft Teams toplantısından atıldığında **otomatik olarak tekrar katılan** Python uygulaması.

## Özellikler

- 🔄 Toplantıdan atılınca **3 saniye içinde** otomatik tekrar katılır
- 🖱️ Fareye **dokunmaz** — oyun oynarken bile kullanılabilir
- 👻 Ekrana **hiçbir şey gelmez** — tamamen arka planda çalışır
- 🎮 **Tam ekran oyun** oynarken bile çalışır
- 🖥️ Arayüz ile kolay kontrol (Başlat / Duraklat / Durdur)
- 📞 Toplantıdan çıkış butonu
- 📋 Olay günlüğü

## Kurulum

**Gereksinim:** [Python 3.8+](https://www.python.org/downloads/) (kurulumda "Add to PATH" seçeneğini işaretleyin)

```bash
git clone https://github.com/KULLANICI/msteams.git
cd msteams
```

İlk çalıştırmada gerekli paketler (`pywinauto`, `pywin32`) **otomatik kurulur**.

## Kullanım

### Arayüz ile (terminal penceresi açılmaz)
```
teams_auto_rejoin.pyw
```
Dosyaya çift tıklayın.

### Terminal ile
```bash
python teams_auto_rejoin.py
```

## Ekran Görüntüsü

Uygulama açıldığında koyu temalı bir arayüz ile karşılaşırsınız:

- **▶ Başlat** — izlemeyi başlatır
- **⏸ Duraklat / ▶ Devam** — izlemeyi duraklatır/devam ettirir
- **⏹ Durdur** — izlemeyi durdurur
- **🔄 Şimdi Katıl** — toplantıya manuel katılır
- **📞 Toplantıdan Çık** — toplantıdan çıkar

## Nasıl Çalışır

1. Windows API ile Teams pencere başlıklarını izler (pencereye dokunmadan)
2. "Daraltılmış görünümle" toplantı penceresi kaybolunca → atıldığınızı algılar
3. UI Automation ile "Anında toplantı" butonuna tıklar (fare hareket etmez)
4. Toplantı penceresi açılınca otomatik minimize eder

## Lisans

MIT
