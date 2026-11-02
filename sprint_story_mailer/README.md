# RTC Sprint Story Mailer

Sprint bitiş günlerinde RTC'deki story query'sini (`_QdyPMNaZEfG_RJ8I86d6LA`)
çalıştırır ve sonuçları Outlook üzerinden şu adreslere mail atar:

- sf01089@external.stellantis.com
- t59612@tofas.com.tr

## Nasıl çalışır

- Görev Zamanlayıcı script'i **her gün 16:00'da** çalıştırır (PC kapalıysa açılınca çalışır).
- Sprint bitiş günleri **RTC'deki sprint takviminden** okunur (proje alanı
  `_sTK-QebpEea8oNaj29Rsbg`). Okunamazsa yedek takvim kullanılır:
  `anchor_end_date` = **2026-11-10**, 14 günde bir.
- Son sprint bitiş gününün maili gönderilmediyse gönderir, gönderildiyse hiçbir şey yapmaz
  (`sent_state.json`). Sprint günü kaçırılırsa (PC kapalı, VPN yok) sonraki
  `catch_up_days` (3) gün içinde telafi edilir.

## VPN / RTC bağlantısı kurulamazsa

Sadece mail günüyse (sprint günü veya telafi günü):

1. İlk başarısız denemede: **Windows bildirimi** + sf01089@external.stellantis.com
   adresine **"[UYARI]" maili** ("VPN'e bağlanın, 15 dk'da bir tekrar denenecek").
2. 15 dakikada bir, toplam 8 kez (yaklaşık 2 saat) tekrar dener. Bu arada VPN'e
   bağlanırsanız mail kendiliğinden gider.
3. Hepsi başarısız olursa: tekrar bildirim + uyarı maili. Ertesi gün yine dener.

Diğer hatalarda (şifre hatalı, query hatası, vb.) de aynı şekilde bildirim yapılır.
Outlook VPN olmadan da internet üzerinden mail gönderebildiği için uyarı maili ulaşır.
Normal günlerde VPN yoksa sizi rahatsız etmez.

## Kurulum

1. Paketleri kurun:
   ```
   pip install -r requirements.txt
   ```
2. `config.example.json` dosyasını `config.json` olarak kopyalayın ve
   `rtc.password` alanına şifrenizi yazın. `config.json` git'e gönderilmez
   (`.gitignore`), sadece sizin PC'nizde durur.
3. Sprint takvimini kontrol edin:
   ```
   python sprint_story_mailer.py --list-sprints
   ```
   "Son gün" sütununda mevcut sprintin bitişi **2026-11-10** görünmeli.
   - 1 gün sonrası (11-11) görünüyorsa: `end_date_offset_days: -1` yapın.
   - Alakasız timeline'lar / iterasyonlar listeleniyorsa: `timeline_name` veya
     `name_filter` (örn. `"Sprint"`) ile daraltın.
4. Test:
   ```
   python sprint_story_mailer.py --test-notify      (uyarı bildirimi + maili dener)
   python sprint_story_mailer.py --force --dry-run  (mail atmadan HTML önizleme)
   python sprint_story_mailer.py --force            (gerçekten mail atar)
   ```
   Not: `--force` ile atılan mail, sprint gününün "gönderildi" kaydını oluşturmaz.
5. Otomatik çalıştırma: `create_scheduled_task.bat` dosyasına çift tıklayın.

## Notlar

- Görev "kullanıcı oturum açıkken" çalışır (Outlook için gerekli).
- Loglar: `sprint_story_mailer.log`
- Şifre değişirse `config.json` içindeki `rtc.password` alanını güncelleyin.
