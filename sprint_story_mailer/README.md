# RTC Sprint Story Mailer

Sprint bitiş günlerinde (2 haftada bir Salı) RTC'deki query'leri çalıştırır ve
çıkan story'leri Outlook üzerinden şu adreslere mail atar:

- sf01089@external.stellantis.com
- t59612@tofas.com.tr

RTC bağlantısı `release_sync_v4.py` ile aynı yöntemi kullanır (`j_security_check`
login, OSLC 2.0 JSON, GCM context).

## Neden PC'de çalışıyor (bulutta değil)?

RTC (`prsse.intra.chrysler.com`) şirket içi ağda / VPN arkasında. Claude'un bulut
ortamı ve GitHub Actions bu sunucuya erişemez; Outlook da sadece PC'de var. Bu yüzden
script PC'de **Windows Görev Zamanlayıcı** ile kendi kendine çalışır.

## Kurulum

1. Paketleri kurun:
   ```
   pip install -r requirements.txt
   ```
2. `config.example.json` dosyasını `config.json` olarak kopyalayın ve doldurun:
   - `sprint.anchor_end_date`: Bilinen bir sprint bitiş Salı'sı (YYYY-MM-DD).
     Script bu tarihten itibaren her 14 günde bir mail atar.
   - `sprint.end_dates`: (opsiyonel) Sprint günleri kayarsa tarihleri buraya liste
     olarak yazın, örn. `["2026-11-03", "2026-11-17"]`. Liste doluysa sadece
     bu tarihlerde mail gider, `anchor_end_date` dikkate alınmaz.
   - `queries`: Her query için ya
     - `saved_query_id` (RTC web'de kayıtlı query'yi çalıştırınca URL'deki
       `id=_xxxxx` değeri), ya da
     - `where` (OSLC where ifadesi, örn. `rtc_cm:plannedFor="<iteration URI>"`).
     `only_stories: true` sadece Story tipindekileri listeler.
3. RTC kullanıcı adı/şifresini ortam değişkeni olarak kaydedin (zamanlanmış görev
   şifre soramaz). Yeni bir terminal açın:
   ```
   setx RTC_USER kullanici_adiniz
   setx RTC_PASS sifreniz
   ```
4. Test edin:
   ```
   python sprint_story_mailer.py --check-date        (bugün sprint günü mü?)
   python sprint_story_mailer.py --force --dry-run   (mail atmadan HTML önizleme üretir)
   python sprint_story_mailer.py --force             (gerçekten mail atar)
   ```
5. Otomatik çalıştırma: `create_scheduled_task.bat` dosyasına çift tıklayın.
   Her Salı 16:00'da çalışan bir görev oluşturur; sprint günü değilse script
   hiçbir şey yapmadan çıkar.

## Notlar

- PC o saatte açık, oturum açık ve VPN bağlı olmalı (Outlook ve RTC için).
- Loglar `sprint_story_mailer.log` dosyasına yazılır.
- Şifre değişirse `setx RTC_PASS ...` komutunu tekrar çalıştırın.
