#!/usr/bin/env python3
"""
Sprint End Day -> RTC Story Mailer
==================================
Sprint bitiş günlerinde (2 haftada bir Salı) RTC'de (IBM EWM) tanımlı
query'leri çalıştırır ve çıkan story'leri HTML tablo olarak Outlook
üzerinden mail atar.

RTC bağlantısı release_sync_v4.py ile aynı yöntemi kullanır:
j_security_check form login + OSLC 2.0 JSON + Global Configuration context.

Windows Görev Zamanlayıcı script'i HER SALI çalıştırır; script bugünün
sprint bitiş günü olup olmadığına config.json'daki takvime göre kendisi
karar verir (sprint günü değilse hiçbir şey yapmadan çıkar).

Kullanım:
    python sprint_story_mailer.py              -> normal (sprint günü değilse çıkar)
    python sprint_story_mailer.py --force      -> tarih kontrolü yapmadan çalıştır
    python sprint_story_mailer.py --dry-run    -> mail atmaz, HTML önizleme dosyası yazar
    python sprint_story_mailer.py --check-date -> sadece bugünün sprint günü olup olmadığını yazar

Gereksinimler:
    pip install requests pywin32
    RTC_USER / RTC_PASS ortam değişkenleri (zamanlanmış görev şifre soramaz):
        setx RTC_USER <kullanici>
        setx RTC_PASS <sifre>
"""
from __future__ import annotations

import argparse
import datetime as dt
import html
import json
import logging
import os
import re
import sys
import time
from pathlib import Path
from typing import Any

import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"
LOG_PATH = BASE_DIR / "sprint_story_mailer.log"

log = logging.getLogger("sprint_mailer")

# Sonuçlarda istenen alanlar (OSLC 2.0)
SELECT = (
    "dcterms:identifier,dcterms:title,dcterms:type,"
    "oslc_cm:status,"
    "rtc_cm:ownedBy{foaf:name},"
    "dcterms:contributor{foaf:name},"
    "rtc_cm:plannedFor{dcterms:title}"
)
RTC_TRANSIENT = {500, 502, 503, 504}


# ==============================================================================
#  Config
# ==============================================================================
def load_config(path: Path) -> dict[str, Any]:
    if not path.exists():
        sys.exit(f"config.json bulunamadı: {path}\n"
                 "config.example.json dosyasını config.json olarak kopyalayıp doldurun.")
    with path.open(encoding="utf-8") as f:
        return json.load(f)


def rtc_credentials() -> tuple[str, str]:
    """Zamanlanmış görevde kimse şifre giremez -> önce ortam değişkeni, sonra keyring."""
    user = os.environ.get("RTC_USER", "").strip()
    password = os.environ.get("RTC_PASS", "")
    if user and not password:
        try:
            import keyring  # opsiyonel: Windows Credential Manager
            password = keyring.get_password("RTC", user) or ""
        except ImportError:
            pass
    if not user or not password:
        raise RuntimeError("RTC_USER / RTC_PASS tanımlı değil (README > Kurulum adım 3).")
    return user, password


# ==============================================================================
#  Sprint bitiş günü kontrolü
# ==============================================================================
def is_sprint_end_day(cfg: dict[str, Any], today: dt.date) -> bool:
    sprint = cfg["sprint"]

    # 1) Açık tarih listesi varsa sadece ona bakılır (tatil / kayma durumları için)
    explicit = sprint.get("end_dates") or []
    if explicit:
        return today.isoformat() in explicit

    # 2) Referans sprint bitiş tarihi + periyot (varsayılan 14 gün)
    anchor = dt.date.fromisoformat(sprint["anchor_end_date"])
    interval = int(sprint.get("interval_days", 14))
    diff = (today - anchor).days
    return diff >= 0 and diff % interval == 0


# ==============================================================================
#  RTC
# ==============================================================================
class Rtc:
    def __init__(self, cfg: dict[str, Any]):
        r = cfg["rtc"]
        self.base = r["base_url"].rstrip("/")          # https://prsse.intra.chrysler.com/ccm
        self.proj_id = r["project_area_id"]
        self.gcm = r.get("gcm_config", "")
        self.session = requests.Session()
        self.headers = {"Accept": "application/json", "OSLC-Core-Version": "2.0"}
        if self.gcm:
            self.headers["oslc-config.context"] = self.gcm
        self._user_cache: dict[str, str] = {}

    # ── HTTP ──────────────────────────────────────────────────────────────────
    def get(self, url: str, params=None, timeout=60, tries=3) -> requests.Response:
        for attempt in range(1, tries + 1):
            try:
                r = self.session.get(url, params=params, headers=self.headers,
                                     verify=False, timeout=timeout)
                if r.status_code not in RTC_TRANSIENT or attempt == tries:
                    return r
            except (requests.Timeout, requests.ConnectionError):
                if attempt == tries:
                    raise
            time.sleep(3 * attempt)
        raise AssertionError("unreachable")

    def login(self) -> None:
        user, password = rtc_credentials()
        self.session.get(f"{self.base}/authenticated/identity", verify=False, timeout=20)
        r = self.session.post(f"{self.base}/j_security_check",
                              data={"j_username": user, "j_password": password},
                              verify=False, allow_redirects=True, timeout=20)
        if ("j_security_check" in r.url
                or r.headers.get("X-com-ibm-team-repository-web-auth-msg") == "authfailed"):
            raise RuntimeError(f"RTC login başarısız: '{user}' kullanıcı adı/şifre kontrol edin.")
        # Gerçek bir API çağrısıyla doğrula (VPN kopuk, şifre süresi dolmuş vb.)
        check = self.get(f"{self.base}/oslc/contexts/{self.proj_id}/workitems",
                         params={"oslc.select": "dcterms:identifier", "oslc.pageSize": "1"},
                         timeout=30)
        try:
            check.json()
        except ValueError:
            raise RuntimeError("RTC login görünürde başarılı ama test çağrısı JSON dönmedi "
                               "-- VPN / şifre durumunu kontrol edin.") from None
        log.info("RTC girişi başarılı: %s", user)

    # ── Query ─────────────────────────────────────────────────────────────────
    def run_query(self, q: dict[str, Any]) -> list[dict[str, Any]]:
        """
        İki tip query desteklenir:
          "saved_query_id": RTC web arayüzünde kayıtlı query'nin ID'si (_xxxxx)
          "where":          OSLC where ifadesi (proje alanı içinde çalışır)
        """
        params: dict[str, str] = {"oslc.select": SELECT, "oslc.pageSize": "100"}
        if self.gcm:
            params["oslc_config.context"] = self.gcm

        if q.get("saved_query_id"):
            url = f"{self.base}/oslc/queries/{q['saved_query_id']}/rtc_cm:results"
        elif q.get("where"):
            url = f"{self.base}/oslc/contexts/{self.proj_id}/workitems"
            params["oslc.where"] = q["where"]
        else:
            raise ValueError(f"Query '{q.get('name')}' için saved_query_id veya where gerekli.")

        items, seen = [], set()
        for page in range(1, 1001):
            r = self.get(url, params=params if page == 1 else None)
            if not r.ok:
                raise RuntimeError(f"Query '{q['name']}' sayfa {page}: HTTP {r.status_code}")
            try:
                data = r.json()
            except ValueError:
                raise RuntimeError(f"Query '{q['name']}' sayfa {page}: JSON dönmedi "
                                   "(oturum düşmüş olabilir -- VPN?)") from None

            members = data.get("oslc:results") or data.get("rdfs:member") or []
            new = 0
            for m in members:
                mid = str(m.get("dcterms:identifier", ""))
                if mid and mid not in seen:
                    seen.add(mid)
                    items.append(m)
                    new += 1

            next_url = (r.links.get("next", {}).get("url")
                        or data.get("oslc:responseInfo", {}).get("oslc:nextPage")
                        or data.get("oslc:nextPage"))
            if not next_url or new == 0:
                break
            url = next_url

        if q.get("only_stories", True):
            items = [m for m in items if _plain(m.get("dcterms:type")).lower() == "story"]
        log.info("Query '%s' -> %d kayıt", q["name"], len(items))
        return [self._row(m) for m in items]

    # ── Alan dönüştürme ───────────────────────────────────────────────────────
    def _owner(self, m: dict) -> str:
        raw = _first_dict(m.get("dcterms:contributor")) or _first_dict(m.get("rtc_cm:ownedBy"))
        name = raw.get("foaf:name", "")
        uri = raw.get("rdf:resource", "")
        if not name and uri:
            name = self._resolve_user(uri)
        return clean_text(name) or "Unassigned"

    def _resolve_user(self, uri: str) -> str:
        if "unassigned" in uri.lower():
            return "Unassigned"
        if uri not in self._user_cache:
            name = ""
            try:
                r = self.session.get(uri, headers={"Accept": "application/json"},
                                     verify=False, timeout=10)
                if r.ok:
                    d = r.json()
                    name = d.get("foaf:name") or d.get("dcterms:title") or ""
            except Exception:
                pass
            self._user_cache[uri] = str(name).strip() or uri.rstrip("/").split("/")[-1]
        return self._user_cache[uri]

    def _row(self, m: dict) -> dict[str, str]:
        wid = str(m.get("dcterms:identifier", ""))
        return {
            "id": wid,
            "title": clean_text(m.get("dcterms:title", "")),
            "status": _plain(m.get("oslc_cm:status")),
            "owner": self._owner(m),
            "planned_for": _plain(m.get("rtc_cm:plannedFor")),
            "url": f"{self.base}/resource/itemName/com.ibm.team.workitem.WorkItem/{wid}",
        }


def clean_text(s: Any) -> str:
    s = html.unescape(str(s or ""))
    s = re.sub(r"<[^>]+>", " ", s)
    return " ".join(s.split())


def _first_dict(v):
    if isinstance(v, list):
        return v[0] if v else {}
    return v if isinstance(v, dict) else {}


def _plain(val) -> str:
    if val is None:
        return ""
    if isinstance(val, list):
        return ", ".join(p for p in (_plain(v) for v in val) if p)
    if isinstance(val, dict):
        for key in ("dcterms:title", "foaf:name"):
            if val.get(key):
                return clean_text(val[key])
        uri = val.get("rdf:resource", "")
        return uri.rstrip("/").split("/")[-1] if uri else ""
    return clean_text(val)


# ==============================================================================
#  Mail
# ==============================================================================
COLUMNS = [("ID", "id"), ("Başlık", "title"), ("Durum", "status"),
           ("Sahibi", "owner"), ("Sprint", "planned_for")]


def build_html(results: list[tuple[str, list[dict]]], today: dt.date) -> str:
    th = "background:#1f4e79;color:#fff;padding:6px 10px;text-align:left;border:1px solid #1f4e79;"
    td = "padding:5px 10px;border:1px solid #d0d0d0;"
    out = [
        "<html><body style='font-family:Segoe UI,Arial,sans-serif;font-size:13px;'>",
        "<p>Merhaba,</p>",
        f"<p>{today:%d.%m.%Y} sprint bitiş günü itibarıyla RTC query sonuçları aşağıdadır.</p>",
    ]
    for name, rows in results:
        out.append(f"<h3 style='color:#1f4e79;margin-bottom:6px'>"
                   f"{html.escape(name)} ({len(rows)} story)</h3>")
        if not rows:
            out.append("<p><i>Kayıt bulunamadı.</i></p>")
            continue
        out.append("<table style='border-collapse:collapse;'><tr>")
        out += [f"<th style='{th}'>{c}</th>" for c, _ in COLUMNS]
        out.append("</tr>")
        for i, row in enumerate(rows):
            bg = "#f5f8fc" if i % 2 else "#ffffff"
            out.append(f"<tr style='background:{bg}'>")
            for _, key in COLUMNS:
                val = html.escape(row.get(key, ""))
                if key == "id":
                    val = f"<a href='{html.escape(row['url'])}'>{val}</a>"
                out.append(f"<td style='{td}'>{val}</td>")
            out.append("</tr>")
        out.append("</table>")
    out.append("<p style='color:#888;font-size:11px;margin-top:20px'>"
               "Bu mail otomatik olarak gönderilmiştir.</p></body></html>")
    return "".join(out)


def send_outlook_mail(to: list[str], cc: list[str], subject: str, body: str) -> None:
    import win32com.client  # pywin32 -- sadece Windows

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # olMailItem
    mail.To = "; ".join(to)
    if cc:
        mail.CC = "; ".join(cc)
    mail.Subject = subject
    mail.HTMLBody = body
    mail.Send()
    log.info("Mail gönderildi -> %s", "; ".join(to))


# ==============================================================================
#  Main
# ==============================================================================
def main() -> int:
    ap = argparse.ArgumentParser(description="Sprint end day RTC story mailer")
    ap.add_argument("--force", action="store_true", help="Sprint günü kontrolünü atla")
    ap.add_argument("--dry-run", action="store_true", help="Mail atma, HTML önizleme yaz")
    ap.add_argument("--check-date", action="store_true", help="Sadece tarih kontrolü yap")
    ap.add_argument("--config", default=str(CONFIG_PATH))
    args = ap.parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.FileHandler(LOG_PATH, encoding="utf-8"), logging.StreamHandler()],
    )
    cfg = load_config(Path(args.config))
    today = dt.date.today()

    sprint_day = is_sprint_end_day(cfg, today)
    if args.check_date:
        print(f"{today}: {'SPRINT BİTİŞ GÜNÜ' if sprint_day else 'sprint bitiş günü değil'}")
        return 0
    if not sprint_day and not args.force:
        log.info("%s sprint bitiş günü değil, çıkılıyor.", today)
        return 0

    try:
        rtc = Rtc(cfg)
        rtc.login()
        results = [(q["name"], rtc.run_query(q)) for q in cfg["queries"]]

        body = build_html(results, today)
        mail_cfg = cfg["mail"]
        subject = mail_cfg.get("subject", "Sprint Sonu Story Listesi - {date}").format(
            date=today.strftime("%d.%m.%Y"))

        if args.dry_run:
            out = BASE_DIR / f"preview_{today}.html"
            out.write_text(body, encoding="utf-8")
            log.info("Dry-run: mail gönderilmedi. Önizleme: %s", out)
        else:
            send_outlook_mail(mail_cfg["to"], mail_cfg.get("cc", []), subject, body)
        return 0
    except Exception:
        log.exception("Hata oluştu")
        return 1


if __name__ == "__main__":
    sys.exit(main())
