#!/usr/bin/env python3
"""
Sprint End Day -> RTC Story Mailer
==================================
Sprint bitiş günlerinde (2 haftada bir Salı) RTC'deki (IBM EWM) kayıtlı
story query'sini çalıştırır ve sonuçları HTML tablo olarak Outlook
üzerinden mail atar.

- Sprint bitiş günleri RTC'deki iterasyon (sprint) takviminden okunur.
  RTC'ye ulaşılamazsa config.json'daki yedek takvim (anchor + 14 gün) kullanılır.
- Görev Zamanlayıcı script'i HER GÜN çalıştırır. Script, son sprint bitiş
  gününün mailinin gönderilip gönderilmediğine bakar (sent_state.json).
  PC o gün kapalıysa / VPN yoksa mail sonraki günlerde telafi edilir
  (catch_up_days kadar).
- VPN / RTC bağlantısı kurulamazsa: Windows bildirimi + size bilgi maili.
  Belirli aralıklarla tekrar dener, sonunda yine başarısızsa tekrar bildirir.

Kullanım:
    python sprint_story_mailer.py                -> normal çalışma
    python sprint_story_mailer.py --force        -> tarih kontrolü olmadan hemen gönder
    python sprint_story_mailer.py --dry-run      -> mail atmaz, HTML önizleme yazar
    python sprint_story_mailer.py --list-sprints -> RTC'deki sprintleri ve bitiş günlerini listeler
    python sprint_story_mailer.py --test-notify  -> hata bildirimini (bildirim + mail) dener

Gereksinimler:
    pip install requests pywin32
"""
from __future__ import annotations

import argparse
import datetime as dt
import html
import json
import logging
import os
import re
import subprocess
import sys
import time
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any

import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"
STATE_PATH = BASE_DIR / "sent_state.json"
LOG_PATH = BASE_DIR / "sprint_story_mailer.log"

log = logging.getLogger("sprint_mailer")

SELECT = (
    "dcterms:identifier,dcterms:title,dcterms:type,"
    "oslc_cm:status,"
    "rtc_cm:ownedBy{foaf:name},"
    "dcterms:contributor{foaf:name},"
    "rtc_cm:plannedFor{dcterms:title}"
)
RTC_TRANSIENT = {500, 502, 503, 504}


class RtcConnectionError(RuntimeError):
    """VPN kopuk / sunucuya ulaşılamıyor / login olmuyor -> tekrar denenebilir."""


# ==============================================================================
#  Config / state
# ==============================================================================
def load_config(path: Path) -> dict[str, Any]:
    if not path.exists():
        sys.exit(f"config.json bulunamadı: {path}\n"
                 "config.example.json dosyasını config.json olarak kopyalayıp doldurun.")
    with path.open(encoding="utf-8") as f:
        return json.load(f)


def load_state() -> dict[str, Any]:
    try:
        return json.loads(STATE_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save_state(state: dict[str, Any]) -> None:
    STATE_PATH.write_text(json.dumps(state, indent=2), encoding="utf-8")


def rtc_credentials(cfg: dict[str, Any]) -> tuple[str, str]:
    """Öncelik: config.json (rtc.username / rtc.password) -> RTC_USER / RTC_PASS."""
    r = cfg["rtc"]
    user = (r.get("username") or os.environ.get("RTC_USER", "")).strip()
    password = r.get("password") or os.environ.get("RTC_PASS", "")
    if not user or not password:
        raise RuntimeError("RTC kullanıcı adı/şifre yok: config.json > rtc.username / rtc.password")
    return user, password


# ==============================================================================
#  Sprint takvimi
# ==============================================================================
def fallback_end_dates(cfg: dict[str, Any], today: dt.date) -> list[dt.date]:
    """config.json'daki yedek takvim: açık liste veya anchor + interval."""
    sprint = cfg["sprint"]
    if sprint.get("end_dates"):
        return sorted(dt.date.fromisoformat(d) for d in sprint["end_dates"])
    anchor = dt.date.fromisoformat(sprint["anchor_end_date"])
    step = dt.timedelta(days=int(sprint.get("interval_days", 14)))
    dates, d = [], anchor
    while d <= today + step:
        dates.append(d)
        d += step
    return dates


def due_sprint_end(end_dates: list[dt.date], today: dt.date, catch_up_days: int,
                   state: dict[str, Any]) -> dt.date | None:
    """Bugün mail gönderilmesi gereken sprint bitiş günü (yoksa None)."""
    past = [d for d in end_dates if d <= today]
    if not past:
        return None
    last = max(past)
    if (today - last).days > catch_up_days:
        return None
    if last.isoformat() in state.get("sent", []):
        return None
    return last


# ==============================================================================
#  RTC
# ==============================================================================
def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _child_text(el: ET.Element, name: str) -> str:
    for c in el:
        if _local(c.tag) == name:
            return (c.text or "").strip()
    return ""


class Rtc:
    def __init__(self, cfg: dict[str, Any]):
        self.cfg = cfg
        r = cfg["rtc"]
        self.base = r["base_url"].rstrip("/")
        self.proj_id = r["project_area_id"]
        self.gcm = r.get("gcm_config", "")
        self.session = requests.Session()
        self.headers = {"Accept": "application/json", "OSLC-Core-Version": "2.0"}
        if self.gcm:
            self.headers["oslc-config.context"] = self.gcm
        self._user_cache: dict[str, str] = {}

    # ── HTTP ──────────────────────────────────────────────────────────────────
    def get(self, url: str, params=None, headers=None, timeout=60, tries=3) -> requests.Response:
        for attempt in range(1, tries + 1):
            try:
                r = self.session.get(url, params=params, headers=headers or self.headers,
                                     verify=False, timeout=timeout)
                if r.status_code not in RTC_TRANSIENT or attempt == tries:
                    return r
            except (requests.Timeout, requests.ConnectionError) as exc:
                if attempt == tries:
                    raise RtcConnectionError(f"RTC sunucusuna ulaşılamıyor ({type(exc).__name__}) "
                                             "-- VPN bağlı mı?") from exc
            time.sleep(3 * attempt)
        raise AssertionError("unreachable")

    def login(self) -> None:
        user, password = rtc_credentials(self.cfg)
        try:
            self.session.get(f"{self.base}/authenticated/identity", verify=False, timeout=20)
            r = self.session.post(f"{self.base}/j_security_check",
                                  data={"j_username": user, "j_password": password},
                                  verify=False, allow_redirects=True, timeout=20)
        except (requests.Timeout, requests.ConnectionError) as exc:
            raise RtcConnectionError(f"RTC sunucusuna ulaşılamıyor ({type(exc).__name__}) "
                                     "-- VPN bağlı mı?") from exc
        if ("j_security_check" in r.url
                or r.headers.get("X-com-ibm-team-repository-web-auth-msg") == "authfailed"):
            # Şifre hatası tekrar denemekle düzelmez -> RuntimeError
            raise RuntimeError(f"RTC login başarısız: '{user}' kullanıcı adı/şifre hatalı "
                               "veya şifrenin süresi dolmuş.")
        check = self.get(f"{self.base}/oslc/contexts/{self.proj_id}/workitems",
                         params={"oslc.select": "dcterms:identifier", "oslc.pageSize": "1"},
                         timeout=30)
        try:
            check.json()
        except ValueError:
            raise RtcConnectionError("RTC login görünürde başarılı ama test çağrısı JSON "
                                     "dönmedi -- VPN bağlantısını kontrol edin.") from None
        log.info("RTC girişi başarılı: %s", user)

    # ── Sprint takvimi (Process REST API) ────────────────────────────────────
    def sprints(self) -> list[dict[str, Any]]:
        """Proje alanının tüm timeline'larındaki yaprak iterasyonlar (sprintler)."""
        sc = self.cfg["sprint"]
        pa = sc.get("project_area_uuid") or self.proj_id
        name_re = re.compile(sc.get("name_filter") or ".*", re.I)
        offset = dt.timedelta(days=int(sc.get("end_date_offset_days", 0)))
        xml_headers = {"Accept": "application/xml"}

        def fetch(url: str) -> ET.Element:
            r = self.get(url, headers=xml_headers, timeout=30)
            if not r.ok:
                raise RuntimeError(f"Sprint takvimi okunamadı: HTTP {r.status_code} ({url})")
            return ET.fromstring(r.content)

        out: list[dict[str, Any]] = []

        def walk(iter_list_url: str, timeline: str, path: str) -> None:
            for it in fetch(iter_list_url):
                if _local(it.tag) != "iteration":
                    continue
                name = _child_text(it, "name")
                child_url = _child_text(it, "iterations-url")
                children = fetch(child_url) if child_url else None
                has_children = children is not None and any(
                    _local(c.tag) == "iteration" for c in children)
                full = f"{path}/{name}" if path else name
                if has_children:
                    walk(child_url, timeline, full)
                    continue
                end_raw = _child_text(it, "end-date")
                if not end_raw or not name_re.search(full):
                    continue
                end = dt.datetime.fromisoformat(end_raw.replace("Z", "+00:00"))
                last_day = (end.astimezone().date() if end.tzinfo else end.date()) + offset
                out.append({"timeline": timeline, "name": full, "end_raw": end_raw,
                            "last_day": last_day,
                            "start_raw": _child_text(it, "start-date")})

        root = fetch(f"{self.base}/process/project-areas/{pa}/timelines")
        wanted_tl = (sc.get("timeline_name") or "").lower()
        for tl in root:
            if _local(tl.tag) != "development-line":
                continue
            tl_name = _child_text(tl, "name")
            if wanted_tl and tl_name.lower() != wanted_tl:
                continue
            url = _child_text(tl, "iterations-url")
            if url:
                walk(url, tl_name, "")
        return sorted(out, key=lambda s: s["last_day"])

    # ── Story query ──────────────────────────────────────────────────────────
    def run_query(self, q: dict[str, Any]) -> list[dict[str, Any]]:
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
                raise RtcConnectionError(f"Query '{q['name']}' sayfa {page}: JSON dönmedi "
                                         "(oturum düştü -- VPN?)") from None
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
#  Mail + bildirim
# ==============================================================================
COLUMNS = [("ID", "id"), ("Başlık", "title"), ("Durum", "status"),
           ("Sahibi", "owner"), ("Sprint", "planned_for")]


def build_html(results: list[tuple[str, list[dict]]], sprint_end: dt.date) -> str:
    th = "background:#1f4e79;color:#fff;padding:6px 10px;text-align:left;border:1px solid #1f4e79;"
    td = "padding:5px 10px;border:1px solid #d0d0d0;"
    out = [
        "<html><body style='font-family:Segoe UI,Arial,sans-serif;font-size:13px;'>",
        "<p>Merhaba,</p>",
        f"<p>{sprint_end:%d.%m.%Y} sprint bitiş günü itibarıyla RTC query sonuçları aşağıdadır.</p>",
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


def send_outlook_mail(to: list[str], cc: list[str], subject: str, body: str,
                      high_importance: bool = False) -> None:
    import win32com.client  # pywin32 -- sadece Windows

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # olMailItem
    mail.To = "; ".join(to)
    if cc:
        mail.CC = "; ".join(cc)
    mail.Subject = subject
    mail.HTMLBody = body
    if high_importance:
        mail.Importance = 2  # olImportanceHigh
    mail.Send()
    log.info("Mail gönderildi -> %s (%s)", "; ".join(to), subject)


def windows_toast(title: str, message: str) -> None:
    """Ek modül gerektirmeyen Windows bildirim balonu (PowerShell)."""
    if os.name != "nt":
        return
    t = title.replace("'", "''")
    m = message.replace("'", "''")
    ps = ("Add-Type -AssemblyName System.Windows.Forms;"
          "$n=New-Object System.Windows.Forms.NotifyIcon;"
          "$n.Icon=[System.Drawing.SystemIcons]::Warning;$n.Visible=$true;"
          f"$n.ShowBalloonTip(30000,'{t}','{m}',[System.Windows.Forms.ToolTipIcon]::Warning);"
          "Start-Sleep -Seconds 30;$n.Dispose()")
    try:
        subprocess.Popen(["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps],
                         creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except Exception:
        log.exception("Windows bildirimi gösterilemedi")


def notify(cfg: dict[str, Any], title: str, message: str) -> None:
    """Hata bildirimi: Windows bildirimi + bilgi maili (Outlook VPN olmadan da gönderebilir)."""
    log.warning("BİLDİRİM: %s - %s", title, message)
    windows_toast(title, message)
    to = cfg["mail"].get("error_notify_to") or []
    if not to:
        return
    body = ("<html><body style='font-family:Segoe UI,Arial,sans-serif;font-size:13px;'>"
            f"<p><b>{html.escape(title)}</b></p><p>{html.escape(message)}</p>"
            f"<p style='color:#888;font-size:11px'>Log: {html.escape(str(LOG_PATH))}</p>"
            "</body></html>")
    try:
        send_outlook_mail(to, [], f"[UYARI] RTC Sprint Mailer: {title}", body,
                          high_importance=True)
    except Exception:
        log.exception("Bildirim maili gönderilemedi")


# ==============================================================================
#  Main
# ==============================================================================
def connect_with_retry(cfg: dict[str, Any], rtc: Rtc, must_succeed: bool) -> bool:
    """
    must_succeed=False: tek deneme, olmazsa sessizce False (bugün mail günü değilse).
    must_succeed=True : bildirerek tekrar dener; hepsi başarısızsa RtcConnectionError.
    """
    retry = cfg.get("retry", {})
    attempts = int(retry.get("attempts", 8)) if must_succeed else 1
    wait_min = int(retry.get("wait_minutes", 15))
    for i in range(1, attempts + 1):
        try:
            rtc.login()
            if i > 1:
                log.info("Bağlantı %d. denemede kuruldu.", i)
            return True
        except RtcConnectionError as exc:
            log.warning("Bağlantı denemesi %d/%d başarısız: %s", i, attempts, exc)
            if i == attempts:
                if must_succeed:
                    raise
                return False
            if i == 1:
                notify(cfg, "RTC'ye bağlanılamadı",
                       f"{exc} Sprint sonu maili gönderilemedi. VPN'e bağlanın; "
                       f"{wait_min} dakikada bir tekrar denenecek "
                       f"(toplam {attempts} deneme).")
            time.sleep(wait_min * 60)
    return False


def run(cfg: dict[str, Any], args: argparse.Namespace) -> int:
    today = dt.date.today()
    state = load_state()
    catch_up = int(cfg["sprint"].get("catch_up_days", 3))
    rtc = Rtc(cfg)

    # Yedek takvime göre bugün mail günü mü? (RTC'ye bağlanmadan önce bilinmesi gerekiyor:
    # bağlantı yoksa sadece mail günüyse kullanıcıyı rahatsız ederiz.)
    local_due = due_sprint_end(fallback_end_dates(cfg, today), today, catch_up, state)
    must = bool(args.force or args.dry_run or local_due)

    if not connect_with_retry(cfg, rtc, must_succeed=must):
        log.info("RTC'ye bağlanılamadı ama bugün mail günü değil, çıkılıyor.")
        return 0

    # Asıl karar: RTC'deki sprint takvimi
    due = local_due
    if cfg["sprint"].get("source", "rtc") == "rtc":
        try:
            ends = sorted({s["last_day"] for s in rtc.sprints()})
            if ends:
                due = due_sprint_end(ends, today, catch_up, state)
            else:
                log.warning("RTC'de sprint bulunamadı, yedek takvim kullanılıyor.")
        except RtcConnectionError:
            raise
        except Exception as exc:
            log.warning("RTC sprint takvimi okunamadı (%s), yedek takvim kullanılıyor.", exc)

    if args.force or args.dry_run:
        due = due or today
    if not due:
        log.info("%s: gönderilecek sprint sonu maili yok.", today)
        return 0

    results = [(q["name"], rtc.run_query(q)) for q in cfg["queries"]]
    body = build_html(results, due)
    mail_cfg = cfg["mail"]
    subject = mail_cfg.get("subject", "Sprint Sonu Story Listesi - {date}").format(
        date=due.strftime("%d.%m.%Y"))

    if args.dry_run:
        out = BASE_DIR / f"preview_{due}.html"
        out.write_text(body, encoding="utf-8")
        log.info("Dry-run: mail gönderilmedi. Önizleme: %s", out)
        return 0

    send_outlook_mail(mail_cfg["to"], mail_cfg.get("cc", []), subject, body)
    state.setdefault("sent", []).append(due.isoformat())
    state["sent"] = state["sent"][-50:]
    save_state(state)
    return 0


def list_sprints(cfg: dict[str, Any]) -> int:
    rtc = Rtc(cfg)
    rtc.login()
    sprints = rtc.sprints()
    print(f"{'Son gün':<12} {'end-date (RTC ham)':<30} Timeline / Sprint")
    for s in sprints:
        print(f"{s['last_day'].isoformat():<12} {s['end_raw']:<30} {s['timeline']} / {s['name']}")
    print(f"\nToplam {len(sprints)} sprint.")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="Sprint end day RTC story mailer")
    ap.add_argument("--force", action="store_true", help="Tarih kontrolü olmadan gönder")
    ap.add_argument("--dry-run", action="store_true", help="Mail atma, HTML önizleme yaz")
    ap.add_argument("--list-sprints", action="store_true", help="RTC sprintlerini listele")
    ap.add_argument("--test-notify", action="store_true", help="Hata bildirimini dene")
    ap.add_argument("--config", default=str(CONFIG_PATH))
    args = ap.parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.FileHandler(LOG_PATH, encoding="utf-8"), logging.StreamHandler()],
    )
    cfg = load_config(Path(args.config))

    if args.test_notify:
        notify(cfg, "Test bildirimi", "Bu bir test bildirimidir, her şey yolunda.")
        return 0
    if args.list_sprints:
        return list_sprints(cfg)

    try:
        return run(cfg, args)
    except RtcConnectionError as exc:
        notify(cfg, "RTC bağlantısı kurulamadı",
               f"{exc} Tüm denemeler başarısız oldu, sprint sonu maili GÖNDERİLEMEDİ. "
               "VPN'e bağlandıktan sonra script yarın otomatik tekrar deneyecek; "
               "hemen göndermek için: python sprint_story_mailer.py")
        return 1
    except Exception as exc:
        log.exception("Hata oluştu")
        notify(cfg, "Sprint mailer hatası",
               f"{type(exc).__name__}: {exc} Sprint sonu maili gönderilemedi.")
        return 1


if __name__ == "__main__":
    sys.exit(main())
