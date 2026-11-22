#!/usr/bin/env python3
"""
Ежедневный мониторинг sellers.json рекламных сетей.

Таблицы (SQLite, файл sellers.db):
  all_sites      — полный текущий список сайтов каждой сети (база для сравнения)
  added_sites    — новые сайты: дата добавления, сеть, домен, владелец
  removed_sites  — удалённые сайты: дата удаления, сеть, домен, владелец
  run_log        — журнал запусков по сетям

Первый запуск по сети = базовая загрузка: всё идёт только в all_sites,
в added_sites ничего не пишется. Дальше — только разница день к дню.

Экспорт:
  sellers_changes.xlsx — листы «Новые», «Удалённые», «Статус сетей» (лёгкий)
  all_sites.xlsx       — полный список (отдельный тяжёлый файл, по флагу --export-all)

Запуск:  python monitor.py              # проверка + экспорт изменений
         python monitor.py --export-all  # плюс выгрузка полного списка
         python monitor.py --check       # только проверить, какие URL отвечают
"""
import argparse
import datetime as dt
import json
import os
import sqlite3
import sys
import time
from pathlib import Path

import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
DB_PATH = BASE / "sellers.db"
CONFIG_PATH = BASE / "networks.json"
CHANGES_XLSX = BASE / "sellers_changes.xlsx"
ALL_XLSX = BASE / "all_sites.xlsx"
TEMPLATE = BASE / "dashboard_template.html"
DASHBOARD = BASE / "docs" / "index.html"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; SellersMonitor/1.0)",
    "Accept": "application/json,*/*",
}

SCHEMA = """
CREATE TABLE IF NOT EXISTS all_sites (
    network     TEXT NOT NULL,
    domain      TEXT NOT NULL,
    name        TEXT,
    seller_id   TEXT,
    seller_type TEXT,
    first_seen  TEXT NOT NULL,
    PRIMARY KEY (network, domain)
);
CREATE TABLE IF NOT EXISTS added_sites (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    date_added  TEXT NOT NULL,
    network     TEXT NOT NULL,
    domain      TEXT NOT NULL,
    name        TEXT,
    seller_id   TEXT,
    seller_type TEXT
);
CREATE TABLE IF NOT EXISTS removed_sites (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    date_removed TEXT NOT NULL,
    network      TEXT NOT NULL,
    domain       TEXT NOT NULL,
    name         TEXT,
    seller_id    TEXT,
    seller_type  TEXT
);
CREATE TABLE IF NOT EXISTS run_log (
    run_date  TEXT NOT NULL,
    network   TEXT NOT NULL,
    status    TEXT NOT NULL,
    url       TEXT,
    total     INTEGER,
    added     INTEGER,
    removed   INTEGER,
    message   TEXT
);
CREATE INDEX IF NOT EXISTS idx_added_date ON added_sites(date_added);
CREATE INDEX IF NOT EXISTS idx_removed_date ON removed_sites(date_removed);
"""


def log(msg):
    print(f"[{dt.datetime.now():%H:%M:%S}] {msg}", flush=True)


def normalize_domain(d):
    if not d:
        return ""
    d = str(d).strip().lower()
    for p in ("https://", "http://"):
        if d.startswith(p):
            d = d[len(p):]
    d = d.split("/")[0].split("?")[0].strip(".")
    if d.startswith("www."):
        d = d[4:]
    return d


def fetch_sellers(urls, retries=3):
    """Пробует URL по очереди, возвращает (url, список sellers) или бросает ошибку."""
    last_err = None
    for url in urls:
        for attempt in range(1, retries + 1):
            try:
                r = requests.get(url, headers=HEADERS, timeout=60)
                r.raise_for_status()
                data = json.loads(r.content.decode("utf-8-sig"))
                sellers = data.get("sellers") if isinstance(data, dict) else None
                if not isinstance(sellers, list):
                    raise ValueError("в файле нет массива 'sellers'")
                return url, sellers
            except Exception as e:  # noqa: BLE001
                last_err = f"{url}: {e}"
                if attempt < retries:
                    time.sleep(3 * attempt)
    raise RuntimeError(last_err or "нет URL")


def parse_sellers(sellers, allowed_types):
    """Возвращает dict domain -> запись. Конфиденциальные и без домена пропускаются."""
    allowed = {t.upper() for t in allowed_types} if allowed_types else None
    out = {}
    for s in sellers:
        if not isinstance(s, dict):
            continue
        if str(s.get("is_confidential", 0)) in ("1", "True", "true"):
            continue
        stype = str(s.get("seller_type", "")).upper()
        if allowed and stype not in allowed:
            continue
        domain = normalize_domain(s.get("domain"))
        if not domain:
            continue
        if domain not in out:  # первое вхождение домена
            out[domain] = {
                "domain": domain,
                "name": (s.get("name") or "").strip(),
                "seller_id": str(s.get("seller_id", "")),
                "seller_type": stype,
            }
    return out


def process_network(conn, net, today, cfg):
    name = net["name"]
    try:
        url, raw = fetch_sellers(net["urls"])
    except Exception as e:  # noqa: BLE001
        log(f"  ✗ {name}: не удалось скачать — {e}")
        conn.execute("INSERT INTO run_log VALUES (?,?,?,?,?,?,?,?)",
                     (today, name, "ERROR", None, None, 0, 0, str(e)[:500]))
        return

    fresh = parse_sellers(raw, cfg.get("include_seller_types"))
    old = {row[0]: row for row in conn.execute(
        "SELECT domain, name, seller_id, seller_type FROM all_sites WHERE network=?", (name,))}

    if not fresh:
        log(f"  ! {name}: файл пустой после фильтра — пропускаю, базу не трогаю")
        conn.execute("INSERT INTO run_log VALUES (?,?,?,?,?,?,?,?)",
                     (today, name, "EMPTY", url, 0, 0, 0, "пустой список"))
        return

    # Первая загрузка — только в all_sites
    if not old:
        conn.executemany(
            "INSERT OR IGNORE INTO all_sites VALUES (?,?,?,?,?,?)",
            [(name, r["domain"], r["name"], r["seller_id"], r["seller_type"], today)
             for r in fresh.values()])
        log(f"  ✓ {name}: базовая загрузка {len(fresh)} сайтов")
        conn.execute("INSERT INTO run_log VALUES (?,?,?,?,?,?,?,?)",
                     (today, name, "BASELINE", url, len(fresh), 0, 0, None))
        return

    added = [fresh[d] for d in fresh.keys() - old.keys()]
    removed = [old[d] for d in old.keys() - fresh.keys()]

    # Защита от сбоя: если исчезла слишком большая доля сайтов — не записываем удаления
    ratio = len(removed) / max(len(old), 1)
    if ratio > cfg.get("max_removal_ratio", 0.3):
        msg = f"исчезло {len(removed)} из {len(old)} ({ratio:.0%}) — похоже на сбой файла, пропускаю"
        log(f"  ! {name}: {msg}")
        conn.execute("INSERT INTO run_log VALUES (?,?,?,?,?,?,?,?)",
                     (today, name, "SUSPICIOUS", url, len(fresh), 0, 0, msg))
        return

    conn.executemany(
        "INSERT INTO added_sites (date_added, network, domain, name, seller_id, seller_type) VALUES (?,?,?,?,?,?)",
        [(today, name, r["domain"], r["name"], r["seller_id"], r["seller_type"]) for r in added])
    conn.executemany(
        "INSERT OR IGNORE INTO all_sites VALUES (?,?,?,?,?,?)",
        [(name, r["domain"], r["name"], r["seller_id"], r["seller_type"], today) for r in added])

    conn.executemany(
        "INSERT INTO removed_sites (date_removed, network, domain, name, seller_id, seller_type) VALUES (?,?,?,?,?,?)",
        [(today, name, r[0], r[1], r[2], r[3]) for r in removed])
    conn.executemany("DELETE FROM all_sites WHERE network=? AND domain=?",
                     [(name, r[0]) for r in removed])

    log(f"  ✓ {name}: всего {len(fresh)}, новых {len(added)}, удалено {len(removed)}")
    conn.execute("INSERT INTO run_log VALUES (?,?,?,?,?,?,?,?)",
                 (today, name, "OK", url, len(fresh), len(added), len(removed), None))


# ---------- Трафик через DataForSEO ----------

DFS_URL = "https://api.dataforseo.com/v3/dataforseo_labs/google/bulk_traffic_estimation/live"


def migrate(conn):
    """Добавляет колонки трафика в старую базу, если их ещё нет."""
    cols = {r[1] for r in conn.execute("PRAGMA table_info(added_sites)")}
    for col, typ in (("traffic", "REAL"), ("keywords", "INTEGER"), ("traffic_date", "TEXT"),
                     ("domain_created", "TEXT"), ("whois_date", "TEXT")):
        if col not in cols:
            conn.execute(f"ALTER TABLE added_sites ADD COLUMN {col} {typ}")
    conn.commit()


def enrich_traffic(conn, cfg, today):
    """Подтягивает оценку органического трафика для новых сайтов, у которых его ещё нет."""
    tcfg = cfg.get("traffic", {})
    if not tcfg.get("enabled", True):
        return
    login, password = os.environ.get("DATAFORSEO_LOGIN"), os.environ.get("DATAFORSEO_PASSWORD")
    if not (login and password):
        log("Трафик: ключи DataForSEO не заданы — пропускаю")
        return

    max_domains = int(tcfg.get("max_domains_per_run", 3000))
    domains = [r[0] for r in conn.execute(
        "SELECT DISTINCT domain FROM added_sites WHERE traffic_date IS NULL "
        "ORDER BY date_added DESC LIMIT ?", (max_domains,))]
    if not domains:
        log("Трафик: новых сайтов без данных нет")
        return

    base = {"item_types": ["organic"]}
    if tcfg.get("location_code"):
        base["location_code"] = tcfg["location_code"]
    if tcfg.get("language_code"):
        base["language_code"] = tcfg["language_code"]

    total_cost, done = 0.0, 0
    for i in range(0, len(domains), 1000):
        chunk = domains[i:i + 1000]
        try:
            r = requests.post(DFS_URL, auth=(login, password), json=[{**base, "targets": chunk}], timeout=120)
            r.raise_for_status()
            data = r.json()
            task = (data.get("tasks") or [{}])[0]
            if task.get("status_code") != 20000:
                raise RuntimeError(f"{task.get('status_code')} {task.get('status_message')}")
            total_cost += float(task.get("cost") or 0)
            items = ((task.get("result") or [{}])[0] or {}).get("items") or []
        except Exception as e:  # noqa: BLE001
            log(f"Трафик: ошибка DataForSEO — {e}. Попробую в следующий запуск")
            break

        found = {}
        for it in items:
            org = (it.get("metrics") or {}).get("organic") or {}
            found[normalize_domain(it.get("target"))] = (round(org.get("etv") or 0), org.get("count") or 0)
        with conn:
            conn.executemany(
                "UPDATE added_sites SET traffic=?, keywords=?, traffic_date=? WHERE domain=? AND traffic_date IS NULL",
                [(*found.get(d, (0, 0)), today, d) for d in chunk])
        done += len(chunk)
    if done:
        log(f"Трафик: обновлено {done} сайтов, стоимость ${total_cost:.4f}")


# ---------- Дата регистрации домена (RDAP / WHOIS) ----------

RDAP_BOOTSTRAP = "https://data.iana.org/rdap/dns.json"


class RetryLater(Exception):
    """Временная ошибка — попробовать в следующий запуск."""


def load_rdap_servers():
    r = requests.get(RDAP_BOOTSTRAP, headers=HEADERS, timeout=30)
    r.raise_for_status()
    servers = {}
    for tlds, urls in r.json().get("services", []):
        for t in tlds:
            servers[t.lower()] = urls[0].rstrip("/") + "/"
    return servers


def rdap_created(domain, servers):
    """Дата регистрации через RDAP. None — если домен не найден / нет RDAP для зоны."""
    labels = domain.split(".")
    server = None
    for i in range(len(labels) - 1):  # самая длинная известная зона: co.uk раньше, чем uk
        server = servers.get(".".join(labels[i + 1:]))
        if server:
            break
    if not server:
        return None
    candidates = [".".join(labels[i:]) for i in range(len(labels) - 1)]  # blog.site.com → site.com
    for name in candidates:
        try:
            r = requests.get(server + "domain/" + name, headers={**HEADERS, "Accept": "application/rdap+json"},
                             timeout=20)
        except requests.RequestException as e:
            raise RetryLater(str(e))
        if r.status_code == 404:
            continue
        if r.status_code == 429 or r.status_code >= 500:
            raise RetryLater(f"HTTP {r.status_code}")
        if not r.ok:
            return None
        for ev in r.json().get("events", []):
            if ev.get("eventAction") == "registration" and ev.get("eventDate"):
                return ev["eventDate"][:10]
        return None
    return None


def whois_created(domain):
    """Запасной вариант через классический WHOIS (для зон без RDAP)."""
    try:
        import whois  # пакет python-whois
    except ImportError:
        return None
    import socket
    socket.setdefaulttimeout(15)  # чтобы зависший WHOIS-сервер не тормозил запуск
    try:
        w = whois.whois(domain)
        d = w.creation_date
        if isinstance(d, list):
            d = min(x for x in d if x)
        return d.strftime("%Y-%m-%d") if d else None
    except Exception:  # noqa: BLE001
        return None


def enrich_whois(conn, cfg, today):
    wcfg = cfg.get("whois", {})
    if not wcfg.get("enabled", True):
        return
    limit = int(wcfg.get("max_domains_per_run", 500))
    domains = [r[0] for r in conn.execute(
        "SELECT DISTINCT domain FROM added_sites WHERE whois_date IS NULL "
        "ORDER BY date_added DESC LIMIT ?", (limit,))]
    if not domains:
        return
    try:
        servers = load_rdap_servers()
    except Exception as e:  # noqa: BLE001
        log(f"WHOIS: не удалось загрузить список RDAP-серверов ({e}) — пропускаю")
        return

    found = retry = 0
    for d in domains:
        try:
            created = rdap_created(d, servers) or whois_created(d)
        except RetryLater:
            retry += 1
            continue
        with conn:
            conn.execute("UPDATE added_sites SET domain_created=?, whois_date=? WHERE domain=?",
                         (created, today, d))
        found += bool(created)
        time.sleep(float(wcfg.get("delay_sec", 0.5)))
    log(f"WHOIS: проверено {len(domains) - retry}, дата найдена у {found}"
        + (f", отложено до следующего запуска {retry}" if retry else ""))


# ---------- Экспорт в Excel ----------

HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", start_color="305496")
BODY_FONT = Font(name="Arial")


def write_sheet(ws, headers, rows, widths):
    ws.append(headers)
    for c in ws[1]:
        c.font, c.fill = HEADER_FONT, HEADER_FILL
    for row in rows:
        ws.append(list(row))
    for col_cells in ws.iter_cols(min_row=2):
        for c in col_cells:
            c.font = BODY_FONT
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions


def export_changes(conn):
    wb = Workbook()
    ws = wb.active
    ws.title = "Новые"
    write_sheet(ws, ["Дата добавления", "Сеть", "Сайт", "Владелец", "Трафик (мес.)", "Ключей",
                     "Домен зарегистрирован", "Seller ID", "Тип"],
                conn.execute("SELECT date_added, network, domain, name, CAST(traffic AS INTEGER), keywords, "
                             "domain_created, seller_id, seller_type "
                             "FROM added_sites ORDER BY date_added DESC, network, domain"),
                [16, 22, 34, 38, 14, 10, 20, 22, 12])
    for c in ws["E"][1:]:
        c.number_format = "#,##0"
    ws = wb.create_sheet("Удалённые")
    write_sheet(ws, ["Дата удаления", "Сеть", "Сайт", "Владелец", "Seller ID", "Тип"],
                conn.execute("SELECT date_removed, network, domain, name, seller_id, seller_type "
                             "FROM removed_sites ORDER BY date_removed DESC, network, domain"),
                [16, 22, 34, 38, 22, 12])
    ws = wb.create_sheet("Статус сетей")
    write_sheet(ws, ["Дата", "Сеть", "Статус", "Сайтов", "Новых", "Удалено", "URL", "Сообщение"],
                conn.execute("SELECT run_date, network, status, total, added, removed, url, message "
                             "FROM run_log WHERE rowid IN (SELECT MAX(rowid) FROM run_log GROUP BY network) "
                             "ORDER BY network"),
                [12, 22, 12, 10, 8, 9, 42, 50])
    wb.save(CHANGES_XLSX)
    log(f"Сохранено: {CHANGES_XLSX.name}")


def export_all(conn):
    wb = Workbook(write_only=False)
    ws = wb.active
    ws.title = "Все сайты"
    write_sheet(ws, ["Сеть", "Сайт", "Владелец", "Seller ID", "Тип", "Впервые замечен"],
                conn.execute("SELECT network, domain, name, seller_id, seller_type, first_seen "
                             "FROM all_sites ORDER BY network, domain"),
                [22, 34, 38, 22, 12, 16])
    wb.save(ALL_XLSX)
    log(f"Сохранено: {ALL_XLSX.name}")


def export_dashboard(conn):
    """Собирает docs/index.html — один файл со встроенными данными (GitHub Pages или открыть локально)."""
    if not TEMPLATE.exists():
        return
    def rows(sql):
        out = []
        for r in conn.execute(sql):
            d = {"date": r[0], "network": r[1], "domain": r[2], "name": r[3] or "", "seller_id": r[4] or ""}
            if len(r) > 5:
                d["traffic"], d["keywords"], d["created"] = r[5], r[6], r[7]
            out.append(d)
        return out
    last = conn.execute("SELECT MAX(run_date) FROM run_log").fetchone()[0]
    data = {
        "last_run": last,
        "total_sites": conn.execute("SELECT COUNT(*) FROM all_sites").fetchone()[0],
        "added": rows("SELECT date_added, network, domain, name, seller_id, "
                      "CAST(traffic AS INTEGER), keywords, domain_created FROM added_sites"),
        "removed": rows("SELECT date_removed, network, domain, name, seller_id FROM removed_sites"),
        "status": [dict(zip(["network", "status", "total", "added", "removed", "url", "message"], r))
                   for r in conn.execute("SELECT network, status, total, added, removed, url, message "
                                         "FROM run_log WHERE rowid IN (SELECT MAX(rowid) FROM run_log GROUP BY network) "
                                         "ORDER BY network")],
    }
    payload = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    DASHBOARD.parent.mkdir(exist_ok=True)
    DASHBOARD.write_text(TEMPLATE.read_text(encoding="utf-8").replace("/*__DATA__*/", payload), encoding="utf-8")
    log(f"Сохранено: docs/index.html")


def main():
    ap = argparse.ArgumentParser(description="Мониторинг sellers.json")
    ap.add_argument("--export-all", action="store_true", help="выгрузить полный список в all_sites.xlsx")
    ap.add_argument("--check", action="store_true", help="только проверить доступность URL")
    ap.add_argument("--date", help="дата запуска YYYY-MM-DD (по умолчанию сегодня, UTC)")
    args = ap.parse_args()

    cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    networks = [n for n in cfg["networks"] if n.get("enabled", True)]

    if args.check:
        for n in networks:
            try:
                url, raw = fetch_sellers(n["urls"], retries=1)
                pubs = parse_sellers(raw, cfg.get("include_seller_types"))
                log(f"✓ {n['name']}: {url} — записей {len(raw)}, сайтов после фильтра {len(pubs)}")
            except Exception as e:  # noqa: BLE001
                log(f"✗ {n['name']}: {e}")
        return

    today = args.date or dt.datetime.utcnow().strftime("%Y-%m-%d")
    conn = sqlite3.connect(DB_PATH)
    conn.executescript(SCHEMA)
    migrate(conn)
    log(f"Запуск за {today}")
    for n in networks:
        with conn:  # транзакция на каждую сеть
            process_network(conn, n, today, cfg)
    enrich_traffic(conn, cfg, today)
    enrich_whois(conn, cfg, today)
    export_changes(conn)
    export_dashboard(conn)
    if args.export_all:
        export_all(conn)
    conn.close()


if __name__ == "__main__":
    sys.exit(main())
