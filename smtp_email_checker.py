#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""smtp_email_checker.py — многопоточная проверка email.

Проверка по этапам:
1) Синтаксис (email-validator)
2) MX (DNS)
3) SMTP диалог: HELO -> MAIL FROM -> RCPT TO

Crash-safe сохранение во время работы:
- results/<ts>/checkpoint.jsonl (append, 1 строка = 1 результат)
- results/<ts>/*.csv по категориям (append)

Проблема "долгой паузы" после 100% прогресса обычно связана с финализацией:
- пересборка Excel (pandas/openpyxl) может быть очень долгой на больших файлах,
- закрытие файлов и финальное flush/fsync,
- финальное создание summary.

В этой версии добавлен дружелюбный вывод с индикатором активности на этапе финализации,
чтобы было понятно, что программа НЕ зависла.
"""

from __future__ import annotations

import csv
import json
import logging
import os
import random
import socket
import smtplib
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, asdict
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import dns.exception
import dns.resolver
import pandas as pd
from colorama import Fore, Style, init
from dotenv import load_dotenv
from email_validator import EmailNotValidError, validate_email

from filters import EmailFilter


# ======================
# НАСТРОЙКИ
# ======================
INPUT_FILE = "emails+1.xlsx"

OUTPUT_DIR = "results"  # results/<timestamp>/
NUM_THREADS = 15

DNS_TIMEOUT_SEC = 8
SMTP_TIMEOUT_SEC = 20
MAX_MX_SERVERS = 3

# Сохранять прогресс сразу (чтобы падение не теряло результат)
FLUSH_EVERY_RESULT = True
FSYNC_EVERY_N = 50            # fsync раз в N результатов (0 = отключить)

# Excel тяжёлый для постоянной дозаписи: делаем периодическую пересборку из CSV/JSONL
REBUILD_XLSX_EVERY_N = 200    # 0 = не пересобирать в процессе, только в конце
FINAL_BUILD_XLSX = True       # пересобрать xlsx в конце

# Небольшой jitter между SMTP-подключениями (уменьшает риск блокировок)
JITTER_SEC_RANGE = (0.0, 0.25)

# Повторять ли при временных кодах 4xx
RETRY_ON_TEMPORARY = False
RETRY_COUNT = 1
RETRY_DELAY_SEC = 2

# Консоль
USE_COLORS = True
PROGRESS_EVERY = 20
VERBOSE_PER_EMAIL = False

# SMTP идентификация
HELO_HOSTNAME = "mailchecker.local"
MAIL_FROM = "checker@mailchecker.local"

# Фильтры
FILTER_CONFIG_PATH = "filters_config.json"

# Дружелюбный индикатор финализации
FINALIZATION_HEARTBEAT_SEC = 1.0


# ======================
# Логирование
# ======================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)
logger = logging.getLogger("smtp_email_checker")


# ======================
# Модели результата
# ======================
@dataclass
class MXRecord:
    priority: int
    host: str


@dataclass
class SMTPCheck:
    mx_host: str
    priority: int
    status: str  # valid/invalid/temporary/unknown
    code: Optional[int] = None
    message: str = ""
    reason: str = ""
    delivery_possible: bool = False
    smtp_dialog: List[str] = None
    time_taken_sec: float = 0.0


@dataclass
class EmailResult:
    email: str
    ts: str

    user: str = ""
    domain: str = ""

    format_ok: bool = False
    mx_ok: bool = False
    mx_records: List[MXRecord] = None

    filtered: bool = False
    filter_reason: str = ""

    final_status: str = "unknown"  # valid/invalid/temporary/unknown/error
    confidence: int = 0

    smtp_checks: List[SMTPCheck] = None
    errors: List[str] = None
    total_time_sec: float = 0.0


# ======================
# UI helpers
# ======================

def c(text: str, color: str) -> str:
    if not USE_COLORS:
        return text
    return color + text + Style.RESET_ALL


class Heartbeat:
    """Периодически печатает 'живой' индикатор, пока идёт длинная операция."""

    def __init__(self, title: str, interval_sec: float = 1.0):
        self.title = title
        self.interval_sec = interval_sec
        self._stop = threading.Event()
        self._t = threading.Thread(target=self._run, daemon=True)
        self._frames = ["|", "/", "-", "\\"]
        self._idx = 0

    def start(self):
        self._t.start()

    def stop(self):
        self._stop.set()
        self._t.join(timeout=2)
        # перенос строки после \r
        print()

    def _run(self):
        while not self._stop.is_set():
            frame = self._frames[self._idx % len(self._frames)]
            self._idx += 1
            # \r чтобы обновлять строку
            print(c(f"{self.title} {frame}", Fore.CYAN), end="\r", flush=True)
            time.sleep(self.interval_sec)


# ======================
# Утилиты
# ======================

def validate_format(email: str) -> Tuple[bool, str]:
    try:
        validate_email(email)
        return True, "OK"
    except EmailNotValidError as e:
        return False, str(e)


def parse_email(email: str) -> Tuple[str, str]:
    user, domain = email.split("@", 1)
    return user, domain.lower().strip().rstrip(".")


def get_mx(domain: str) -> Optional[List[MXRecord]]:
    try:
        resolver = dns.resolver.Resolver()
        resolver.lifetime = DNS_TIMEOUT_SEC
        resolver.timeout = DNS_TIMEOUT_SEC

        answers = resolver.resolve(domain, "MX")
        mxs: List[MXRecord] = []
        for r in answers:
            mxs.append(MXRecord(priority=int(r.preference), host=str(r.exchange).rstrip(".")))
        mxs.sort(key=lambda x: x.priority)
        return mxs
    except (dns.resolver.NXDOMAIN, dns.resolver.NoAnswer):
        return None
    except dns.exception.Timeout:
        return None
    except Exception as e:
        logger.warning("MX lookup error for %s: %s", domain, e)
        return None


def smtp_probe(mx_host: str, priority: int, rcpt_email: str) -> SMTPCheck:
    started = time.time()
    dialog: List[str] = []
    res = SMTPCheck(mx_host=mx_host, priority=priority, status="unknown", smtp_dialog=dialog)

    try:
        with smtplib.SMTP(mx_host, 25, timeout=SMTP_TIMEOUT_SEC) as server:
            code, msg = server.helo(HELO_HOSTNAME)
            msg_s = msg.decode(errors="ignore") if isinstance(msg, (bytes, bytearray)) else str(msg)
            dialog.append(f"HELO {HELO_HOSTNAME} -> {code} {msg_s}")

            code, msg = server.mail(MAIL_FROM)
            msg_s = msg.decode(errors="ignore") if isinstance(msg, (bytes, bytearray)) else str(msg)
            dialog.append(f"MAIL FROM:<{MAIL_FROM}> -> {code} {msg_s}")
            if code != 250:
                res.code = int(code)
                res.message = msg_s
                res.reason = "MAIL FROM rejected"
                res.status = "unknown"
                return res

            code, msg = server.rcpt(rcpt_email)
            msg_s = msg.decode(errors="ignore") if isinstance(msg, (bytes, bytearray)) else str(msg)
            dialog.append(f"RCPT TO:<{rcpt_email}> -> {code} {msg_s}")

            res.code = int(code)
            res.message = msg_s

            if code == 250:
                res.status = "valid"
                res.delivery_possible = True
            elif code in (421, 450, 451, 452):
                res.status = "temporary"
                res.reason = f"Temporary SMTP error ({code})"
            elif 500 <= code <= 599:
                res.status = "invalid"
                res.reason = f"Permanent SMTP error ({code})"
            else:
                res.status = "unknown"
                res.reason = f"Unhandled SMTP code ({code})"

    except socket.timeout:
        res.status = "temporary"
        res.reason = "SMTP timeout"
    except (smtplib.SMTPConnectError, ConnectionRefusedError, OSError) as e:
        res.status = "temporary"
        res.reason = f"SMTP connection error: {e}"
    except smtplib.SMTPServerDisconnected:
        res.status = "temporary"
        res.reason = "SMTP server disconnected"
    except Exception as e:
        res.status = "unknown"
        res.reason = f"SMTP error: {e}"
    finally:
        res.time_taken_sec = round(time.time() - started, 3)

    return res


def read_emails_from_excel(path: str) -> List[str]:
    df = pd.read_excel(path)
    cols = [c for c in df.columns if "email" in str(c).lower()]
    if not cols:
        raise ValueError("Не найден столбец, содержащий 'email' в названии")

    emails = df[cols[0]].dropna().astype(str).map(lambda x: x.strip()).tolist()
    emails = [e for e in emails if e]

    seen = set()
    out = []
    for e in emails:
        if e not in seen:
            out.append(e)
            seen.add(e)
    return out


# ======================
# Crash-safe writer
# ======================
class ResultWriter:
    """Пишет результаты ПО МЕРЕ ГОТОВНОСТИ (append)."""

    def __init__(self, run_dir: Path):
        self.run_dir = run_dir
        self.lock = threading.Lock()
        self.count = 0

        self.jsonl_path = run_dir / "checkpoint.jsonl"
        self.checked_path = run_dir / "checked.txt"

        self.csv_paths = {
            "valid": run_dir / "valid.csv",
            "invalid": run_dir / "invalid.csv",
            "temporary": run_dir / "temporary.csv",
            "unknown": run_dir / "unknown.csv",
            "error": run_dir / "error.csv",
        }

        if not self.jsonl_path.exists():
            self.jsonl_path.write_text("", encoding="utf-8")

        for k, p in self.csv_paths.items():
            if not p.exists():
                with p.open("w", newline="", encoding="utf-8") as f:
                    csv.writer(f).writerow(["Email"])

        if not self.checked_path.exists():
            self.checked_path.write_text("", encoding="utf-8")

        self._jsonl_f = self.jsonl_path.open("a", encoding="utf-8")
        self._checked_f = self.checked_path.open("a", encoding="utf-8")

        self._csv_f = {k: self.csv_paths[k].open("a", newline="", encoding="utf-8") for k in self.csv_paths}
        self._csv_w = {k: csv.writer(self._csv_f[k]) for k in self.csv_paths}

    def close(self) -> None:
        with self.lock:
            for f in [self._jsonl_f, self._checked_f, *self._csv_f.values()]:
                try:
                    f.close()
                except Exception:
                    pass

    def write(self, res: EmailResult) -> None:
        with self.lock:
            self.count += 1

            self._jsonl_f.write(json.dumps(_result_to_dict(res), ensure_ascii=False) + "\n")
            self._checked_f.write(res.email + "\n")

            st = res.final_status if res.final_status in self.csv_paths else "unknown"
            self._csv_w[st].writerow([res.email])

            if FLUSH_EVERY_RESULT:
                self._jsonl_f.flush()
                self._checked_f.flush()
                for f in self._csv_f.values():
                    f.flush()

            if FSYNC_EVERY_N and (self.count % FSYNC_EVERY_N == 0):
                try:
                    os.fsync(self._jsonl_f.fileno())
                except Exception:
                    pass

    def rebuild_xlsx_from_csv(self, ts: str, verbose: bool = True) -> None:
        """Пересобирает xlsx из CSV."""

        def csv_to_xlsx(csv_path: Path, out_path: Path):
            if not csv_path.exists():
                return
            if verbose:
                print(c(f"  -> Создаю {out_path.name} ...", Fore.CYAN))
            df = pd.read_csv(csv_path)
            df.to_excel(out_path, index=False)
            if verbose:
                print(c(f"  -> Готово: {out_path.name} (строк: {len(df)})", Fore.GREEN))

        csv_to_xlsx(self.csv_paths["valid"], self.run_dir / f"01_valid_emails_{ts}.xlsx")
        csv_to_xlsx(self.csv_paths["invalid"], self.run_dir / f"02_invalid_emails_{ts}.xlsx")
        csv_to_xlsx(self.csv_paths["temporary"], self.run_dir / f"03_temporary_emails_{ts}.xlsx")
        csv_to_xlsx(self.csv_paths["unknown"], self.run_dir / f"04_unknown_emails_{ts}.xlsx")
        csv_to_xlsx(self.csv_paths["error"], self.run_dir / f"05_error_emails_{ts}.xlsx")


def _result_to_dict(r: EmailResult) -> Dict[str, Any]:
    d = asdict(r)
    d["mx_records"] = [asdict(x) for x in (r.mx_records or [])]
    d["smtp_checks"] = [asdict(x) for x in (r.smtp_checks or [])]
    return d


# ======================
# Основная валидация одного email
# ======================

def validate_one(email: str, email_filter: EmailFilter) -> EmailResult:
    t0 = time.time()
    r = EmailResult(
        email=email,
        ts=datetime.now().isoformat(timespec="seconds"),
        mx_records=[],
        smtp_checks=[],
        errors=[],
    )

    # jitter
    jmin, jmax = JITTER_SEC_RANGE
    if jmax > 0:
        time.sleep(random.uniform(jmin, jmax))

    # 0) фильтр
    decision = email_filter.check(email)
    if decision.blocked:
        r.filtered = True
        r.filter_reason = decision.reason
        r.final_status = "invalid"
        r.confidence = 0
        r.errors.append(decision.reason)
        r.total_time_sec = round(time.time() - t0, 3)
        return r

    # 1) формат
    ok, msg = validate_format(email)
    r.format_ok = ok
    if not ok:
        r.final_status = "invalid"
        r.confidence = 0
        r.errors.append(f"Syntax error: {msg}")
        r.total_time_sec = round(time.time() - t0, 3)
        return r

    # 2) parse
    try:
        r.user, r.domain = parse_email(email)
    except Exception:
        r.final_status = "invalid"
        r.confidence = 0
        r.errors.append("Cannot parse email")
        r.total_time_sec = round(time.time() - t0, 3)
        return r

    # 3) MX
    mx = get_mx(r.domain)
    if not mx:
        r.mx_ok = False
        r.final_status = "invalid"
        r.confidence = 0
        r.errors.append("No MX records or domain does not exist")
        r.total_time_sec = round(time.time() - t0, 3)
        return r

    r.mx_ok = True
    r.mx_records = mx

    # 4) SMTP
    mx_to_check = mx[: max(1, min(MAX_MX_SERVERS, len(mx)))]

    def run_checks_once() -> List[SMTPCheck]:
        out: List[SMTPCheck] = []
        for rec in mx_to_check:
            out.append(smtp_probe(rec.host, rec.priority, email))
        return out

    checks = run_checks_once()

    if RETRY_ON_TEMPORARY:
        need_retry = any(c.status == "temporary" for c in checks)
        tries = 0
        while need_retry and tries < RETRY_COUNT:
            tries += 1
            time.sleep(RETRY_DELAY_SEC)
            checks = run_checks_once()
            need_retry = any(c.status == "temporary" for c in checks)

    r.smtp_checks = checks

    valid_cnt = sum(1 for c in checks if c.status == "valid")
    invalid_cnt = sum(1 for c in checks if c.status == "invalid")
    temp_cnt = sum(1 for c in checks if c.status == "temporary")

    if valid_cnt > 0:
        r.final_status = "valid"
        r.confidence = 100
    elif invalid_cnt > 0 and temp_cnt == 0:
        r.final_status = "invalid"
        r.confidence = 95
    elif temp_cnt > 0 and invalid_cnt == 0:
        r.final_status = "temporary"
        r.confidence = 50
    else:
        r.final_status = "unknown"
        r.confidence = 30

    r.total_time_sec = round(time.time() - t0, 3)
    return r


# ======================
# Summary
# ======================

def write_summary(run_dir: Path, ts: str, total: int) -> None:
    """Сводка на базе CSV, чтобы не держать всё в памяти."""
    def count_rows(csv_path: Path) -> int:
        if not csv_path.exists():
            return 0
        try:
            df = pd.read_csv(csv_path)
            return int(len(df))
        except Exception:
            n = 0
            with csv_path.open("r", encoding="utf-8", errors="ignore") as f:
                for i, _ in enumerate(f):
                    n = i
            return max(0, n)

    valid_n = count_rows(run_dir / "valid.csv")
    invalid_n = count_rows(run_dir / "invalid.csv")
    temp_n = count_rows(run_dir / "temporary.csv")
    unknown_n = count_rows(run_dir / "unknown.csv")
    error_n = count_rows(run_dir / "error.csv")

    summary_path = run_dir / f"00_summary_{ts}.txt"
    with summary_path.open("w", encoding="utf-8") as f:
        f.write("Email validation summary\n")
        f.write(f"Timestamp: {ts}\n")
        f.write(f"Total in input: {total}\n\n")
        f.write(f"Valid: {valid_n}\n")
        f.write(f"Invalid: {invalid_n}\n")
        f.write(f"Temporary: {temp_n}\n")
        f.write(f"Unknown: {unknown_n}\n")
        f.write(f"Error: {error_n}\n")


# ======================
# main
# ======================

def main() -> None:
    if USE_COLORS:
        init(autoreset=True)
    load_dotenv()

    if not os.path.exists(INPUT_FILE):
        print(c(f"Файл не найден: {INPUT_FILE}", Fore.RED))
        return

    emails = read_emails_from_excel(INPUT_FILE)
    total = len(emails)
    if total == 0:
        print(c("В файле нет email адресов", Fore.RED))
        return

    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    run_dir = Path(OUTPUT_DIR) / ts
    run_dir.mkdir(parents=True, exist_ok=True)

    email_filter = EmailFilter(FILTER_CONFIG_PATH)
    writer = ResultWriter(run_dir)

    print(c(f"Проверка: {total} email | потоков={NUM_THREADS} | SMTP timeout={SMTP_TIMEOUT_SEC}s", Fore.MAGENTA))
    print(c(f"Результаты пишутся сразу (checkpoint.jsonl + CSV) в: {run_dir}", Fore.GREEN))

    start = time.time()
    done = 0

    try:
        with ThreadPoolExecutor(max_workers=NUM_THREADS) as ex:
            futures = {ex.submit(validate_one, e, email_filter): e for e in emails}

            for f in as_completed(futures):
                done += 1
                e = futures[f]

                try:
                    res = f.result()
                except Exception as exc:
                    res = EmailResult(
                        email=e,
                        ts=datetime.now().isoformat(timespec="seconds"),
                        final_status="error",
                        confidence=0,
                        format_ok=False,
                        mx_ok=False,
                        mx_records=[],
                        smtp_checks=[],
                        errors=[f"Unhandled exception: {exc}"],
                    )

                writer.write(res)

                if VERBOSE_PER_EMAIL:
                    color = {
                        "valid": Fore.GREEN,
                        "invalid": Fore.RED,
                        "temporary": Fore.YELLOW,
                        "unknown": Fore.CYAN,
                        "error": Fore.RED,
                    }.get(res.final_status, Fore.WHITE)
                    print(c(f"{res.final_status.upper():9} {res.email}", color))

                if done % PROGRESS_EVERY == 0 or done == total:
                    elapsed = time.time() - start
                    avg = elapsed / done
                    remaining = avg * (total - done)
                    print(c(f"{done}/{total} ({done*100/total:.1f}%), осталось ~{timedelta(seconds=int(remaining))}", Fore.BLUE))

                if REBUILD_XLSX_EVERY_N and done % REBUILD_XLSX_EVERY_N == 0:
                    print(c("Финализация (промежуточно): пересборка Excel из CSV...", Fore.CYAN))
                    hb = Heartbeat("Идёт пересборка Excel", FINALIZATION_HEARTBEAT_SEC)
                    hb.start()
                    try:
                        writer.rebuild_xlsx_from_csv(ts, verbose=False)
                    finally:
                        hb.stop()
                    print(c("Промежуточная пересборка Excel завершена.", Fore.GREEN))

    finally:
        print(c("\n100% проверок завершено. Начинаю финализацию результатов...", Fore.MAGENTA))
        print(c("Это может занять время на больших базах (особенно создание Excel).", Fore.MAGENTA))

        if FINAL_BUILD_XLSX:
            print(c("Шаг 1/3: Финальная пересборка Excel файлов из CSV.", Fore.CYAN))
            hb = Heartbeat("Финальная пересборка Excel", FINALIZATION_HEARTBEAT_SEC)
            hb.start()
            try:
                writer.rebuild_xlsx_from_csv(ts, verbose=True)
            except Exception as e_reb:
                logger.warning("Final XLSX rebuild failed: %s", e_reb)
            finally:
                hb.stop()
            print(c("Шаг 1/3 завершён.", Fore.GREEN))
        else:
            print(c("Шаг 1/3: Пропущено (FINAL_BUILD_XLSX=False).", Fore.YELLOW))

        print(c("Шаг 2/3: Создаю summary файл...", Fore.CYAN))
        hb = Heartbeat("Запись summary", FINALIZATION_HEARTBEAT_SEC)
        hb.start()
        try:
            write_summary(run_dir, ts, total)
        except Exception as e_sum:
            logger.warning("Summary write failed: %s", e_sum)
        finally:
            hb.stop()
        print(c("Шаг 2/3 завершён.", Fore.GREEN))

        print(c("Шаг 3/3: Закрываю файлы и освобождаю ресурсы...", Fore.CYAN))
        hb = Heartbeat("Закрытие файлов", FINALIZATION_HEARTBEAT_SEC)
        hb.start()
        try:
            writer.close()
        finally:
            hb.stop()
        print(c("Шаг 3/3 завершён.", Fore.GREEN))

    elapsed_total = time.time() - start
    print(c(f"\nГотово за {timedelta(seconds=int(elapsed_total))}", Fore.MAGENTA))
    print(c(f"Папка результатов: {run_dir}", Fore.GREEN))
    print(c("Если программа упадёт — частичные результаты всё равно останутся в checkpoint.jsonl и CSV.", Fore.GREEN))
    print(c("Подсказка: если Excel делается слишком долго — поставьте FINAL_BUILD_XLSX=False.", Fore.BLUE))


if __name__ == "__main__":
    main()
