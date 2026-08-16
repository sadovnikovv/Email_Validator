#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""smtp_email_checker.py — многопоточная проверка email.

Проверка по этапам:
1) Синтаксис (email-validator)
2) MX (DNS)
3) SMTP-диалог в одном из режимов:

   SMTP_CHECK_MODE = "mx_probe"
       Исходная логика:
       MX получателя:25 -> HELO -> MAIL FROM -> RCPT TO

   SMTP_CHECK_MODE = "submission_envelope"
       Проверка через SMTP REG.RU:
       mail.hosting.reg.ru:465 -> SSL -> AUTH -> MAIL FROM -> RCPT TO -> RSET
       Команда DATA не вызывается, письмо фактически не отправляется.

   SMTP_CHECK_MODE = "submission_send"
       То же подключение через SMTP REG.RU, но выполняется DATA и отправляется
       реальное техническое письмо. Использовать только для базы контактов,
       давших согласие на получение рассылки.

Crash-safe сохранение во время работы:
- results/<ts>/checkpoint.jsonl
- results/<ts>/*.csv по категориям

Про Excel:
- BUILD_EXCEL_FILES=False: быстрый режим, остаются CSV/JSONL.
- BUILD_EXCEL_FILES=True: формируются XLSX-отчёты.

Про входной файл:
- INPUT_FILE может быть .xlsx/.xls или .csv.
- Заголовок может быть любым либо отсутствовать.
- Столбец с email определяется по содержимому.
"""

from __future__ import annotations

import csv
import json
import logging
import os
import random
import re
import socket
import smtplib
import ssl
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict, dataclass
from datetime import datetime, timedelta
from email.message import EmailMessage
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import functools
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

# Вход: .xlsx/.xls или .csv
# INPUT_FILE = "техно.csv"
INPUT_FILE = str(Path(r"C:\Users\user\Downloads\ххх.csv"))

OUTPUT_DIR = "results"
NUM_THREADS = 15

DNS_TIMEOUT_SEC = 8
SMTP_TIMEOUT_SEC = 5
MAX_MX_SERVERS = 2

# Сохранять прогресс сразу
FLUSH_EVERY_RESULT = True
FSYNC_EVERY_N = 50

# Excel
BUILD_EXCEL_FILES = False
REBUILD_XLSX_EVERY_N = 200
FINAL_BUILD_XLSX = True

# Пауза между SMTP-подключениями
JITTER_SEC_RANGE = (0.0, 0.05)

# Повтор временных ошибок
RETRY_ON_TEMPORARY = False
RETRY_COUNT = 1
RETRY_DELAY_SEC = 2

# Консоль
USE_COLORS = True
PROGRESS_EVERY = 20
VERBOSE_PER_EMAIL = False

# Исходная SMTP-идентификация
HELO_HOSTNAME = "mailchecker.zehvk.ru"
MAIL_FROM = "info@zehvk.ru"


# ======================
# РАЗЛИЧИЯ МЕЖДУ РЕЖИМАМИ
# ======================
#
# SMTP_CHECK_MODE = "submission_envelope"
# ─────────────────────────────────────────
# Назначение: Проверка валидности email БЕЗ реальной отправки письма
#
# SMTP-диалог:
#   1. EHLO                           (приветствие сервера)
#   2. AUTH LOGIN                     (авторизация на SMTP REG.RU)
#   3. MAIL FROM:<info@zehvk.ru>      (отправитель)
#   4. RCPT TO:<проверяемый@email>    (получатель)
#   5. RSET                           (сброс сессии, ОТМЕНА отправки)
#   ❌ DATA НЕ вызывается
#   ❌ Письмо НЕ формируется
#   ❌ Письмо НЕ отправляется
#
# Результат:
#   - Если RCPT TO принят (код 2xx) → статус "submission_accepted"
#   - Сервер REG.RU подтвердил, что адрес существует и готов принять почту
#   - Но фактически получателю ничего не приходит
#
# Когда использовать:
#   ✅ Для проверки больших баз контактов
#   ✅ Когда нужно только валидировать адреса
#   ✅ Когда не хотите беспокоить получателей
#   ✅ Безопасный режим для любой базы
#
#
# SMTP_CHECK_MODE = "submission_send"
# ─────────────────────────────────────────
# Назначение: РЕАЛЬНАЯ отправка технического письма через SMTP REG.RU
#
# SMTP-диалог:
#   1. EHLO                           (приветствие сервера)
#   2. AUTH LOGIN                     (авторизация на SMTP REG.RU)
#   3. MAIL FROM:<info@zehvk.ru>      (отправитель)
#   4. RCPT TO:<проверяемый@email>    (получатель)
#   5. DATA                           (начало передачи письма)
#   6. Отправка EmailMessage          (реальное письмо с Subject и Body)
#   ✅ Письмо ФОРМИРУЕТСЯ
#   ✅ Письмо ОТПРАВЛЯЕТСЯ получателю
#
# Результат:
#   - Если сервер принял письмо → статус "submitted"
#   - Получатель ПОЛУЧИТ письмо с темой "Проверка доставки"
#   - Письмо появится в его почтовом ящике (возможно в спаме)
#
# Когда использовать:
#   ⚠️  ТОЛЬКО для базы контактов, давших явное согласие на рассылку
#   ⚠️  Когда нужно проверить не только существование адреса, но и доставляемость
#   ⚠️  Требует осторожности — каждая проверка = реальное письмо
#   ❌ НЕ использовать для проверки произвольных баз без согласия
#
#
# КЛЮЧЕВОЕ ОТЛИЧИЕ:
# ─────────────────────────────────────────
# submission_envelope → Проверяет, готов ли сервер ПРИНЯТЬ письмо (безопасно)
# submission_send     → Фактически ОТПРАВЛЯЕТ письмо получателю (требует согласия)

# ======================
# ДОБАВЛЕНО: ВЫБОР SMTP-РЕЖИМА
# ======================

# Допустимые значения:
# "mx_probe"
# "submission_envelope"
# "submission_send"
SMTP_CHECK_MODE = "mx_probe"

# Настройки REG.RU используются только в submission_envelope/submission_send
SMTP_HOST = "mail.hosting.reg.ru"
SMTP_PORT = 465
SMTP_USE_SSL = True
SMTP_USERNAME = "info@zehvk.ru"

# Пароль не хранится в этом файле.
# Создайте .env рядом со скриптом:
# SMTP_PASSWORD=ваш_реальный_пароль
SMTP_PASSWORD_ENV = "SMTP_PASSWORD"

# Одновременно открывать не больше этого числа соединений к REG.RU.
# NUM_THREADS при этом может оставаться 15.
SMTP_SUBMISSION_MAX_CONNECTIONS = 10
SMTP_SUBMISSION_SEMAPHORE = threading.BoundedSemaphore(
    SMTP_SUBMISSION_MAX_CONNECTIONS
)

# Используется только в режиме submission_send
SEND_SUBJECT = "Проверка доставки"
SEND_BODY = (
    "Здравствуйте!\n\n"
    "Это техническое сообщение для проверки доставки на адрес, "
    "указанный в подписке.\n"
)

# Фильтры
FILTER_CONFIG_PATH = "filters_config.json"

# Индикатор финализации
FINALIZATION_HEARTBEAT_SEC = 1.0

# Поиск email в файле
EMAIL_LIKE_REGEX = re.compile(
    r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}",
    re.IGNORECASE,
)
DETECT_ROWS = 200


# ======================
# ЛОГИРОВАНИЕ
# ======================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)
logger = logging.getLogger("smtp_email_checker")


# ======================
# МОДЕЛИ РЕЗУЛЬТАТА
# ======================

@dataclass
class MXRecord:
    priority: int
    host: str


@dataclass
class SMTPCheck:
    mx_host: str
    priority: int

    # valid/invalid/temporary/unknown/submission_accepted/submitted
    status: str

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

    # valid/invalid/temporary/unknown/error/submission_accepted/submitted
    final_status: str = "unknown"
    confidence: int = 0

    smtp_checks: List[SMTPCheck] = None
    errors: List[str] = None
    total_time_sec: float = 0.0


# ======================
# UI HELPERS
# ======================

def c(text: str, color: str) -> str:
    if not USE_COLORS:
        return text
    return color + text + Style.RESET_ALL


class Heartbeat:
    """Периодически печатает индикатор, пока идёт длинная операция."""

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
        print()

    def _run(self):
        while not self._stop.is_set():
            frame = self._frames[self._idx % len(self._frames)]
            self._idx += 1
            print(
                c(f"{self.title} {frame}", Fore.CYAN),
                end="\r",
                flush=True,
            )
            time.sleep(self.interval_sec)


# ======================
# УТИЛИТЫ
# ======================

def _decode_smtp_message(value: Any) -> str:
    if isinstance(value, (bytes, bytearray)):
        return value.decode("utf-8", errors="replace")
    return str(value or "")


def _smtp_success(code: int) -> bool:
    return 200 <= int(code) < 300


def validate_format(email: str) -> Tuple[bool, str]:
    try:
        validate_email(email)
        return True, "OK"
    except EmailNotValidError as e:
        return False, str(e)


def parse_email(email: str) -> Tuple[str, str]:
    user, domain = email.split("@", 1)
    return user, domain.lower().strip().rstrip(".")


@functools.lru_cache(maxsize=4096)
def get_mx(domain: str) -> Optional[List[MXRecord]]:
    try:
        resolver = dns.resolver.Resolver()

        # Явно задаём публичные DNS
        resolver.nameservers = ["8.8.8.8", "1.1.1.1"]
        resolver.lifetime = DNS_TIMEOUT_SEC
        resolver.timeout = DNS_TIMEOUT_SEC

        # Ответ без DNSSEC-проверки подписи
        resolver.use_edns(edns=0, ednsflags=0, payload=512)

        answers = resolver.resolve(domain, "MX")

        mxs: List[MXRecord] = []
        for record in answers:
            mxs.append(
                MXRecord(
                    priority=int(record.preference),
                    host=str(record.exchange).rstrip("."),
                )
            )

        mxs.sort(key=lambda x: x.priority)
        return mxs

    except dns.resolver.NXDOMAIN:
        logger.warning("Domain %s does not exist (NXDOMAIN)", domain)
        return None

    except dns.resolver.NoAnswer:
        logger.warning("No MX records found for %s", domain)
        return None

    except dns.exception.Timeout as e:
        logger.warning("DNS timeout for %s: %s", domain, e)
        return None

    except Exception as e:
        logger.warning(
            "MX lookup error for %s: %s - %s",
            domain,
            type(e).__name__,
            e,
        )
        return None


# ======================
# SMTP: ИСХОДНЫЙ MX-РЕЖИМ
# ======================

def smtp_probe(
    mx_host: str,
    priority: int,
    rcpt_email: str,
) -> SMTPCheck:
    """Исходная проверка: прямое подключение к MX-серверу на порт 25."""

    started = time.time()
    dialog: List[str] = []

    res = SMTPCheck(
        mx_host=mx_host,
        priority=priority,
        status="unknown",
        smtp_dialog=dialog,
    )

    try:
        with smtplib.SMTP(
            mx_host,
            25,
            timeout=SMTP_TIMEOUT_SEC,
        ) as server:
            code, msg = server.helo(HELO_HOSTNAME)
            msg_s = _decode_smtp_message(msg)

            dialog.append(
                f"HELO {HELO_HOSTNAME} -> {code} {msg_s}"
            )

            code, msg = server.mail(MAIL_FROM)
            msg_s = _decode_smtp_message(msg)

            dialog.append(
                f"MAIL FROM:<{MAIL_FROM}> -> {code} {msg_s}"
            )

            if not _smtp_success(code):
                res.code = int(code)
                res.message = msg_s
                res.reason = "MAIL FROM rejected"
                res.status = "unknown"
                return res

            code, msg = server.rcpt(rcpt_email)
            msg_s = _decode_smtp_message(msg)

            dialog.append(
                f"RCPT TO:<{rcpt_email}> -> {code} {msg_s}"
            )

            res.code = int(code)
            res.message = msg_s

            if _smtp_success(code):
                res.status = "valid"
                res.delivery_possible = True

            elif 400 <= code < 500:
                res.status = "temporary"
                res.reason = f"Temporary SMTP error ({code})"

            elif 500 <= code < 600:
                res.status = "invalid"
                res.reason = f"Permanent SMTP error ({code})"

            else:
                res.status = "unknown"
                res.reason = f"Unhandled SMTP code ({code})"

    except socket.timeout:
        res.status = "temporary"
        res.reason = "SMTP timeout"

    except (
        smtplib.SMTPConnectError,
        ConnectionRefusedError,
        OSError,
    ) as e:
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
# ======================
# SMTP: ДОБАВЛЕННЫЙ РЕЖИМ REG.RU
# ======================

def smtp_submission_probe(rcpt_email: str) -> SMTPCheck:
    """Проверяет адрес через авторизованный SMTP REG.RU.

    submission_envelope:
        AUTH -> MAIL FROM -> RCPT TO -> RSET.
        DATA не вызывается, письмо не создаётся.

    submission_send:
        AUTH -> MAIL FROM -> RCPT TO -> DATA.
        Отправляется реальное техническое письмо.

    Важно: принятие SMTP-конверта сервером REG.RU не гарантирует
    окончательную доставку в папку «Входящие».
    """

    started = time.time()
    dialog: List[str] = []

    res = SMTPCheck(
        mx_host=SMTP_HOST,
        priority=0,
        status="unknown",
        smtp_dialog=dialog,
    )

    password = os.getenv(SMTP_PASSWORD_ENV, "").strip()

    if not password:
        res.status = "error"
        res.reason = (
            f"Не задана переменная окружения {SMTP_PASSWORD_ENV}. "
            f"Создайте .env рядом со скриптом: "
            f"{SMTP_PASSWORD_ENV}=пароль_ящика"
        )
        res.time_taken_sec = round(time.time() - started, 3)
        return res

    try:
        # Ограничиваем число одновременных подключений именно к SMTP REG.RU.
        with SMTP_SUBMISSION_SEMAPHORE:
            ssl_context = ssl.create_default_context()

            if SMTP_USE_SSL:
                client: smtplib.SMTP = smtplib.SMTP_SSL(
                    SMTP_HOST,
                    SMTP_PORT,
                    timeout=SMTP_TIMEOUT_SEC,
                    context=ssl_context,
                )
            else:
                client = smtplib.SMTP(
                    SMTP_HOST,
                    SMTP_PORT,
                    timeout=SMTP_TIMEOUT_SEC,
                )

            with client as server:
                code, msg = server.ehlo()
                msg_s = _decode_smtp_message(msg)

                dialog.append(f"EHLO -> {code} {msg_s}")

                if not _smtp_success(code):
                    res.code = int(code)
                    res.message = msg_s
                    res.reason = "SMTP server rejected EHLO"
                    res.status = "unknown"
                    return res

                # Этот блок оставлен для SMTP-серверов с STARTTLS.
                # Для mail.hosting.reg.ru:465 SMTP_USE_SSL=True,
                # поэтому STARTTLS не вызывается.
                if not SMTP_USE_SSL:
                    code, msg = server.starttls(context=ssl_context)
                    msg_s = _decode_smtp_message(msg)

                    dialog.append(f"STARTTLS -> {code} {msg_s}")

                    if not _smtp_success(code):
                        res.code = int(code)
                        res.message = msg_s
                        res.reason = "SMTP server rejected STARTTLS"
                        res.status = "unknown"
                        return res

                    code, msg = server.ehlo()
                    msg_s = _decode_smtp_message(msg)

                    dialog.append(
                        f"EHLO after STARTTLS -> {code} {msg_s}"
                    )

                    if not _smtp_success(code):
                        res.code = int(code)
                        res.message = msg_s
                        res.reason = "SMTP server rejected EHLO after STARTTLS"
                        res.status = "unknown"
                        return res

                code, msg = server.login(
                    SMTP_USERNAME,
                    password,
                )
                msg_s = _decode_smtp_message(msg)

                dialog.append(
                    f"AUTH {SMTP_USERNAME} -> {code} {msg_s}"
                )

                code, msg = server.mail(MAIL_FROM)
                msg_s = _decode_smtp_message(msg)

                dialog.append(
                    f"MAIL FROM:<{MAIL_FROM}> -> {code} {msg_s}"
                )

                if not _smtp_success(code):
                    res.code = int(code)
                    res.message = msg_s
                    res.reason = (
                        "Authenticated SMTP server rejected MAIL FROM"
                    )
                    res.status = "invalid"
                    return res

                code, msg = server.rcpt(rcpt_email)
                msg_s = _decode_smtp_message(msg)

                dialog.append(
                    f"RCPT TO:<{rcpt_email}> -> {code} {msg_s}"
                )

                res.code = int(code)
                res.message = msg_s

                if not _smtp_success(code):
                    if 400 <= code < 500:
                        res.status = "temporary"
                        res.reason = (
                            "SMTP submission server temporary rejected "
                            "recipient"
                        )

                    elif 500 <= code < 600:
                        res.status = "invalid"
                        res.reason = (
                            "SMTP submission server permanently rejected "
                            "recipient"
                        )

                    else:
                        res.status = "unknown"
                        res.reason = (
                            f"Unhandled SMTP submission code ({code})"
                        )

                    return res

                # Режим 1: проверка SMTP-конверта без отправки письма.
                if SMTP_CHECK_MODE == "submission_envelope":
                    reset_code, reset_msg = server.rset()
                    reset_msg_s = _decode_smtp_message(reset_msg)

                    dialog.append(
                        f"RSET (without DATA) -> "
                        f"{reset_code} {reset_msg_s}"
                    )

                    res.status = "valid"
                    res.delivery_possible = True
                    res.reason = (
                        "REG.RU accepted SMTP envelope after authentication; "
                        "DATA was not called, no message was sent"
                    )
                    return res

                # Режим 2: фактическая отправка технического письма.
                if SMTP_CHECK_MODE == "submission_send":
                    message = EmailMessage()
                    message["From"] = MAIL_FROM
                    message["To"] = rcpt_email
                    message["Subject"] = SEND_SUBJECT

                    message.set_content(SEND_BODY)

                    refused = server.send_message(
                        message,
                        from_addr=MAIL_FROM,
                        to_addrs=[rcpt_email],
                    )

                    dialog.append(
                        f"DATA/send_message -> refused: {refused}"
                    )

                    if refused:
                        res.status = "invalid"
                        res.reason = (
                            "SMTP server refused message for recipient: "
                            f"{refused}"
                        )
                        return res

                    res.status = "valid"
                    res.delivery_possible = True
                    res.reason = (
                        "REG.RU accepted real message for onward delivery; "
                        "final delivery is not confirmed"
                    )
                    return res

                res.status = "unknown"
                res.reason = (
                    f"Unsupported submission mode: {SMTP_CHECK_MODE}"
                )

    except smtplib.SMTPAuthenticationError as e:
        res.status = "error"
        res.code = int(e.smtp_code) if e.smtp_code else None
        res.message = _decode_smtp_message(e.smtp_error)
        res.reason = (
            "SMTP authentication failed: check SMTP_USERNAME, "
            "SMTP_PASSWORD and mailbox SMTP access"
        )

    except smtplib.SMTPRecipientsRefused as e:
        res.status = "invalid"
        res.reason = f"SMTP recipient refused: {e.recipients}"

    except smtplib.SMTPSenderRefused as e:
        res.status = "invalid"
        res.code = int(e.smtp_code) if e.smtp_code else None
        res.message = _decode_smtp_message(e.smtp_error)
        res.reason = "SMTP server refused MAIL FROM"

    except socket.timeout:
        res.status = "temporary"
        res.reason = "SMTP submission timeout"

    except (
        smtplib.SMTPConnectError,
        smtplib.SMTPServerDisconnected,
        ConnectionRefusedError,
        OSError,
        ssl.SSLError,
    ) as e:
        res.status = "temporary"
        res.reason = (
            f"SMTP submission connection error: "
            f"{type(e).__name__}: {e}"
        )

    except Exception as e:
        res.status = "unknown"
        res.reason = (
            f"SMTP submission error: {type(e).__name__}: {e}"
        )

    finally:
        res.time_taken_sec = round(time.time() - started, 3)

    return res


# ======================
# ЧТЕНИЕ EMAIL ИЗ ФАЙЛА
# ======================

def _email_like(s: str) -> bool:
    if not s:
        return False

    return bool(EMAIL_LIKE_REGEX.search(str(s)))


def _unique_keep_order(values: List[str]) -> List[str]:
    seen = set()
    out: List[str] = []

    for v in values:
        if v not in seen:
            out.append(v)
            seen.add(v)

    return out


def _detect_delimiter(sample: str) -> str:
    try:
        dialect = csv.Sniffer().sniff(
            sample,
            delimiters=[",", ";", "\t", "|"],
        )
        return dialect.delimiter

    except Exception:
        return ";" if sample.count(";") > sample.count(",") else ","


def read_emails_from_file(path: str) -> List[str]:
    """Читает .xlsx/.xls или .csv и возвращает список email.

    Логика:
    - Заголовок может быть любым или отсутствовать.
    - Определяем столбец с email по содержимому.
    - Если email есть в нескольких столбцах — выбираем лучший столбец.
    """

    p = Path(path)

    if not p.exists():
        raise FileNotFoundError(path)

    ext = p.suffix.lower()

    if ext in (".xlsx", ".xls"):
        df = pd.read_excel(
            path,
            header=None,
            dtype=str,
        )

        if df.shape[1] == 0:
            return []

        sample = df.head(DETECT_ROWS)

        best_col = None
        best_score = -1

        for col in df.columns:
            score = int(
                sample[col]
                .astype(str)
                .map(_email_like)
                .sum()
            )

            if score > best_score:
                best_score = score
                best_col = col

        if best_col is None or best_score <= 0:
            emails: List[str] = []

            for col in df.columns:
                emails.extend(
                    [
                        x.strip()
                        for x in df[col]
                        .dropna()
                        .astype(str)
                        .tolist()
                        if _email_like(x)
                    ]
                )

            return _unique_keep_order(
                [email.lower() for email in emails]
            )

        col_vals = (
            df[best_col]
            .dropna()
            .astype(str)
            .map(lambda x: x.strip())
            .tolist()
        )

        emails = [
            value
            for value in col_vals
            if _email_like(value)
        ]

        return _unique_keep_order(
            [email.lower() for email in emails]
        )

    if ext == ".csv":
        with p.open(
            "r",
            encoding="utf-8-sig",
            errors="ignore",
            newline="",
        ) as f:
            sample_text = f.read(4096)
            f.seek(0)

            delimiter = _detect_delimiter(sample_text)
            reader = csv.reader(f, delimiter=delimiter)

            preview_rows: List[List[str]] = []

            for _ in range(DETECT_ROWS):
                try:
                    preview_rows.append(next(reader))
                except StopIteration:
                    break

            if not preview_rows:
                return []

            max_cols = max(
                len(row)
                for row in preview_rows
            )

            scores = [0] * max_cols

            for row in preview_rows:
                for i in range(max_cols):
                    value = row[i] if i < len(row) else ""

                    if _email_like(value):
                        scores[i] += 1

            best_col = int(
                max(
                    range(max_cols),
                    key=lambda i: scores[i],
                )
            )
            best_score = scores[best_col]

            emails: List[str] = []

            def consume_row(row: List[str]) -> None:
                nonlocal emails

                if best_score > 0:
                    if best_col < len(row):
                        value = (row[best_col] or "").strip()

                        if _email_like(value):
                            emails.append(value.lower())

                else:
                    for value in row:
                        value = (value or "").strip()

                        if _email_like(value):
                            emails.append(value.lower())

            for row in preview_rows:
                consume_row(row)

            for row in reader:
                consume_row(row)

            return _unique_keep_order(emails)

    raise ValueError(
        f"Неподдерживаемый формат файла: {ext} "
        f"(нужен .xlsx/.xls или .csv)"
    )


# ======================
# CRASH-SAFE WRITER
# ======================

class ResultWriter:
    """Пишет результаты по мере готовности, append-режим."""

    def __init__(self, run_dir: Path):
        self.run_dir = run_dir
        self.lock = threading.Lock()
        self.count = 0

        self.jsonl_path = run_dir / "checkpoint.jsonl"
        self.checked_path = run_dir / "checked.txt"

        # Старые категории сохранены.
        # Новые добавлены для submission-режимов.
        self.csv_paths = {
            "valid": run_dir / "valid.csv",
            "invalid": run_dir / "invalid.csv",
            "temporary": run_dir / "temporary.csv",
            "unknown": run_dir / "unknown.csv",
            "error": run_dir / "error.csv",
            "submission_accepted": (
                run_dir / "submission_accepted.csv"
            ),
            "submitted": run_dir / "submitted.csv",
        }

        if not self.jsonl_path.exists():
            self.jsonl_path.write_text(
                "",
                encoding="utf-8",
            )

        for path in self.csv_paths.values():
            if not path.exists():
                with path.open(
                    "w",
                    newline="",
                    encoding="utf-8",
                ) as f:
                    csv.writer(f).writerow(["Email"])

        if not self.checked_path.exists():
            self.checked_path.write_text(
                "",
                encoding="utf-8",
            )

        self._jsonl_f = self.jsonl_path.open(
            "a",
            encoding="utf-8",
        )

        self._checked_f = self.checked_path.open(
            "a",
            encoding="utf-8",
        )

        self._csv_f = {
            key: self.csv_paths[key].open(
                "a",
                newline="",
                encoding="utf-8",
            )
            for key in self.csv_paths
        }

        self._csv_w = {
            key: csv.writer(self._csv_f[key])
            for key in self.csv_paths
        }

    def close(self) -> None:
        with self.lock:
            for f in [
                self._jsonl_f,
                self._checked_f,
                *self._csv_f.values(),
            ]:
                try:
                    f.close()
                except Exception:
                    pass

    def write(self, res: EmailResult) -> None:
        with self.lock:
            self.count += 1

            self._jsonl_f.write(
                json.dumps(
                    _result_to_dict(res),
                    ensure_ascii=False,
                )
                + "\n"
            )

            self._checked_f.write(res.email + "\n")

            status = (
                res.final_status
                if res.final_status in self.csv_paths
                else "unknown"
            )

            self._csv_w[status].writerow([res.email])

            if FLUSH_EVERY_RESULT:
                self._jsonl_f.flush()
                self._checked_f.flush()

                for f in self._csv_f.values():
                    f.flush()

            if (
                FSYNC_EVERY_N
                and self.count % FSYNC_EVERY_N == 0
            ):
                try:
                    os.fsync(self._jsonl_f.fileno())
                except Exception:
                    pass

    def rebuild_xlsx_from_csv(
            self,
            ts: str,
            verbose: bool = True,
    ) -> None:
        """Пересобирает XLSX-файлы из CSV."""

        def csv_to_xlsx(
                csv_path: Path,
                out_path: Path,
        ) -> None:
            if not csv_path.exists():
                return

            if verbose:
                print(
                    c(
                        f"  -> Создаю {out_path.name} ...",
                        Fore.CYAN,
                    )
                )

            df = pd.read_csv(csv_path)
            df.to_excel(out_path, index=False)

            if verbose:
                print(
                    c(
                        f"  -> Готово: {out_path.name} "
                        f"(строк: {len(df)})",
                        Fore.GREEN,
                    )
                )

        # Исходные пять Excel-отчётов.
        csv_to_xlsx(
            self.csv_paths["valid"],
            self.run_dir / f"01_valid_emails_{ts}.xlsx",
        )

        csv_to_xlsx(
            self.csv_paths["invalid"],
            self.run_dir / f"02_invalid_emails_{ts}.xlsx",
        )

        csv_to_xlsx(
            self.csv_paths["temporary"],
            self.run_dir / f"03_temporary_emails_{ts}.xlsx",
        )

        csv_to_xlsx(
            self.csv_paths["unknown"],
            self.run_dir / f"04_unknown_emails_{ts}.xlsx",
        )

        csv_to_xlsx(
            self.csv_paths["error"],
            self.run_dir / f"05_error_emails_{ts}.xlsx",
        )

        # Добавленные Excel-отчёты.
        csv_to_xlsx(
            self.csv_paths["submission_accepted"],
            self.run_dir / (
                f"06_submission_accepted_{ts}.xlsx"
            ),
        )

        csv_to_xlsx(
            self.csv_paths["submitted"],
            self.run_dir / f"07_submitted_{ts}.xlsx",
        )

def _result_to_dict(r: EmailResult) -> Dict[str, Any]:
    d = asdict(r)

    d["mx_records"] = [
        asdict(record)
        for record in (r.mx_records or [])
    ]

    d["smtp_checks"] = [
        asdict(check)
        for check in (r.smtp_checks or [])
    ]

    return d

# ======================
# ОСНОВНАЯ ВАЛИДАЦИЯ EMAIL
# ======================

def validate_one(
        email: str,
        email_filter: EmailFilter,
) -> EmailResult:
    t0 = time.time()

    result = EmailResult(
        email=email,
        ts=datetime.now().isoformat(timespec="seconds"),
        mx_records=[],
        smtp_checks=[],
        errors=[],
    )

    if SMTP_CHECK_MODE == "mx_probe":
        jmin, jmax = JITTER_SEC_RANGE
        
        if jmax > 0:
                time.sleep(random.uniform(jmin, jmax))

    # 1. Фильтрация.
    decision = email_filter.check(email)

    if decision.blocked:
        result.filtered = True
        result.filter_reason = decision.reason
        result.final_status = "invalid"
        result.confidence = 0
        result.errors.append(decision.reason)
        result.total_time_sec = round(
            time.time() - t0,
            3,
        )
        return result

    # 2. Синтаксис.
    ok, message = validate_format(email)
    result.format_ok = ok

    if not ok:
        result.final_status = "invalid"
        result.confidence = 0
        result.errors.append(
            f"Syntax error: {message}"
        )
        result.total_time_sec = round(
            time.time() - t0,
            3,
        )
        return result

    # 3. Разбор local-part/domain.
    try:
        result.user, result.domain = parse_email(email)

    except Exception:
        result.final_status = "invalid"
        result.confidence = 0
        result.errors.append("Cannot parse email")
        result.total_time_sec = round(
            time.time() - t0,
            3,
        )
        return result

    # 4. MX нужен только для mx_probe
    if SMTP_CHECK_MODE == "mx_probe":
        mx = get_mx(result.domain)

        if not mx:
            result.mx_ok = False
            result.final_status = "invalid"
            result.confidence = 0
            result.errors.append(
                "No MX records or domain does not exist"
            )
            result.total_time_sec = round(
                time.time() - t0,
                3,
            )
            return result

        result.mx_ok = True
        result.mx_records = mx

        mx_to_check = mx[
            :max(
                1,
                min(MAX_MX_SERVERS, len(mx)),
            )
        ]
    else:
        # Для submission_envelope / submission_send MX не нужен
        result.mx_ok = True
        result.mx_records = []
        mx_to_check = []

    def run_checks_once() -> List[SMTPCheck]:
        # Исходный режим: каждый MX-сервер проверяется напрямую.
        if SMTP_CHECK_MODE == "mx_probe":
            checks: List[SMTPCheck] = []

            for record in mx_to_check:
                checks.append(
                    smtp_probe(
                        record.host,
                        record.priority,
                        email,
                    )
                )

            return checks

        # Новый режим: проверка через SMTP submission REG.RU.
        if SMTP_CHECK_MODE == "submission_envelope":
            return [smtp_submission_probe(email)]

        # Новый режим: реальная отправка через SMTP submission REG.RU.
        if SMTP_CHECK_MODE == "submission_send":
            return [smtp_submission_probe(email)]

        return [
            SMTPCheck(
                mx_host="",
                priority=0,
                status="unknown",
                reason=(
                    "Unknown SMTP_CHECK_MODE: "
                    f"{SMTP_CHECK_MODE}"
                ),
                smtp_dialog=[],
            )
        ]

    checks = run_checks_once()

    # Повтор только временных ошибок.
    if RETRY_ON_TEMPORARY:
        need_retry = any(
            check.status == "temporary"
            for check in checks
        )

        tries = 0

        while need_retry and tries < RETRY_COUNT:
            tries += 1

            time.sleep(RETRY_DELAY_SEC)

            checks = run_checks_once()

            need_retry = any(
                check.status == "temporary"
                for check in checks
            )

    result.smtp_checks = checks

    # Старые статусы сохраняются, новые добавлены отдельно.
    valid_count = sum(
        1
        for check in checks
        if check.status == "valid"
    )

    invalid_count = sum(
        1
        for check in checks
        if check.status == "invalid"
    )

    temporary_count = sum(
        1
        for check in checks
        if check.status == "temporary"
    )

    submission_accepted_count = sum(
        1
        for check in checks
        if check.status == "submission_accepted"
    )

    submitted_count = sum(
        1
        for check in checks
        if check.status == "submitted"
    )

    if submitted_count > 0:
        result.final_status = "submitted"
        result.confidence = 70

    elif submission_accepted_count > 0:
        result.final_status = "submission_accepted"
        result.confidence = 50

    elif valid_count > 0:
        result.final_status = "valid"
        result.confidence = 100

    elif invalid_count > 0 and temporary_count == 0:
        result.final_status = "invalid"
        result.confidence = 95

    elif temporary_count > 0 and invalid_count == 0:
        result.final_status = "temporary"
        result.confidence = 50

    else:
        result.final_status = "unknown"
        result.confidence = 30

    result.total_time_sec = round(
        time.time() - t0,
        3,
    )

    return result

# ======================
# SUMMARY
# ======================

def write_summary(
        run_dir: Path,
        ts: str,
        total: int,
) -> None:
    def count_rows(csv_path: Path) -> int:
        if not csv_path.exists():
            return 0

        n = 0

        with csv_path.open(
                "r",
                encoding="utf-8",
                errors="ignore",
        ) as f:
            for i, _ in enumerate(f):
                n = i

        return max(0, n)

    valid_n = count_rows(run_dir / "valid.csv")
    invalid_n = count_rows(run_dir / "invalid.csv")
    temp_n = count_rows(run_dir / "temporary.csv")
    unknown_n = count_rows(run_dir / "unknown.csv")
    error_n = count_rows(run_dir / "error.csv")

    submission_accepted_n = count_rows(
        run_dir / "submission_accepted.csv"
    )

    submitted_n = count_rows(
        run_dir / "submitted.csv"
    )

    summary_path = run_dir / f"00_summary_{ts}.txt"

    with summary_path.open(
            "w",
            encoding="utf-8",
    ) as f:
        f.write("Email validation summary\n")
        f.write(f"Timestamp: {ts}\n")
        f.write(f"Input file: {INPUT_FILE}\n")
        f.write(
            f"Total in input (unique detected): {total}\n"
        )
        f.write(f"Threads: {NUM_THREADS}\n")
        f.write(f"SMTP mode: {SMTP_CHECK_MODE}\n")
        f.write(f"SMTP timeout: {SMTP_TIMEOUT_SEC}s\n")
        f.write(f"Build Excel: {BUILD_EXCEL_FILES}\n")

        if SMTP_CHECK_MODE in (
                "submission_envelope",
                "submission_send",
        ):
            f.write(f"Submission SMTP host: {SMTP_HOST}\n")
            f.write(f"Submission SMTP port: {SMTP_PORT}\n")
            f.write(f"Submission SMTP SSL: {SMTP_USE_SSL}\n")
            f.write(
                "Submission parallel connections: "
                f"{SMTP_SUBMISSION_MAX_CONNECTIONS}\n"
            )

        f.write("\n")

        f.write(f"Valid: {valid_n}\n")
        f.write(f"Invalid: {invalid_n}\n")
        f.write(f"Temporary: {temp_n}\n")
        f.write(f"Unknown: {unknown_n}\n")
        f.write(f"Error: {error_n}\n")
        f.write(
            "Submission accepted (without DATA): "
            f"{submission_accepted_n}\n"
        )
        f.write(
            "Submitted (real message accepted by REG.RU): "
            f"{submitted_n}\n"
        )

# ======================
# MAIN
# ======================

def main() -> None:
    global REBUILD_XLSX_EVERY_N
    global FINAL_BUILD_XLSX

    if USE_COLORS:
        init(autoreset=True)

    # Загружает SMTP_PASSWORD из .env.
    load_dotenv()

    allowed_modes = {
        "mx_probe",
        "submission_envelope",
        "submission_send",
    }

    if SMTP_CHECK_MODE not in allowed_modes:
        print(
            c(
                "Неизвестный SMTP_CHECK_MODE: "
                f"{SMTP_CHECK_MODE}\n"
                "Допустимо: mx_probe, "
                "submission_envelope, submission_send",
                Fore.RED,
            )
        )
        return

    # Для submission-режимов пароль обязателен.
    if SMTP_CHECK_MODE in (
            "submission_envelope",
            "submission_send",
    ):
        password = os.getenv(
            SMTP_PASSWORD_ENV,
            "",
        ).strip()

        if not password:
            print(
                c(
                    f"Не задан {SMTP_PASSWORD_ENV}.\n"
                    "Создайте файл .env рядом со скриптом:\n"
                    f"{SMTP_PASSWORD_ENV}=пароль_ящика",
                    Fore.RED,
                )
            )
            return

    if SMTP_CHECK_MODE == "submission_send":
        print(
            c(
                "ВНИМАНИЕ: submission_send отправит "
                "реальное письмо каждому адресу "
                "из входного файла.",
                Fore.RED,
            )
        )

    if not BUILD_EXCEL_FILES:
        REBUILD_XLSX_EVERY_N = 0
        FINAL_BUILD_XLSX = False

    if not os.path.exists(INPUT_FILE):
        print(
            c(
                f"Файл не найден: {INPUT_FILE}",
                Fore.RED,
            )
        )
        return

    try:
        emails = read_emails_from_file(INPUT_FILE)

    except Exception as e:
        print(
            c(
                f"Не удалось прочитать входной файл: {e}",
                Fore.RED,
            )
        )
        return

    total = len(emails)

    if total == 0:
        print(
            c(
                "Не найдено email адресов во входном файле",
                Fore.RED,
            )
        )
        return

    ts = datetime.now().strftime(
        "%Y-%m-%d_%H-%M-%S"
    )

    run_dir = Path(OUTPUT_DIR) / ts
    run_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    email_filter = EmailFilter(FILTER_CONFIG_PATH)
    writer = ResultWriter(run_dir)

    print(
        c(
            f"Входной файл: {INPUT_FILE}",
            Fore.BLUE,
        )
    )

    print(
        c(
            f"Найдено email (уникальных): {total}",
            Fore.BLUE,
        )
    )

    print(
        c(
            "Проверка: "
            f"mode={SMTP_CHECK_MODE} | "
            f"потоков={NUM_THREADS} | "
            f"SMTP timeout={SMTP_TIMEOUT_SEC}s",
            Fore.MAGENTA,
        )
    )

    if SMTP_CHECK_MODE in (
            "submission_envelope",
            "submission_send",
    ):
        print(
            c(
                "SMTP submission: "
                f"{SMTP_HOST}:{SMTP_PORT} | "
                f"SSL={SMTP_USE_SSL} | "
                f"max connections="
                f"{SMTP_SUBMISSION_MAX_CONNECTIONS}",
                Fore.MAGENTA,
            )
        )

    print(
        c(
            f"Excel отчёты: {BUILD_EXCEL_FILES}",
            Fore.MAGENTA,
        )
    )

    print(
        c(
            "Результаты пишутся сразу "
            "(checkpoint.jsonl + CSV) в: "
            f"{run_dir}",
            Fore.GREEN,
        )
    )

    start = time.time()
    done = 0

    try:
        with ThreadPoolExecutor(
                max_workers=NUM_THREADS
        ) as executor:
            futures = {
                executor.submit(
                    validate_one,
                    email,
                    email_filter,
                ): email
                for email in emails
            }

            for future in as_completed(futures):
                done += 1
                email = futures[future]

                try:
                    result = future.result()

                except Exception as exc:
                    result = EmailResult(
                        email=email,
                        ts=datetime.now().isoformat(
                            timespec="seconds"
                        ),
                        final_status="error",
                        confidence=0,
                        format_ok=False,
                        mx_ok=False,
                        mx_records=[],
                        smtp_checks=[],
                        errors=[
                            f"Unhandled exception: {exc}"
                        ],
                    )

                writer.write(result)

                if VERBOSE_PER_EMAIL:
                    color = {
                        "valid": Fore.GREEN,
                        "submitted": Fore.GREEN,
                        "submission_accepted": Fore.CYAN,
                        "invalid": Fore.RED,
                        "temporary": Fore.YELLOW,
                        "unknown": Fore.CYAN,
                        "error": Fore.RED,
                    }.get(
                        result.final_status,
                        Fore.WHITE,
                    )

                    print(
                        c(
                            f"{result.final_status.upper():22} "
                            f"{result.email}",
                            color,
                        )
                    )

                if (
                        done % PROGRESS_EVERY == 0
                        or done == total
                ):
                    elapsed = time.time() - start
                    average = elapsed / done
                    remaining = average * (total - done)

                    print(
                        c(
                            f"{done}/{total} "
                            f"({done * 100 / total:.1f}%), "
                            f"осталось ~"
                            f"{timedelta(seconds=int(remaining))}",
                            Fore.BLUE,
                        )
                    )

                if (
                        BUILD_EXCEL_FILES
                        and REBUILD_XLSX_EVERY_N
                        and done % REBUILD_XLSX_EVERY_N == 0
                ):
                    print(
                        c(
                            "Финализация (промежуточно): "
                            "пересборка Excel из CSV...",
                            Fore.CYAN,
                        )
                    )

                    heartbeat = Heartbeat(
                        "Идёт пересборка Excel",
                        FINALIZATION_HEARTBEAT_SEC,
                    )

                    heartbeat.start()

                    try:
                        writer.rebuild_xlsx_from_csv(
                            ts,
                            verbose=False,
                        )

                    finally:
                        heartbeat.stop()

                    print(
                        c(
                            "Промежуточная пересборка "
                            "Excel завершена.",
                            Fore.GREEN,
                        )
                    )

    finally:
        print(
            c(
                "\n100% проверок завершено. "
                "Начинаю финализацию результатов...",
                Fore.MAGENTA,
            )
        )

        if FINAL_BUILD_XLSX:
            print(
                c(
                    "Шаг 1/3: Финальная пересборка "
                    "Excel файлов из CSV.",
                    Fore.CYAN,
                )
            )

            heartbeat = Heartbeat(
                "Финальная пересборка Excel",
                FINALIZATION_HEARTBEAT_SEC,
            )

            heartbeat.start()

            try:
                writer.rebuild_xlsx_from_csv(
                    ts,
                    verbose=True,
                )

            except Exception as rebuild_error:
                logger.warning(
                    "Final XLSX rebuild failed: %s",
                    rebuild_error,
                )

            finally:
                heartbeat.stop()

            print(
                c(
                    "Шаг 1/3 завершён.",
                    Fore.GREEN,
                )
            )

        else:
            print(
                c(
                    "Шаг 1/3: Пропущено "
                    "(Excel отключён).",
                    Fore.YELLOW,
                )
            )

        print(
            c(
                "Шаг 2/3: Создаю summary файл...",
                Fore.CYAN,
            )
        )

        heartbeat = Heartbeat(
            "Запись summary",
            FINALIZATION_HEARTBEAT_SEC,
        )

        heartbeat.start()

        try:
            write_summary(
                run_dir,
                ts,
                total,
            )

        except Exception as summary_error:
            logger.warning(
                "Summary write failed: %s",
                summary_error,
            )

        finally:
            heartbeat.stop()

        print(
            c(
                "Шаг 2/3 завершён.",
                Fore.GREEN,
            )
        )

        print(
            c(
                "Шаг 3/3: Закрываю файлы "
                "и освобождаю ресурсы...",
                Fore.CYAN,
            )
        )

        heartbeat = Heartbeat(
            "Закрытие файлов",
            FINALIZATION_HEARTBEAT_SEC,
        )

        heartbeat.start()

        try:
            writer.close()

        finally:
            heartbeat.stop()

        print(
            c(
                "Шаг 3/3 завершён.",
                Fore.GREEN,
            )
        )

    elapsed_total = time.time() - start

    print(
        c(
            f"\nГотово за "
            f"{timedelta(seconds=int(elapsed_total))}",
            Fore.MAGENTA,
        )
    )

    print(
        c(
            f"Папка результатов: {run_dir}",
            Fore.GREEN,
        )
    )

    print(
        c(
            "Если программа упадёт — частичные "
            "результаты останутся в checkpoint.jsonl и CSV.",
            Fore.GREEN,
        )
    )

if __name__ == "__main__":
    main()
