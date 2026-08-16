# Email_Validator — многопоточная проверка email

Проект предназначен для технической проверки базы email-адресов перед работой с согласованными контактами.

Проверка выполняется в несколько этапов:

1. Проверка синтаксиса адреса через библиотеку `email-validator`
2. Проверка домена через DNS и MX-записи
3. SMTP-проверка в выбранном режиме
4. Фильтрация служебных и нецелевых адресов
5. Crash-safe сохранение результатов в процессе работы

> Важно: ни DNS/MX, ни ответ SMTP `250` не дают 100% гарантии, что письмо попадёт во «Входящие». Почтовые серверы могут применять catch-all, антиспам-защиту, greylisting, отложенную проверку адреса или не раскрывать существование ящика на этапе `RCPT TO`.

---

## Состав проекта

| Файл | Назначение |
|---|---|
| `smtp_email_checker.py` | Основной многопоточный скрипт: чтение базы, синтаксис, DNS/MX, SMTP, отчёты и сохранение прогресса |
| `filters.py` | Фильтрация служебных, role-based и нежелательных адресов |
| `filters_config.json` | Настройки фильтрации адресов и доменов |
| `rebuild_reports.py` | Пересборка Excel-отчётов из `checkpoint.jsonl` после остановки или падения |
| `requirements.txt` | Зависимости Python |
| `.env.example` | Пример переменных окружения для SMTP-настроек |
| `.env` | Локальный файл с секретами, не должен попадать в Git |
| `results/` | Папка с результатами запусков |

---

## Требования

- Windows 10/11, Linux или macOS
- Python 3.9 или новее
- Доступ в интернет для DNS- и SMTP-проверок
- Для submission-режимов — настроенный SMTP-ящик

Проверить версию Python:

```bash
python --version
```

---

## Установка

Создайте и активируйте виртуальное окружение.

### Windows PowerShell

```powershell
cd C:\Users\user\PycharmProjects\Email_Validator

python -m venv .venv

.\.venv\Scripts\Activate.ps1
```

Установите зависимости:

```powershell
pip install -r requirements.txt
```

### Windows CMD

```bat
cd C:\Users\user\PycharmProjects\Email_Validator

python -m venv .venv

.venv\Scripts\activate.bat

pip install -r requirements.txt
```

> Не называйте основной файл `email_validator.py`, поскольку это может конфликтовать с пакетом `email-validator`.

---

## Входной файл

В переменной `INPUT_FILE` в начале `smtp_email_checker.py` указывается путь к исходной базе:

```python
INPUT_FILE = str(
    Path(r"C:\Users\user\Downloads\MailPoet_export.csv")
)
```

Поддерживаются форматы:

- Excel: `.xlsx`
- Excel: `.xls`
- CSV: `.csv`

### Автоматическое определение столбца

Программа не зависит от названия колонки.

Она анализирует первые строки файла и выбирает столбец, где чаще всего встречаются значения, похожие на email-адреса:

```text
name@example.com
sales@company.ru
info@factory.com
```

Поэтому допустимы разные заголовки:

```text
Email
Почта
Contact
E-mail
Значение
Адрес
```

Допускается и файл без заголовка, где первая строка уже содержит email.

### Определение разделителя CSV

Для CSV программа пытается автоматически определить разделитель:

```text
,
;
Tab
|
```

---

## Режимы SMTP-проверки

Режим выбирается в настройках файла `smtp_email_checker.py`:

```python
SMTP_CHECK_MODE = "mx_probe"
```

Доступны три режима.

### `mx_probe`

Исходный режим прямой SMTP-проверки через MX-серверы получателя.

```python
SMTP_CHECK_MODE = "mx_probe"
```

Сценарий:

```text
DNS MX
↓
MX-сервер получателя:25
↓
HELO
↓
MAIL FROM
↓
RCPT TO
```

Пример:

```text
Ваш ПК
↓
gmail-smtp-in.l.google.com:25
↓
HELO mailchecker.zehvk.ru
↓
MAIL FROM:<info@zehvk.ru>
↓
RCPT TO:<recipient@example.com>
```

Преимущества:

- Не требует SMTP-пароля
- Не использует SMTP-лимиты вашего почтового ящика
- Позволяет получить ответ непосредственно от принимающего MX-сервера
- Сохраняет исходную логику проекта

Ограничения:

- Порт `25` может блокироваться провайдером, роутером, антивирусом или корпоративной сетью
- Серверы могут временно отвечать кодами `4xx`
- Некоторые сервисы не раскрывают существование ящика
- Catch-all-домены могут принимать письма на несуществующие адреса

---

### `submission_envelope`

Проверка через авторизованный SMTP-шлюз без реальной отправки письма.

```python
SMTP_CHECK_MODE = "submission_envelope"
```

Сценарий для REG.RU:

```text
mail.hosting.reg.ru:465
↓
SSL/TLS
↓
EHLO
↓
SMTP AUTH
↓
MAIL FROM
↓
RCPT TO
↓
RSET
```

После команды `RCPT TO` скрипт отправляет `RSET`, поэтому команда `DATA` не вызывается и письмо не формируется.

Преимущества:

- Используется легитимный SMTP-шлюз вашего почтового ящика
- Работает, если прямой порт `25` заблокирован
- Проверяет возможность выполнения SMTP-конверта через ваш сервер
- Не отправляет тестовое письмо в режиме `submission_envelope`

Ограничения:

- Требует пароль SMTP-ящика
- Может расходовать лимиты SMTP-подключений
- Принятие `RCPT TO` SMTP-шлюзом не подтверждает окончательную доставку
- Некоторые почтовые системы принимают адрес на раннем этапе, а затем проверяют его позже

---

### `submission_send`

Режим реальной отправки технического письма.

```python
SMTP_CHECK_MODE = "submission_send"
```

Сценарий:

```text
mail.hosting.reg.ru:465
↓
SSL/TLS
↓
EHLO
↓
SMTP AUTH
↓
MAIL FROM
↓
RCPT TO
↓
DATA
↓
Отправка письма
```

> ⚠️ Внимание: этот режим отправляет реальное сообщение каждому адресу во входном файле. Используйте его только для базы контактов, по которой у вас есть законные основания и согласие на получение сообщений.

Тема и текст технического письма задаются в настройках:

```python
SEND_SUBJECT = "Проверка доставки"

SEND_BODY = (
    "Здравствуйте!\n\n"
    "Это техническое сообщение для проверки доставки на адрес, "
    "указанный в подписке.\n"
)
```

---

## SMTP-настройки REG.RU

Для SMTP submission-режимов используются следующие настройки:

```python
SMTP_HOST = "mail.hosting.reg.ru"
SMTP_PORT = 465
SMTP_USE_SSL = True
SMTP_USERNAME = "info@zehvk.ru"
MAIL_FROM = "info@zehvk.ru"
```

Параметры соответствуют типовой настройке доменного ящика REG.RU:

```text
SMTP server: mail.hosting.reg.ru
SMTP port: 465
Security: SSL/TLS
Username: полный email-адрес
Password: пароль почтового ящика
```

Для ограничения нагрузки на SMTP-сервер используется семафор:

```python
SMTP_SUBMISSION_MAX_CONNECTIONS = 3
```

Даже если основной скрипт использует:

```python
NUM_THREADS = 15
```

одновременно будет открыто не более трёх SMTP-подключений к `mail.hosting.reg.ru`.

---

## Настройка `.env`

Для режимов `submission_envelope` и `submission_send` необходимо создать файл `.env` в корне проекта.

Путь:

```text
C:\Users\user\PycharmProjects\Email_Validator\.env
```

Содержимое:

```env
SMTP_PASSWORD=ваш_реальный_пароль_от_info@zehvk.ru
```

При необходимости можно задать SMTP-логин отдельно:

```env
SMTP_USERNAME=info@zehvk.ru
SMTP_PASSWORD=ваш_реальный_пароль_от_info@zehvk.ru
```

### Безопасность секретов

Никогда не добавляйте `.env` в Git.

Добавьте в `.gitignore`:

```gitignore
.env
.venv/
.idea/
__pycache__/
results/
```

Пароль не должен находиться в:

- `smtp_email_checker.py`
- `filters_config.json`
- Excel-файлах
- CSV-отчётах
- Git-репозитории
- скриншотах
- сообщениях в мессенджерах
- логах

---

## Запуск

### Через активированное виртуальное окружение

```powershell
python .\smtp_email_checker.py
```

### Явный запуск через интерпретатор `.venv`

```powershell
.\.venv\Scripts\python.exe .\smtp_email_checker.py
```

---

## Результаты работы

После каждого запуска создаётся отдельная папка:

```text
results/<YYYY-mm-dd_HH-MM-SS>/
```

Пример:

```text
results/2026-08-16_15-00-00/
```

### Crash-safe файлы

Результаты записываются по мере готовности. При аварийной остановке или падении уже готовые данные остаются в папке запуска.

| Файл | Назначение |
|---|---|
| `checkpoint.jsonl` | Полный результат по каждому адресу, одна JSON-строка на один email |
| `checked.txt` | Список всех обработанных email |
| `valid.csv` | Адреса с успешным результатом в режиме `mx_probe` |
| `invalid.csv` | Адреса с постоянным отказом, ошибкой формата или отсутствующим MX |
| `temporary.csv` | Временные SMTP/DNS ошибки |
| `unknown.csv` | Неоднозначные результаты |
| `error.csv` | Непредвиденные ошибки выполнения |
| `submission_accepted.csv` | SMTP-конверт принят шлюзом в режиме `submission_envelope` |
| `submitted.csv` | Реальное письмо принято SMTP-шлюзом в режиме `submission_send` |
| `00_summary_<timestamp>.txt` | Итоговая сводка по запуску |

---

## Статусы результатов

| Статус | Значение |
|---|---|
| `valid` | Хотя бы один MX-сервер принял `RCPT TO` в режиме `mx_probe` |
| `invalid` | Неверный синтаксис, нет MX, домен не существует или сервер вернул постоянный SMTP-отказ |
| `temporary` | Временная ошибка SMTP/DNS, код `4xx`, таймаут или временная недоступность |
| `unknown` | Невозможно однозначно определить результат |
| `error` | Ошибка конфигурации, например отсутствует `SMTP_PASSWORD` |
| `submission_accepted` | SMTP-шлюз после авторизации принял конверт, но письмо не отправлялось |
| `submitted` | SMTP-шлюз принял реальное письмо для последующей доставки |

> `submission_accepted` и `submitted` не означают гарантированное попадание сообщения во «Входящие». Для контроля окончательной доставки нужно учитывать bounce-уведомления, DSN, отписки и историю предыдущих отправок.

---

## Excel-отчёты

Создание `.xlsx` на больших базах может занимать заметное время. Поэтому по умолчанию Excel отключён:

```python
BUILD_EXCEL_FILES = False
```

Основные быстрые результаты доступны сразу в CSV и JSONL.

Чтобы включить Excel-отчёты, измените настройку:

```python
BUILD_EXCEL_FILES = True
```

При включённой опции создаются файлы вида:

```text
01_valid_emails_<timestamp>.xlsx
02_invalid_emails_<timestamp>.xlsx
03_temporary_emails_<timestamp>.xlsx
04_unknown_emails_<timestamp>.xlsx
05_error_emails_<timestamp>.xlsx
06_submission_accepted_<timestamp>.xlsx
07_submitted_<timestamp>.xlsx
```

---

## Восстановление отчётов

Если программа была прервана, Excel-файлы можно пересобрать из `checkpoint.jsonl`.

```powershell
.\.venv\Scripts\python.exe .\rebuild_reports.py .\results\<timestamp>
```

Пример:

```powershell
.\.venv\Scripts\python.exe .\rebuild_reports.py .\results\2026-08-16_15-00-00
```

---

## Фильтрация адресов

Перед DNS- и SMTP-проверками программа использует `filters.py`.

Фильтр позволяет не тратить сетевые запросы на явно служебные или нежелательные адреса:

```text
no-reply
noreply
postmaster
abuse
unsubscribe
mailer-daemon
```

Настройки находятся в файле:

```text
filters_config.json
```

Пример пользовательской фильтрации:

```json
{
  "user_block_local_exact": [
    "hr",
    "accounting",
    "ceo",
    "noreply"
  ],
  "user_block_domain_suffix": [
    "tempmail.com",
    "10minutemail.com"
  ]
}
```

> Для доменных суффиксов используйте параметр `user_block_domain_suffix`, а не `user_block_domain_endswith`.

---

## Настройка производительности

Основные параметры находятся в начале `smtp_email_checker.py`.

```python
NUM_THREADS = 15
DNS_TIMEOUT_SEC = 5
SMTP_TIMEOUT_SEC = 8
MAX_MX_SERVERS = 3
JITTER_SEC_RANGE = (0.0, 0.25)
```

| Параметр | Назначение |
|---|---|
| `NUM_THREADS` | Число одновременных задач проверки |
| `DNS_TIMEOUT_SEC` | Таймаут DNS-запроса |
| `SMTP_TIMEOUT_SEC` | Таймаут SMTP-подключения |
| `MAX_MX_SERVERS` | Максимум MX-серверов одного домена для проверки |
| `JITTER_SEC_RANGE` | Случайная небольшая задержка перед SMTP-подключением |
| `RETRY_ON_TEMPORARY` | Повторять ли временные ошибки |
| `RETRY_COUNT` | Количество повторных попыток |
| `RETRY_DELAY_SEC` | Пауза перед повторной попыткой |
| `SMTP_SUBMISSION_MAX_CONNECTIONS` | Максимум параллельных соединений с вашим SMTP-шлюзом |

### Рекомендованные параметры

Для обычной базы:

```python
NUM_THREADS = 10
MAX_MX_SERVERS = 2
SMTP_TIMEOUT_SEC = 8
```

Для submission-режима:

```python
NUM_THREADS = 10
SMTP_SUBMISSION_MAX_CONNECTIONS = 2
JITTER_SEC_RANGE = (0.1, 0.5)
```

Для очень большой базы:

```python
BUILD_EXCEL_FILES = False
```

Рекомендуется:

- работать с CSV/JSONL;
- запускать проверку частями;
- не переводить компьютер в спящий режим;
- не использовать слишком большое число потоков;
- сохранять результаты каждого запуска отдельно.

---

## Типичные проблемы

### Ошибка: `SMTP_PASSWORD not set`

Причина: отсутствует `.env` или в нём не указана переменная.

Создайте файл `.env`:

```env
SMTP_PASSWORD=пароль_почтового_ящика
```

Проверьте, что `.env` находится рядом с `smtp_email_checker.py`.

---

### Ошибка SMTP AUTH

Пример:

```text
SMTP authentication failed
```

Проверьте:

- SMTP-логин указан полным email-адресом;
- пароль актуален;
- используется `mail.hosting.reg.ru`;
- порт указан как `465`;
- `SMTP_USE_SSL = True`;
- ящик не заблокирован;
- SMTP-доступ разрешён у почтового провайдера.

---

### DNS timeout

Пример:

```text
DNS timeout for gmail.com
```

Что можно сделать:

```python
DNS_TIMEOUT_SEC = 8
```

или:

```python
DNS_TIMEOUT_SEC = 10
```

Также проверьте:

- стабильность интернет-соединения;
- работу DNS через VPN, если он используется;
- блокировки DNS со стороны антивируса или корпоративной сети.

---

### Порт 25 недоступен

Пример:

```text
SMTP connection error
```

или:

```text
SMTP timeout
```

в режиме:

```python
SMTP_CHECK_MODE = "mx_probe"
```

Возможные причины:

- исходящий порт `25` блокируется провайдером;
- порт блокируется роутером;
- порт блокируется антивирусом;
- корпоративная сеть запрещает SMTP-трафик;
- удалённый сервер не принимает соединения.

В таком случае можно использовать:

```python
SMTP_CHECK_MODE = "submission_envelope"
```

---

### Временные SMTP-отказы

Пример:

```text
Temporary SMTP error (421)
Temporary SMTP error (450)
Temporary SMTP error (451)
Temporary SMTP error (452)
```

Можно включить повторные попытки:

```python
RETRY_ON_TEMPORARY = True
RETRY_COUNT = 2
RETRY_DELAY_SEC = 5
```

И снизить нагрузку:

```python
NUM_THREADS = 5
JITTER_SEC_RANGE = (0.2, 0.8)
```

---

### Ошибка при запуске PowerShell

Если PowerShell блокирует `.ps1`-файлы:

```text
running scripts is disabled on this system
```

Для запуска Python-скрипта используйте явный путь к интерпретатору:

```powershell
.\.venv\Scripts\python.exe .\smtp_email_checker.py
```

---

## Правовая и техническая заметка

Используйте программу только для адресов, по которым у вас есть законные основания для обработки, деловой коммуникации или согласие на получение сообщений.

Режим:

```python
SMTP_CHECK_MODE = "submission_send"
```

создаёт реальные письма. Перед его использованием убедитесь, что:

- адресаты дали согласие на коммуникацию;
- письмо содержит корректный механизм отписки, если это рассылка;
- соблюдены требования к рекламе и персональным данным;
- настроены SPF, DKIM и DMARC для домена отправителя;
- вы готовы обрабатывать bounce-уведомления и отписки.

---

## Лицензия

MIT