# -*- coding: utf-8 -*-
"""
filters.py — модуль фильтрации email перед дорогостоящими DNS/SMTP проверками.

Задачи:
- Отсеивать заведомо нецелевые ящики (no-reply, postmaster, abuse, unsubscribe и т.п.)
- Поддерживать пользовательские списки блокировки (например: hr, ceo, accounting...)
- Уметь объяснять причину (reason) — для отчёта и отладки

Основа дефолтного списка role-based/служебных локальных частей взята из публичного списка,
который MailPoet игнорирует (role-based email addresses) [page:0] — см. README.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple


DEFAULT_CONFIG_PATH = "filters_config.json"


@dataclass
class FilterDecision:
    blocked: bool
    reason: str = ""


def _norm_local(local: str) -> str:
    return (local or "").strip().lower()


def _norm_domain(domain: str) -> str:
    return (domain or "").strip().lower().rstrip(".")


class EmailFilter:
    def __init__(self, config_path: str = DEFAULT_CONFIG_PATH):
        self.config_path = config_path
        self.cfg: Dict = {}
        self._compiled_local_regex: List[re.Pattern] = []
        self._compiled_domain_regex: List[re.Pattern] = []

        self.load()

    def load(self) -> None:
        p = Path(self.config_path)
        if not p.exists():
            # Если конфига нет — создаём минимальный "пустой" режим
            self.cfg = {"enabled": False}
            self._compiled_local_regex = []
            self._compiled_domain_regex = []
            return

        self.cfg = json.loads(p.read_text(encoding="utf-8"))
        self._compiled_local_regex = [
            re.compile(x, re.IGNORECASE) for x in (self.cfg.get("block_local_regex") or [])
        ] + [
            re.compile(x, re.IGNORECASE) for x in (self.cfg.get("user_block_local_regex") or [])
        ]

        self._compiled_domain_regex = [
            re.compile(x, re.IGNORECASE) for x in (self.cfg.get("block_domain_regex") or [])
        ] + [
            re.compile(x, re.IGNORECASE) for x in (self.cfg.get("user_block_domain_regex") or [])
        ]

    def is_enabled(self) -> bool:
        return bool(self.cfg.get("enabled", True))

    def check(self, email: str) -> FilterDecision:
        """
        Возвращает blocked=True если email надо сразу отправить в invalid по фильтру,
        чтобы не тратить время на DNS/SMTP.
        """
        if not self.is_enabled():
            return FilterDecision(False, "")

        if not email or "@" not in email:
            # формат отловит основной валидатор; здесь только мягкая защита
            return FilterDecision(False, "")

        local, domain = email.split("@", 1)
        local_n = _norm_local(local)
        domain_n = _norm_domain(domain)

        # 1) exact local block (default + user)
        block_local_exact = set(map(_norm_local, self.cfg.get("block_local_exact") or []))
        user_block_local_exact = set(map(_norm_local, self.cfg.get("user_block_local_exact") or []))
        if local_n in block_local_exact:
            return FilterDecision(True, f"Filtered: local-part '{local_n}' is role/system (exact)")
        if local_n in user_block_local_exact:
            return FilterDecision(True, f"Filtered: local-part '{local_n}' blocked by user (exact)")

        # 2) local contains
        for sub in (self.cfg.get("block_local_contains") or []):
            sub_n = _norm_local(sub)
            if sub_n and sub_n in local_n:
                return FilterDecision(True, f"Filtered: local-part contains '{sub_n}'")

        for sub in (self.cfg.get("user_block_local_contains") or []):
            sub_n = _norm_local(sub)
            if sub_n and sub_n in local_n:
                return FilterDecision(True, f"Filtered: local-part contains '{sub_n}' (user)")

        # 3) local regex
        for rx in self._compiled_local_regex:
            if rx.search(local_n):
                return FilterDecision(True, f"Filtered: local-part matched regex '{rx.pattern}'")

        # 4) domain exact/suffix/regex (редко нужно, но добавлено как вы просили)
        block_domain_exact = set(map(_norm_domain, self.cfg.get("block_domain_exact") or []))
        user_block_domain_exact = set(map(_norm_domain, self.cfg.get("user_block_domain_exact") or []))
        if domain_n in block_domain_exact:
            return FilterDecision(True, f"Filtered: domain '{domain_n}' blocked (exact)")
        if domain_n in user_block_domain_exact:
            return FilterDecision(True, f"Filtered: domain '{domain_n}' blocked by user (exact)")

        for suf in (self.cfg.get("block_domain_suffix") or []):
            suf_n = _norm_domain(suf)
            if suf_n and domain_n.endswith(suf_n):
                return FilterDecision(True, f"Filtered: domain endswith '{suf_n}'")

        for suf in (self.cfg.get("user_block_domain_suffix") or []):
            suf_n = _norm_domain(suf)
            if suf_n and domain_n.endswith(suf_n):
                return FilterDecision(True, f"Filtered: domain endswith '{suf_n}' (user)")

        for rx in self._compiled_domain_regex:
            if rx.search(domain_n):
                return FilterDecision(True, f"Filtered: domain matched regex '{rx.pattern}'")

        return FilterDecision(False, "")
