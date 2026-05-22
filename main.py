from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from urllib.parse import unquote
from typing import Optional
from pathlib import Path

import importlib.util
import unicodedata
import subprocess
import importlib
import time
import sys
import re
import os

MAX_WORKERS = 35
MAX_RESULTS = 4
TIMEOUT = 4

EMAIL_REGEX = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"
PROJECT_REGEX = re.compile(r"\d{4}-\d-[A-Z]{2}\d{2}-KA1\d{2}-[A-Z]{3}-\d{5,}")

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}

REQUIRED_PACKAGES = {
    "pandas": "pandas",
    "requests": "requests",
    "xlsxwriter": "xlsxwriter",
    "ddgs": "ddgs",
    "questionary": "questionary",
    "openpyxl": "openpyxl",
}

pd = None
requests = None
xlsxwriter = None
DDGS = None
questionary = None

def clear_console() -> None:
    os.system("cls" if os.name == "nt" else "clear")

def load_runtime_dependencies() -> None:
    global pd, requests, xlsxwriter, DDGS, questionary

    pd = importlib.import_module("pandas")
    requests = importlib.import_module("requests")
    xlsxwriter = importlib.import_module("xlsxwriter")
    DDGS = getattr(importlib.import_module("ddgs"), "DDGS")
    questionary = importlib.import_module("questionary")

@dataclass(frozen=True)
class OrganizationRecord:
    source_file: str
    sheet_name: str
    project_number: str
    org_name: str
    org_address: str
    org_city: str

@dataclass(frozen=True)
class EmailLookupResult:
    source_file: str
    sheet_name: str
    project_number: str
    org_name: str
    org_city: str
    email: Optional[str]
    verified: bool

class EmailFinder:
    BLOCKED_TLDS = {
        "png",
        "jpg",
        "jpeg",
        "gif",
        "svg",
        "webp",
        "ico",
        "bmp",
        "tiff",
        "pdf",
        "css",
        "js",
        "json",
        "xml",
        "txt",
        "zip",
        "rar",
        "7z",
    }

    GENERIC_ALLOWED_TLDS = {"com", "net"}

    PLACEHOLDER_DOMAINS = {
        "mysite",
        "mysite.com",
        "example.com",
        "example.org",
        "example.net",
        "yourdomain.com",
        "domain.com",
        "test.com",
        "localhost",
    }

    PLACEHOLDER_DOMAIN_ROOTS = {
        "mysite",
        "example",
        "domain",
        "test",
        "localhost",
    }

    def clean_email(self, email: str) -> Optional[str]:
        if not email:
            return None

        cleaned = unquote(email)
        cleaned = cleaned.replace("mailto:", "")
        cleaned = cleaned.replace("%20", "")
        cleaned = cleaned.strip().strip(".,;:()[]<>\"'")
        return cleaned

    def is_valid_email(self, email: Optional[str]) -> bool:
        if not email or "@" not in email:
            return False

        email = email.strip().lower()
        if not re.fullmatch(EMAIL_REGEX, email):
            return False

        local_part, domain = email.split("@", 1)
        domain = domain.strip().lower()
        local_part = local_part.strip().lower()

        if len(local_part) > 25:
            return False

        if len(local_part) < 2:
            return False

        if re.fullmatch(r"[a-fA-F0-9]{20,}", local_part):
            return False

        if ".." in domain or domain.startswith(".") or domain.endswith("."):
            return False

        if domain.startswith("-") or domain.endswith("-"):
            return False

        labels = domain.split(".")
        if len(labels) < 2:
            return False

        if any(not label or label.startswith("-") or label.endswith("-") for label in labels):
            return False

        tld = labels[-1]
        if tld in self.BLOCKED_TLDS:
            return False

        if domain in self.PLACEHOLDER_DOMAINS:
            return False

        if labels[0] in self.PLACEHOLDER_DOMAIN_ROOTS:
            return False

        return True

    def is_preferred_email(self, email: str) -> bool:
        email_lower = email.lower()
        local_part = email_lower.split("@", 1)[0]
        domain = email_lower.split("@", 1)[1] if "@" in email_lower else ""

        blocked = [
            "noreply",
            "no-reply",
            "support",
            "cloudflare",
            "hostmaster",
            "webmaster",
            "abuse",
        ]
        if any(word in email_lower for word in blocked):
            return False

        if local_part == "admin" and (
            domain in self.PLACEHOLDER_DOMAINS
            or domain.split(".", 1)[0] in self.PLACEHOLDER_DOMAIN_ROOTS
        ):
            return False

        return True

    def is_country_compatible_email(self, email: str, country_code: Optional[str]) -> bool:
        if "@" not in email:
            return False

        domain = email.split("@", 1)[1].strip().lower()
        if "." not in domain:
            return False

        tld = domain.rsplit(".", 1)[-1]
        if not tld:
            return False

        allowed = set(self.GENERIC_ALLOWED_TLDS)
        if country_code:
            code = country_code.strip().upper()
            if code == "SI":
                allowed.update({"si", "sl"})
            elif code == "SL":
                allowed.update({"sl", "si"})
            else:
                allowed.add(code.lower())

        return tld in allowed

    def score_email(self, email: str) -> int:
        email_lower = email.lower()
        local_part = email_lower.split("@", 1)[0]

        score = 0

        strong_keywords = [
            "office",
            "sekretariat",
            "secretary",
            "secretariat",
            "direktion",
            "director",
            "post",
            "info",
            "kontakt",
            "contact",
            "verwaltung",
            "admin",
            "school",
        ]

        for keyword in strong_keywords:
            if keyword in local_part:
                score += 100

        official_domain_parts = [
            "schule",
            "edu",
            "ac.at",
            "gv.at",
            "school",
            "gym",
        ]

        for part in official_domain_parts:
            if part in email_lower:
                score += 20

        if "." in local_part or "_" in local_part:
            score -= 10

        if len(local_part) > 20:
            score -= 20

        return score

    def find_email(self, org_name: str, country_code: Optional[str] = None) -> tuple[Optional[str], bool]:
        found_emails: list[str] = []
        queries = [f"{org_name} email", f'"{org_name}" contact']

        for query in queries:
            try:
                with DDGS() as ddgs:
                    results = list(ddgs.text(query, max_results=MAX_RESULTS))

                for result in results:
                    url = result.get("href", "")
                    text = " ".join([
                        result.get("title", ""),
                        result.get("body", ""),
                        url,
                    ])

                    for email in re.findall(EMAIL_REGEX, text):
                        candidate = self.clean_email(email)
                        if (
                            candidate
                            and self.is_valid_email(candidate)
                            and self.is_preferred_email(candidate)
                            and self.is_country_compatible_email(candidate, country_code)
                        ):
                            if self.score_email(candidate) >= 100:
                                return candidate, True
                            found_emails.append(candidate)

                    if not url:
                        continue

                    try:
                        response = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
                        for email in re.findall(EMAIL_REGEX, response.text):
                            candidate = self.clean_email(email)
                            if (
                                candidate
                                and self.is_valid_email(candidate)
                                and self.is_preferred_email(candidate)
                                and self.is_country_compatible_email(candidate, country_code)
                            ):
                                if self.score_email(candidate) >= 100:
                                    return candidate, True
                                found_emails.append(candidate)
                    except Exception:
                        continue

            except Exception:
                continue

        if found_emails:
            unique = sorted(set(found_emails), key=self.score_email, reverse=True)
            best = unique[0]
            return best, self.score_email(best) >= 100

        return None, False