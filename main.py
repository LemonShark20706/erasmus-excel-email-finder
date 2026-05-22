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

class ExcelStructureDetector:
    PROJECT_KEYWORDS = (
        "palyazat szama",
        "palyazat azonosito",
        "sztevilka projekta",
        "stevilka projekta",
        "st. projekta",
        "st projekta",
        "nr wniosku",
        "project",
    )

    ORG_KEYWORDS = (
        "szervezet neve",
        "palyazo szervezet",
        "organizacija",
        "organizcija",
        "nazwa wnioskodawcy",
        "organization",
    )

    CITY_KEYWORDS = (
        "szekhely",
        "miasto",
        "kraj",
        "city",
    )

    IGNORE_CELL_MARKERS = (
        "tamogatott palyazatok",
        "tartaleklista",
        "table",
        "tabela",
        "zap. st",
        "lp.",
    )

    def extract_records(self, excel_path: Path) -> list[OrganizationRecord]:
        records: list[OrganizationRecord] = []

        try:
            workbook = pd.ExcelFile(excel_path)
        except Exception as exc:
            print(f"[WARN] Nem sikerult megnyitni: {excel_path.name} ({exc})")
            return records

        for sheet_name in workbook.sheet_names:
            sheet_records = self._extract_records_from_sheet(excel_path, sheet_name)
            records.extend(sheet_records)

        return self._deduplicate(records)

    def _extract_records_from_sheet(self, excel_path: Path, sheet_name: str) -> list[OrganizationRecord]:
        try:
            raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, dtype=str)
        except Exception as exc:
            print(f"[WARN] Nem sikerult beolvasni a lapot: {excel_path.name} / {sheet_name} ({exc})")
            return []

        raw = raw.fillna("")
        records: list[OrganizationRecord] = []

        for row_idx in range(len(raw)):
            row_values = [self._clean_cell(value) for value in raw.iloc[row_idx].tolist()]
            project_hit = self._find_project_in_row(row_values)

            if not project_hit:
                continue

            project_col, project_number = project_hit
            header_map = self._find_nearby_header_map(raw, row_idx)

            org_col = self._pick_org_column(row_values, project_col, header_map.get("org"))
            if org_col is None:
                continue

            org_name = self._clean_cell(row_values[org_col])
            if not org_name or self._looks_like_noise(org_name):
                continue

            address = self._pick_address(row_values, org_col, header_map.get("city"))
            city = self._pick_city(row_values, org_col, header_map.get("city"), address)

            records.append(
                OrganizationRecord(
                    source_file=excel_path.name,
                    sheet_name=sheet_name,
                    project_number=project_number,
                    org_name=org_name,
                    org_address=address,
                    org_city=city,
                )
            )

        return records

    def _find_project_in_row(self, row_values: list[str]) -> Optional[tuple[int, str]]:
        for idx, cell in enumerate(row_values):
            project = self._extract_project_id(cell)
            if project:
                return idx, project
        return None

    def _extract_project_id(self, text: str) -> Optional[str]:
        if not text:
            return None

        compact = re.sub(r"\s+", "", text.upper())
        match = PROJECT_REGEX.search(compact)
        if not match:
            return None

        return match.group(0)

    def _find_nearby_header_map(self, raw: pd.DataFrame, row_idx: int) -> dict[str, int]:
        best_map: dict[str, int] = {}
        best_score = 0

        for idx in range(max(0, row_idx - 8), row_idx):
            values = [self._normalize_text(v) for v in raw.iloc[idx].tolist()]
            mapping: dict[str, int] = {}

            for col, cell in enumerate(values):
                if not cell:
                    continue

                if "project" not in mapping and any(keyword in cell for keyword in self.PROJECT_KEYWORDS):
                    mapping["project"] = col
                if "org" not in mapping and any(keyword in cell for keyword in self.ORG_KEYWORDS):
                    mapping["org"] = col
                if "city" not in mapping and any(keyword in cell for keyword in self.CITY_KEYWORDS):
                    mapping["city"] = col

            score = len(mapping)
            if score > best_score:
                best_map = mapping
                best_score = score

        return best_map

    def _pick_org_column(self, row_values: list[str], project_col: int, header_org_col: Optional[int]) -> Optional[int]:
        if header_org_col is not None and header_org_col < len(row_values):
            candidate = self._clean_cell(row_values[header_org_col])
            if self._is_org_candidate(candidate):
                return header_org_col

        for col in range(project_col + 1, len(row_values)):
            candidate = self._clean_cell(row_values[col])
            if self._is_org_candidate(candidate):
                return col

        return None

    def _pick_address(self, row_values: list[str], org_col: int, city_col: Optional[int]) -> str:
        if city_col is not None and city_col < len(row_values):
            candidate = self._clean_cell(row_values[city_col])
            if self._looks_like_address(candidate):
                return candidate

        for col in range(org_col + 1, len(row_values)):
            if city_col is not None and col == city_col:
                continue

            candidate = self._clean_cell(row_values[col])
            if self._looks_like_address(candidate):
                return candidate

        return ""

    def _pick_city( self, row_values: list[str], org_col: int, header_city_col: Optional[int], address: str ) -> str:
        if header_city_col is not None and header_city_col < len(row_values):
            candidate = self._clean_cell(row_values[header_city_col])
            if self._is_city_candidate(candidate):
                return candidate

        if address:
            extracted = self._extract_city_from_address(address)
            if extracted:
                return extracted

        fallback: Optional[str] = None
        for col in range(org_col + 1, len(row_values)):
            candidate = self._clean_cell(row_values[col])
            if not self._is_city_candidate(candidate):
                continue

            if len(candidate) <= 40 and not any(ch.isdigit() for ch in candidate):
                return candidate

            if fallback is None:
                fallback = candidate

        return fallback or ""