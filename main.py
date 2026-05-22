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

    def _extract_city_from_address(self, address: str) -> str:
        address = self._clean_cell(address)
        if not address:
            return ""

        match = re.search(r"\b\d{4,5}\s+([^,]+)", address)
        if match:
            return match.group(1).strip()

        parts = [part.strip() for part in address.split(",") if part.strip()]
        if len(parts) >= 2:
            first = parts[0]
            first = re.sub(r"^\d{4,5}\s+", "", first)
            if first:
                return first
            return parts[1]

        return ""

    def _is_org_candidate(self, value: str) -> bool:
        value = self._clean_cell(value)
        if len(value) < 3:
            return False

        if self._looks_like_noise(value):
            return False

        if self._extract_project_id(value):
            return False

        has_letters = any(ch.isalpha() for ch in value)
        mostly_digits = bool(re.fullmatch(r"[\d\s.,-]+", value))

        return has_letters and not mostly_digits

    def _is_city_candidate(self, value: str) -> bool:
        value = self._clean_cell(value)
        if not value:
            return False

        if self._extract_project_id(value):
            return False

        if self._looks_like_noise(value):
            return False

        if len(value) > 80:
            return False

        has_letters = any(ch.isalpha() for ch in value)
        mostly_digits = bool(re.fullmatch(r"[\d\s.,-]+", value))
        return has_letters and not mostly_digits

    def _looks_like_address(self, value: str) -> bool:
        value = self._clean_cell(value)
        if len(value) < 6:
            return False

        has_digit = any(ch.isdigit() for ch in value)
        has_letter = any(ch.isalpha() for ch in value)

        return has_digit and has_letter

    def _looks_like_noise(self, value: str) -> bool:
        folded = self._normalize_text(value)
        if not folded:
            return True

        return any(marker in folded for marker in self.IGNORE_CELL_MARKERS)

    def _clean_cell(self, value: object) -> str:
        text = str(value) if value is not None else ""
        text = text.replace("\n", " ").replace("\t", " ")
        text = re.sub(r"\s+", " ", text).strip()
        if text.lower() == "nan":
            return ""
        return text

    def _normalize_text(self, value: object) -> str:
        text = self._clean_cell(value).lower()
        decomposed = unicodedata.normalize("NFKD", text)
        without_accents = "".join(ch for ch in decomposed if not unicodedata.combining(ch))
        without_accents = re.sub(r"\s+", " ", without_accents)
        return without_accents

    def _deduplicate(self, records: list[OrganizationRecord]) -> list[OrganizationRecord]:
        deduped: dict[tuple[str, str, str], OrganizationRecord] = {}

        for record in records:
            key = (
                record.project_number,
                self._normalize_text(record.org_name),
                self._normalize_text(record.org_city),
            )
            if key not in deduped:
                deduped[key] = record

        return list(deduped.values())

class OutputExcelWriter:
    def write(self, rows: list[EmailLookupResult], output_path: Path) -> None:
        workbook = xlsxwriter.Workbook(str(output_path))
        worksheet = workbook.add_worksheet("Emails")

        headers = [
            "ProjectNumber",
            "OrgName",
            "Email",
            "OrgCity",
            "Verified",
        ]

        bold = workbook.add_format({"bold": True})
        center = workbook.add_format({"align": "center"})

        for col, header in enumerate(headers):
            worksheet.write(0, col, header, bold)

        worksheet.set_column("A:A", 30)
        worksheet.set_column("B:B", 55)
        worksheet.set_column("C:C", 45)
        worksheet.set_column("D:D", 25)
        worksheet.set_column("E:E", 12)

        for row_idx, row in enumerate(rows, start=1):
            worksheet.write(row_idx, 0, row.project_number)
            worksheet.write(row_idx, 1, row.org_name.replace('"', ""))
            worksheet.write(row_idx, 2, row.email or "")
            worksheet.write(row_idx, 3, row.org_city)
            worksheet.write(row_idx, 4, "YES" if row.verified else "NO", center)

        workbook.close()

class StartupFolderValidator:
    def __init__(self, source_dir: Path, done_dir: Path) -> None:
        self.source_dir = source_dir
        self.done_dir = done_dir

    def validate(self) -> None:
        self._ensure_required_packages()

        if self.source_dir.exists() and not self.source_dir.is_dir():
            raise NotADirectoryError(f"A forras utvonal nem mappa: {self.source_dir}")

        if self.done_dir.exists() and not self.done_dir.is_dir():
            raise NotADirectoryError(f"A done utvonal nem mappa: {self.done_dir}")

        if not self.source_dir.exists():
            self.source_dir.mkdir(parents=True, exist_ok=True)
            print(f"[INFO] Letrehozva: {self.source_dir}")

        if not self.done_dir.exists():
            self.done_dir.mkdir(parents=True, exist_ok=True)
            print(f"[INFO] Letrehozva: {self.done_dir}")

    def ensure_package(self, package_name: str, import_name: str) -> None:
        if importlib.util.find_spec(import_name) is not None:
            return

        print(f"[INFO] Hianyzo csomag, telepites indul: {package_name}")
        try:
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", package_name]
            )
        except subprocess.CalledProcessError as exc:
            raise RuntimeError(
                f"Nem sikerult telepiteni a csomagot: {package_name}"
            ) from exc

class SourseFolderProcessor:
    def __init__( self, folder_path: str = "sources", output_dir: str = "done", max_workers: int = MAX_WORKERS ) -> None:
        self.folder = Path(folder_path)
        self.output_dir = Path(output_dir)
        self.max_workers = max_workers
        self.finder = EmailFinder()
        self.detector = ExcelStructureDetector()
        self.writer = OutputExcelWriter()
        self.startup_validator = StartupFolderValidator(self.folder, self.output_dir)

    def run_project_config(self) -> None:
        self.startup_validator.validate()
        load_runtime_dependencies()
        print("[OK] Project config kesz: csomagok es mappak rendben.")

    def process_all(self) -> None:
        start = time.perf_counter()
        self.run_project_config()

        excel_files = sorted(self.folder.glob("*.xlsx"))
        if not excel_files:
            print(f"Nincs feldolgozhato Excel: {self.folder}")
            return

        print(f"Feldolgozas indul: {len(excel_files)} fajl")

        for index, excel_file in enumerate(excel_files, start=1):
            self._render_progress_header(index, len(excel_files), excel_file.name)
            records = self.detector.extract_records(excel_file)
            if not records:
                print(f"[SKIP] {excel_file.name}: nincs feldolgozhato sor")
                continue

            print(f"[INFO] {excel_file.name}: {len(records)} rekord")
            file_results = self._process_records(records)

            output_name = f"{excel_file.stem}_emails_output.xlsx"
            output_path = self.output_dir / output_name
            self.writer.write(file_results, output_path)
            print(f"[OK] Elkeszult: {output_path}")

        elapsed = time.perf_counter() - start
        print(f"Befejezve: {self._format_duration(elapsed)}")

    def _process_records(self, records: list[OrganizationRecord]) -> list[EmailLookupResult]:
        output_rows: list[EmailLookupResult] = []
        total_records = len(records)
        processed = 0
        found_count = 0
        verified_count = 0

        self._render_record_progress(
            processed=processed,
            total=total_records,
            found_count=found_count,
            verified_count=verified_count,
        )

        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            future_map = {
                executor.submit(
                    self.finder.find_email,
                    record.org_name,
                    self._extract_country_code(record.project_number),
                ): record
                for record in records
            }

            for future in as_completed(future_map):
                record = future_map[future]
                try:
                    email, verified = future.result()
                except Exception as exc:
                    print("")
                    print(f"[WARN] Email hiba: {record.org_name} ({exc})")
                    email, verified = None, False

                output_email = email if verified else None
                processed += 1
                if email:
                    found_count += 1
                if verified:
                    verified_count += 1

                output = EmailLookupResult(
                    source_file=record.source_file,
                    sheet_name=record.sheet_name,
                    project_number=record.project_number,
                    org_name=record.org_name,
                    org_city=record.org_city,
                    email=output_email,
                    verified=verified,
                )
                output_rows.append(output)
                self._render_record_progress(
                    processed=processed,
                    total=total_records,
                    found_count=found_count,
                    verified_count=verified_count,
                )

        print("")

        return output_rows

    def _extract_country_code(self, project_number: str) -> Optional[str]:
        compact = re.sub(r"\s+", "", str(project_number).upper())
        match = re.search(r"^\d{4}-\d-([A-Z]{2})\d{2}-KA1\d{2}-[A-Z]{3}-\d{5,}$", compact)
        if not match:
            return None
        return match.group(1)

    def _format_duration(self, seconds: float) -> str:
        whole_seconds = int(seconds)
        hours = whole_seconds // 3600
        minutes = (whole_seconds % 3600) // 60
        secs = whole_seconds % 60

        if hours > 0:
            return f"{hours}h {minutes}m {secs}s"
        if minutes > 0:
            return f"{minutes}m {secs}s"
        return f"{secs}s"

    def _render_progress_header(self, current: int, total: int, file_name: str) -> None:
        clear_console()
        completed = max(0, current - 1)
        percent = int((completed / total) * 100) if total > 0 else 0
        progress_bar = self._build_progress_bar(percent)
        header = f"Progress: {progress_bar} {percent}% | {current}/{total} | {file_name}"
        print(header)
        print("")

    def _build_progress_bar(self, percent: int, width: int = 24) -> str:
        safe_percent = max(0, min(100, percent))
        filled = int((safe_percent / 100) * width)
        return "[" + ("#" * filled) + ("-" * (width - filled)) + "]"

    def _render_record_progress( self, processed: int, total: int, found_count: int, verified_count: int ) -> None:
        percent = int((processed / total) * 100) if total > 0 else 100
        bar = self._build_progress_bar(percent, width=18)
        line = (
            f"Records: {bar} {percent}% | {processed}/{total} | "
            f"Found: {found_count} | Verified: {verified_count}"
        )
        print(f"\r{line.ljust(120)}", end="", flush=True)

class ConsoleMenuApp:
    def __init__(self, processor: SourseFolderProcessor) -> None:
        self.processor = processor
        self.readme_path = Path("README.md")

    def run(self) -> None:
        self.processor.startup_validator.ensure_package("questionary", "questionary")
        load_runtime_dependencies()

        while True:
            clear_console()
            choice = questionary.select(
                "Valassz egy lehetoseget:",
                choices=[
                    "Project Config",
                    "Show README",
                    "Start Processing",
                    "Exit",
                ],
            ).ask()

            if choice == "Project Config":
                self.processor.run_project_config()
                self.wait_for_continue()
                continue

            if choice == "Show README":
                self.show_readme()
                continue

            if choice == "Start Processing":
                confirmed = questionary.confirm(
                    "Minden keszen all, induljon a feldolgozas?",
                    default=True,
                ).ask()

                if confirmed:
                    self.processor.process_all()
                else:
                    print("[INFO] A feldolgozas megszakitva.")
                continue

            print("Kilepes.")
            break

    def show_readme(self) -> None:
        if not self.readme_path.exists():
            print(f"[WARN] Nem talalhato README: {self.readme_path}")
            input("Nyomj Entert a visszalepeshez...")
            return

        content = self.readme_path.read_text(encoding="utf-8")
        print("\n" + "=" * 60)
        print("README.md")
        print("=" * 60)
        print(content)
        print("=" * 60 + "\n")
        self.wait_for_continue()