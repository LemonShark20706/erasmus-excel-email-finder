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

