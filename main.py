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