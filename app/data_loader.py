from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from hashlib import sha1
from pathlib import Path
import re
import threading

import openpyxl

HW_RE = re.compile(r"HW\s*0*(\d+)", re.IGNORECASE)
DATE_RE = re.compile(r"(20\d{2})\.(\d{2})\.(\d{2})(?:\s+(\d{2})\s+(\d{2}))?")


@dataclass(frozen=True)
class HwFile:
    idx: int
    label: str
    path: Path
    date: datetime | None
    mtime: float


def _normalize_header(val: object) -> str:
    if val is None:
        return ""
    text = str(val).strip().lower()
    text = text.replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text)
    return text


def _clean_name(val: object) -> str:
    if val is None:
        return ""
    text = str(val).strip()
    text = text.replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text)
    return text


def _parse_result(val: object) -> tuple[bool | None, str]:
    if val is None:
        return None, ""
    if isinstance(val, bool):
        return val, str(val)
    if isinstance(val, (int, float)):
        if val == 0:
            return False, str(val)
        return True, str(val)
    raw = str(val).strip()
    if not raw:
        return None, ""
    s = raw.lower().replace("\u00a0", " ")
    s = re.sub(r"\s+", " ", s)
    negative = [
        "не зач",
        "незач",
        "не сдан",
        "не принят",
        "нет",
        "fail",
        "0",
    ]
    positive = [
        "зач",
        "прин",
        "сдан",
        "ok",
        "passed",
        "1",
        "yes",
        "да",
    ]
    if any(tok in s for tok in negative):
        return False, raw
    if any(tok in s for tok in positive):
        return True, raw
    return None, raw


def _is_noise_name(name: str) -> bool:
    if not name:
        return True
    lowered = name.lower()
    return lowered in {"итого", "всего"} or lowered.startswith("итого ")


def discover_hw_files(base_dir: Path) -> list[HwFile]:
    candidates: dict[int, HwFile] = {}
    for path in base_dir.glob("*.xlsx"):
        match = HW_RE.search(path.name)
        if not match:
            continue
        idx = int(match.group(1))
        date = None
        date_match = DATE_RE.search(path.stem)
        if date_match:
            y, m, d, hh, mm = date_match.groups()
            if hh is None:
                hh, mm = "00", "00"
            try:
                date = datetime(int(y), int(m), int(d), int(hh), int(mm))
            except ValueError:
                date = None
        hw = HwFile(
            idx=idx,
            label=f"HW{idx:02d}",
            path=path,
            date=date,
            mtime=path.stat().st_mtime,
        )
        prev = candidates.get(idx)
        if prev is None or hw.mtime > prev.mtime:
            candidates[idx] = hw
    return [candidates[idx] for idx in sorted(candidates)]


def _find_header(row: tuple[object, ...]) -> tuple[int | None, int | None]:
    fio_idx = None
    res_idx = None
    for i, cell in enumerate(row):
        norm = _normalize_header(cell)
        if "фио" in norm or "ф.и.о" in norm or "name" in norm:
            fio_idx = i
        if "результ" in norm or "статус" in norm or "оцен" in norm:
            res_idx = i
    return fio_idx, res_idx


def load_hw_results(hw: HwFile) -> dict[str, dict[str, object]]:
    wb = openpyxl.load_workbook(hw.path, data_only=True, read_only=True)

    try:
        ws = wb.active
        fio_idx = None
        res_idx = None
        header_found = False
        results: dict[str, dict[str, object]] = {}
        for row in ws.iter_rows(values_only=True):
            if fio_idx is None or res_idx is None:
                fio_idx, res_idx = _find_header(row)
                if fio_idx is not None and res_idx is not None:
                    header_found = True
                    continue
            if fio_idx is None or res_idx is None:
                continue
            if fio_idx >= len(row):
                continue
            name = _clean_name(row[fio_idx])
            if _is_noise_name(name):
                continue
            raw_val = row[res_idx] if res_idx < len(row) else None
            value, raw = _parse_result(raw_val)
            results[name] = {"value": value, "raw": raw}
        if not header_found:
            raise ValueError(
                "Не найдены колонки ФИО и Результат (или Статус/Оценка) в первой таблице."
            )
        return results
    finally:
        # Ensure file handles are released (especially important for refresh loops).
        wb.close()


def _failed_first_n(per_hw: list[int | None], n: int) -> bool:
    if n <= 0:
        return False
    if not per_hw:
        return True
    limit = min(n, len(per_hw))
    for i in range(limit):
        if per_hw[i] == 1:
            return False
    return True


def _status_for_ratio(accepted: int, total_hw: int, per_hw: list[int | None]) -> str:
    if total_hw <= 0:
        return "низкие показатели"
    if _failed_first_n(per_hw, min(4, total_hw)):
        return "низкие показатели"
    ratio = accepted / total_hw
    if ratio >= (5 / 7):
        return "хорошие показатели"
    if ratio >= (3 / 7):
        return "средние показатели"
    return "низкие показатели"


def _signature(hw_files: list[HwFile]) -> str:
    base = "|".join(f"{hw.path.name}:{int(hw.mtime)}" for hw in hw_files)
    return sha1(base.encode("utf-8")).hexdigest()[:12] if base else "empty"


def build_data(base_dir: Path, hw_files: list[HwFile] | None = None) -> dict[str, object]:
    hw_files = hw_files or discover_hw_files(base_dir)
    generated_at = (
        datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")
    )
    if not hw_files:
        return {
            "meta": {
                "generated_at": generated_at,
                "hw_count": 0,
                "signature": "empty",
                "source_files": [],
                "warnings": ["Не найдено ни одного HW*.xlsx файла."],
            },
            "hws": [],
            "students": [],
        }

    warnings: list[str] = []
    results_by_hw = []
    for hw in hw_files:
        try:
            results_by_hw.append(load_hw_results(hw))
        except Exception as exc:
            warnings.append(f"{hw.label} ({hw.path.name}): {exc}")
            results_by_hw.append({})
    all_names = set()
    for result in results_by_hw:
        all_names.update(result.keys())

    hw_count = len(hw_files)
    students = []
    for name in sorted(all_names):
        per_hw = []
        per_hw_raw = []
        for result in results_by_hw:
            entry = result.get(name)
            if entry is None:
                per_hw.append(None)
                per_hw_raw.append("")
            else:
                val = entry["value"]
                per_hw.append(1 if val is True else 0 if val is False else None)
                per_hw_raw.append(entry["raw"] or "")
        accepted = sum(1 for v in per_hw if v == 1)
        percent = round((accepted / hw_count) * 100, 1) if hw_count else 0.0
        status = _status_for_ratio(accepted, hw_count, per_hw)
        students.append(
            {
                "name": name,
                "accepted": accepted,
                "percent": percent,
                "status": status,
                "group": f"{accepted}/{hw_count}" if hw_count else "0/0",
                "per_hw": per_hw,
                "per_hw_raw": per_hw_raw,
            }
        )

    students.sort(key=lambda s: (-s["accepted"], s["name"].lower()))
    for idx, student in enumerate(students, start=1):
        student["rank"] = idx

    counts_by_accepted: dict[str, int] = {}
    for student in students:
        key = str(student["accepted"])
        counts_by_accepted[key] = counts_by_accepted.get(key, 0) + 1

    hws = []
    for hw in hw_files:
        hws.append(
            {
                "id": hw.idx,
                "label": hw.label,
                "file": hw.path.name,
                "date": hw.date.isoformat(sep=" ") if hw.date else None,
            }
        )

    return {
        "meta": {
            "generated_at": generated_at,
            "hw_count": hw_count,
            "signature": _signature(hw_files),
            "source_files": [hw.path.name for hw in hw_files],
            "warnings": warnings,
        },
        "hws": hws,
        "students": students,
        "stats": {
            "total": len(students),
            "counts_by_accepted": counts_by_accepted,
        },
    }


class DataStore:
    def __init__(self, base_dir: Path) -> None:
        self._base_dir = base_dir
        self._lock = threading.Lock()
        self._cache: dict[str, object] | None = None
        self._signature: str | None = None

    def load(self, force: bool = False) -> dict[str, object]:
        with self._lock:
            hw_files = discover_hw_files(self._base_dir)
            signature = _signature(hw_files)
            if not force and self._cache is not None and signature == self._signature:
                return self._cache
            data = build_data(self._base_dir, hw_files=hw_files)
            self._cache = data
            self._signature = signature
            return data
