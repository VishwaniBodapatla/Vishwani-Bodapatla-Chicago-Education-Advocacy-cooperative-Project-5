"""
Wrapper around pipelinescrapper_mod.py (UNCHANGED).

After the original Excel is generated:
- Split 'branch' by commas
- Ignore any comma-part that has more than 5 words
- Load canonical majors from: canonical_majors_custom_full.xlsx (sheet: canonical_list_only)
- Map each branch term to canonical majors using:
    * custom alias rules (info systems vs info sciences vs IT etc.)
    * exact canonical match
    * fuzzy match (difflib)
- ✅ Only keep terms that successfully map to the custom canonical list
- Create/extend one column per kept canonical branch
- ✅ Mark the row under those columns with the branch FIELD NAME (not 1)
- Save back IN-PLACE to the SAME Excel file

NEW (ROLE CANONICAL):
- Map 'role_name' into popular canonical role columns (like branch canonical)
- Supports wildcard patterns like *security analyst* (anything before/after matches)
- Creates/updates one column per canonical role and writes the role bucket name into the cell
- Drops old role-canonical columns (based on current ROLE_PATTERNS keys) before rebuilding

NEW (SENIORITY CANONICAL):
- Map 'seniority' into popular canonical seniority columns
- Supports wildcard patterns like *senior*, *intern*, *manager*, etc.
- Creates/updates one column per canonical seniority and writes the bucket name into the cell
- Drops old seniority-canonical columns (based on current SENIORITY_PATTERNS keys) before rebuilding
"""

from __future__ import annotations

import re
from difflib import get_close_matches
from pathlib import Path
from typing import Optional, Callable, Any, List, Tuple

import pandas as pd

from pipelinescrapper_mod import run_pipeline as _run_pipeline_original, OUT_EXCEL as _OUT_EXCEL

# -----------------------------
# Base (non-canonical) columns
# -----------------------------
# We keep ONLY these columns from jobright_jobs.xlsx before rebuilding canonical branch columns.
# This prevents old canonical columns from previous runs from sticking around after you edit
# canonical_majors_custom_full.xlsx.
BASE_HEADERS = [
    "platform",
    "role_selected",
    "role_name",
    "role_type",
    "work_model",
    "company_name",
    "company_url",
    "job_url",
    "salary",
    "salary_min",
    "salary_max",
    "degree_level",
    "branch",
    "location",
    "posted",
    "seniority",
]

# -----------------------------
# Custom canonical (XLSX)
# -----------------------------
BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent
INPUT_DIR = PROJECT_DIR / "input"
CUSTOM_CANONICAL_XLSX = INPUT_DIR / "canonical_majors_custom_full.xlsx"
CUSTOM_CANONICAL_SHEET = "canonical_list_only"   # sheet created in your new excel
CUSTOM_CANONICAL_COL = "canonical"               # column name in that sheet

# -----------------------------
# Your key buckets (BRANCH)
# -----------------------------
INFO_SCI = "information sciences"
INFO_SYS = "information systems"
INFO_TECH = "information technology"
INFO_CSYS = "computer information systems"

# -----------------------------
# Alias rules (BRANCH) (edit anytime)
# -----------------------------
# These map raw terms into the exact canonical bucket names above.
_INFO_SCI_ALIASES = [
    "information science",
    "information sciences",
]

_INFO_SYS_ALIASES = [
    "information system",
    "information systems",
    "cis",
    "mis",
    "management information systems",
    "management information systems and statistics",
]

_INFO_TECH_ALIASES = [
    "information technology",
    "it",           # special case: only when term is exactly "it"
    "informatics",  # you can move this to INFO_SYS if you prefer
]

_INFO_CSYS_ALIASES = [
    "computer information systems",
    "computer information system",
    "computer & information systems",
    "computer and information systems",
    "computer info systems",
]

# -----------------------------
# ROLE canonical buckets (wildcards supported)
# -----------------------------
# Pattern rules use "*" to mean "anything before/after is allowed".
# Example: "*security analyst*" matches:
#   "Senior Security Analyst", "Security Analyst II", "Lead Security Analyst - Cloud"
ROLE_PATTERNS: dict[str, List[str]] = {
    "data analyst": ["*data analyst*", "*analytics analyst*", "*reporting analyst*"],
    "data scientist": ["*data scientist*"],
    "security analyst": ["*security analyst*", "*soc analyst*", "*information security analyst*", "*infosec analyst*"],
    "cybersecurity analyst": ["*cybersecurity analyst*", "*cyber analyst*"],
    "security engineer": ["*security engineer*", "*application security engineer*", "*appsec engineer*"],
    "software engineer": ["*software engineer*", "*backend engineer*", "*frontend engineer*"],
    "software developer": ["*software developer*", "*developer*"],
    "machine learning engineer": ["*machine learning engineer*", "*ml engineer*"],
    "devops engineer": ["*devops engineer*", "*site reliability engineer*", "*sre*"],
    "cloud engineer": ["*cloud engineer*", "*aws engineer*", "*azure engineer*", "*gcp engineer*"],
    "business analyst": ["*business analyst*"],
    "product manager": ["*product manager*", "*product owner*"],
}

# Extra token-only abbreviations (to avoid false positives like "company" matching "pm")
# These are checked with word boundaries.
ROLE_TOKEN_ALIASES: dict[str, List[str]] = {
    "software engineer": ["swe"],
    "devops engineer": ["sre"],
    "data scientist": ["ds"],
    "product manager": ["pm"],
}

# -----------------------------
# SENIORITY canonical buckets (wildcards supported)
# -----------------------------
SENIORITY_PATTERNS: dict[str, List[str]] = {
    "intern": ["*intern*", "*internship*", "*co-op*", "*coop*"],
    "entry": ["*entry*", "*junior*", "*jr*", "*associate*", "*new grad*", "*graduate*"],
    "mid": ["*mid*", "*intermediate*", "*mid-level*", "*mid level*"],
    "senior": ["*senior*", "*sr*", "*principal*", "*staff*"],
    "lead": ["*lead*", "*tech lead*", "*team lead*"],
    "manager": ["*manager*", "*mgr*", "*management*"],
    "director": ["*director*", "*head of*"],
    "vp": ["*vp*", "*vice president*"],
    "c-level": ["*cto*", "*ciso*", "*ceo*", "*cfo*", "*coo*", "*chief*"],
}

SENIORITY_TOKEN_ALIASES: dict[str, List[str]] = {
    "entry": ["jr"],
    "senior": ["sr"],
    "vp": ["vp"],
}

# -----------------------------
# Common utils
# -----------------------------
def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip().lower())


def _word_count(s: str) -> int:
    return len((s or "").split())


def _split_branch_cell(cell: Any, max_words: int = 5) -> List[str]:
    """Split by comma. Ignore segments with > max_words."""
    if cell is None:
        return []
    s = str(cell).strip()
    if not s or s.lower() in {"nan", "none"}:
        return []
    parts = [p.strip() for p in s.split(",")]
    out: List[str] = []
    seen = set()
    for p in parts:
        v = _norm(p)
        if not v:
            continue
        if _word_count(v) > max_words:
            continue
        if v not in seen:
            out.append(v)
            seen.add(v)
    return out


def _contains_alias(term: str, aliases: List[str]) -> bool:
    t = _norm(term)
    for a in aliases:
        a2 = _norm(a)
        if a2 == "it":
            continue
        if a2 and a2 in t:
            return True
    return False


def _alias_bucket(term: str) -> str | None:
    """
    Returns a canonical bucket name (must exist in canonical set) or None.
    """
    t = _norm(term)

    # "it" only when it is exactly "it"
    if re.fullmatch(r"it", t):
        return INFO_TECH

    # ✅ Computer Information Systems bucket support
    if _contains_alias(t, _INFO_CSYS_ALIASES):
        return INFO_CSYS

    if _contains_alias(t, _INFO_SCI_ALIASES):
        return INFO_SCI

    if _contains_alias(t, _INFO_SYS_ALIASES) or re.search(r"\b(cis|mis)\b", t):
        return INFO_SYS

    if _contains_alias(t, _INFO_TECH_ALIASES):
        return INFO_TECH

    return None


def _load_custom_canonical_xlsx(
    path: Path = CUSTOM_CANONICAL_XLSX,
    sheet: str = CUSTOM_CANONICAL_SHEET,
    col: str = CUSTOM_CANONICAL_COL,
) -> Tuple[List[str], set[str]]:
    """
    Loads canonical list from your XLSX.
    Returns (canonical_list, canonical_set) all lowercased.
    """
    if not path.exists():
        return [], set()

    try:
        df = pd.read_excel(path, sheet_name=sheet)
    except Exception:
        return [], set()

    if col not in df.columns:
        # try case-insensitive match
        cols_l = {c.lower(): c for c in df.columns}
        real = cols_l.get(col.lower())
        if not real:
            return [], set()
        col = real

    vals = []
    for v in df[col].dropna().astype(str).tolist():
        v2 = _norm(v)
        if v2:
            vals.append(v2)

    # unique preserve order
    vals = list(dict.fromkeys(vals))
    return vals, set(vals)


def _map_to_canonical_only(
    raw_terms: List[str],
    canonical_list: List[str],
    canonical_set: set[str],
    fuzzy_cutoff: float = 0.70,
) -> List[str]:
    """
    Map raw terms to your canonical list ONLY.
    Priority:
    1) alias buckets (information systems vs sciences vs IT vs CIS)
    2) exact match
    3) fuzzy match
    """
    out: List[str] = []
    seen = set()

    for t0 in raw_terms:
        t = _norm(t0)
        if not t:
            continue

        # 1) alias bucket
        bucket = _alias_bucket(t)
        if bucket and bucket in canonical_set:
            if bucket not in seen:
                out.append(bucket)
                seen.add(bucket)
            continue

        # 2) exact match
        if t in canonical_set:
            if t not in seen:
                out.append(t)
                seen.add(t)
            continue

        # 3) fuzzy match
        match = get_close_matches(t, canonical_list, n=1, cutoff=fuzzy_cutoff)
        if match:
            m = match[0]
            if m in canonical_set and m not in seen:
                out.append(m)
                seen.add(m)

    return out


def branch_preprocess_inplace_custom_xlsx(
    excel_path: Path,
    sheet_name: str = "jobs",
    branch_col: str = "branch",
    marker: Any = 1,  # kept for compatibility, not used now
    max_words: int = 5,
    fuzzy_cutoff: float = 0.70,
) -> None:
    """
    IN-PLACE update of excel_path with canonical branch columns from your XLSX.
    ✅ Writes the branch name itself into the canonical columns (instead of 1).
    """
    if not excel_path.exists():
        return

    canonical_list, canonical_set = _load_custom_canonical_xlsx()
    if not canonical_list:
        # no canonical list, do nothing
        return

    # Ensure key buckets exist even if user forgot to add them
    for must in (INFO_SCI, INFO_SYS, INFO_TECH, INFO_CSYS):
        if must not in canonical_set:
            canonical_list.append(must)
            canonical_set.add(must)

    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    # ✅ IMPORTANT: Drop old canonical columns so canonical-sheet edits actually reflect.
    # Keep only the base headers that exist in the sheet.
    keep_cols = [c for c in BASE_HEADERS if c in df.columns]
    if keep_cols:
        df = df[keep_cols].copy()

    if branch_col not in df.columns:
        return

    per_row: List[List[str]] = []
    all_branches = set()

    for cell in df[branch_col].tolist():
        raw = _split_branch_cell(cell, max_words=max_words)
        mapped = _map_to_canonical_only(
            raw,
            canonical_list,
            canonical_set,
            fuzzy_cutoff=fuzzy_cutoff,
        )
        per_row.append(mapped)
        all_branches.update(mapped)

    if not all_branches:
        return

    # Create missing canonical columns
    for b in sorted(all_branches):
        if b not in df.columns:
            df[b] = ""

    # ✅ Fill with branch FIELD NAME
    for i, branches in enumerate(per_row):
        for b in branches:
            df.at[i, b] = b

    # Write back in-place
    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)


# -----------------------------
# ROLE canonical helpers
# -----------------------------
def _pattern_core(pat: str) -> str:
    """Remove '*' wildcards and normalize."""
    return _norm(pat.replace("*", " "))


def _contains_token(txt: str, token: str) -> bool:
    """True if token appears as a whole word in txt."""
    token = _norm(token)
    if not token:
        return False
    return re.search(rf"\b{re.escape(token)}\b", txt) is not None


def _map_role_to_canonical(role_text: Any) -> List[str]:
    """
    Map role_name -> canonical role buckets using wildcard patterns.
    Example: "*security analyst*" matches anything containing "security analyst".
    Also supports strict token aliases (pm, swe, ds, sre) using word boundaries.
    """
    if role_text is None:
        return []

    txt = _norm(str(role_text))
    if not txt or txt in {"nan", "none"}:
        return []

    found: List[str] = []
    seen = set()

    # 1) token aliases (word-boundary)
    for canonical, toks in ROLE_TOKEN_ALIASES.items():
        for t in toks:
            if _contains_token(txt, t):
                if canonical not in seen:
                    found.append(canonical)
                    seen.add(canonical)

    # 2) wildcard patterns (contains)
    for canonical, patterns in ROLE_PATTERNS.items():
        for p in patterns:
            core = _pattern_core(p)
            if not core:
                continue

            # If core is short, use word boundaries to reduce false positives
            if len(core) <= 3:
                hit = _contains_token(txt, core)
            else:
                hit = core in txt

            if hit:
                if canonical not in seen:
                    found.append(canonical)
                    seen.add(canonical)

    # 3) fallback fuzzy match (against canonical keys only)
    if not found:
        match = get_close_matches(txt, list(ROLE_PATTERNS.keys()), n=1, cutoff=0.65)
        if match:
            found.append(match[0])

    return found


def role_preprocess_inplace(
    excel_path: Path,
    sheet_name: str = "jobs",
    role_col: str = "role_name",
) -> None:
    """
    IN-PLACE update of excel_path with canonical ROLE columns.
    ✅ Supports wildcard patterns like *security analyst*.
    ✅ Drops old role-canonical columns (based on ROLE_PATTERNS keys) before rebuilding.
    ✅ Writes the role bucket name into the canonical column (not 1).
    """
    if not excel_path.exists():
        return

    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    if role_col not in df.columns:
        return

    # Drop existing role-canonical columns so edits to ROLE_PATTERNS reflect immediately
    role_cols_existing = [c for c in ROLE_PATTERNS.keys() if c in df.columns]
    if role_cols_existing:
        df = df.drop(columns=role_cols_existing)

    per_row: List[List[str]] = []
    all_roles = set()

    for cell in df[role_col].tolist():
        mapped = _map_role_to_canonical(cell)
        per_row.append(mapped)
        all_roles.update(mapped)

    if not all_roles:
        return

    # Create missing role columns
    for r in sorted(all_roles):
        if r not in df.columns:
            df[r] = ""

    # Fill rows
    for i, roles in enumerate(per_row):
        for r in roles:
            df.at[i, r] = r

    # Write back in-place
    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)


# -----------------------------
# SENIORITY canonical helpers
# -----------------------------
def _map_seniority_to_canonical(sen_text: Any) -> List[str]:
    """
    Map seniority -> canonical seniority buckets.
    Supports wildcard patterns like *senior* and token aliases like sr/jr/vp.
    Also supports comma-separated values (if the cell has multiple labels).
    """
    if sen_text is None:
        return []

    txt_raw = str(sen_text).strip()
    if not txt_raw or txt_raw.lower() in {"nan", "none"}:
        return []

    # Allow comma-separated
    parts = [p.strip() for p in txt_raw.split(",") if p.strip()]
    if not parts:
        parts = [txt_raw]

    found: List[str] = []
    seen = set()

    for part in parts:
        txt = _norm(part)
        if not txt:
            continue

        # 1) token aliases (word-boundary)
        for canonical, toks in SENIORITY_TOKEN_ALIASES.items():
            for t in toks:
                if _contains_token(txt, t):
                    if canonical not in seen:
                        found.append(canonical)
                        seen.add(canonical)

        # 2) wildcard patterns (contains)
        for canonical, patterns in SENIORITY_PATTERNS.items():
            for p in patterns:
                core = _pattern_core(p)
                if not core:
                    continue

                # If core is short, use word boundaries to reduce false positives
                if len(core) <= 3:
                    hit = _contains_token(txt, core)
                else:
                    hit = core in txt

                if hit:
                    if canonical not in seen:
                        found.append(canonical)
                        seen.add(canonical)

    # 3) fallback fuzzy match (against keys only)
    if not found:
        match = get_close_matches(_norm(txt_raw), list(SENIORITY_PATTERNS.keys()), n=1, cutoff=0.65)
        if match:
            found.append(match[0])

    return found


def seniority_preprocess_inplace(
    excel_path: Path,
    sheet_name: str = "jobs",
    seniority_col: str = "seniority",
) -> None:
    """
    IN-PLACE update of excel_path with canonical SENIORITY columns.
    ✅ Supports wildcard patterns like *senior*.
    ✅ Drops old seniority-canonical columns (based on SENIORITY_PATTERNS keys) before rebuilding.
    ✅ Writes the seniority bucket name into the canonical column (not 1).
    """
    if not excel_path.exists():
        return

    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    if seniority_col not in df.columns:
        return

    # Drop existing seniority-canonical columns so edits to patterns reflect immediately
    sen_cols_existing = [c for c in SENIORITY_PATTERNS.keys() if c in df.columns]
    if sen_cols_existing:
        df = df.drop(columns=sen_cols_existing)

    per_row: List[List[str]] = []
    all_sen = set()

    for cell in df[seniority_col].tolist():
        mapped = _map_seniority_to_canonical(cell)
        per_row.append(mapped)
        all_sen.update(mapped)

    if not all_sen:
        return

    # Create missing columns
    for s in sorted(all_sen):
        if s not in df.columns:
            df[s] = ""

    # Fill rows
    for i, sels in enumerate(per_row):
        for s in sels:
            df.at[i, s] = s

    # Write back in-place
    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)


def run_pipeline(
    target_jobs: int,
    headless: bool = False,
    wait_for_login: Optional[Callable[[], None]] = None,
    on_status: Optional[Callable[[str], None]] = None,
) -> None:
    """
    Same signature as original run_pipeline, but custom canonical postprocess after Excel generation.
    """
    _run_pipeline_original(
        target_jobs=target_jobs,
        headless=headless,
        wait_for_login=wait_for_login,
        on_status=on_status,
    )

    try:
        if on_status:
            on_status(
                "[BRANCH] Custom XLSX canonical postprocess: splitting 'branch' into canonical columns (in-place)..."
            )

        branch_preprocess_inplace_custom_xlsx(
            _OUT_EXCEL,
            sheet_name="jobs",
            branch_col="branch",
            marker=1,
            max_words=5,
            fuzzy_cutoff=0.70,
        )

        if on_status:
            on_status(f"[BRANCH] Done. Updated Excel: {_OUT_EXCEL.resolve()}")

        if on_status:
            on_status("[ROLE] Canonical role postprocess: splitting 'role_name' into popular role columns (in-place)...")

        role_preprocess_inplace(
            _OUT_EXCEL,
            sheet_name="jobs",
            role_col="role_name",
        )

        if on_status:
            on_status(f"[ROLE] Done. Updated Excel: {_OUT_EXCEL.resolve()}")

        if on_status:
            on_status("[SENIORITY] Canonical seniority postprocess: splitting 'seniority' into popular seniority columns (in-place)...")

        seniority_preprocess_inplace(
            _OUT_EXCEL,
            sheet_name="jobs",
            seniority_col="seniority",
        )

        if on_status:
            on_status(f"[SENIORITY] Done. Updated Excel: {_OUT_EXCEL.resolve()}")

    except Exception as e:
        if on_status:
            on_status(f"[POSTPROCESS][WARN] Postprocess failed: {e}")