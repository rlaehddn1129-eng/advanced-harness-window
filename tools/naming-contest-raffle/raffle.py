#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""네이밍 공모전 참가상 추첨 프로그램 (1인 1표 · 검증 가능 추첨).

핵심 원칙
  1) 다량 제출로 당첨 확률을 올릴 수 없다.
     - 응모 '건' 단위가 아니라 '사람' 단위로 추첨한다. 10건을 내도 응모권은 1장.
  2) 동명이인은 다른 사람으로 취급한다.
     - 이름이 아니라 이메일/휴대폰번호(정규화 후)로 사람을 식별하고 병합한다.
  3) 최상위 1건만 시상한다.
     - 최우수/우수 수상자는 사람 단위로 제외되므로, 그 사람의 다른 제출건도 함께 빠진다.
  4) 추첨 결과는 누구나 재현·검증할 수 있다.
     - 공개 시드(seed)와 HMAC-SHA256 기반 결정적 추첨. 명부 해시로 사후 변조도 확인 가능.

의존성 없음(파이썬 3.9+ 표준 라이브러리만). 이 파일 하나만 복사해도 동작한다.

사용 예시
  # 1) 추첨 전 응모 현황 점검(중복 제출·이상 응모 확인)
  python3 raffle.py inspect --input 응모현황.xlsx

  # 2) 참가상 10명 + 예비 5명 추첨
  python3 raffle.py draw --input 응모현황.xlsx \
      --winners 10 --reserve 5 \
      --seed "2026-09-15 추첨 입회 하 공개시드" \
      --exclude-winner "hong@example.com" \
      --exclude-winner "010-1234-5678" \
      --out 결과

  # 3) 결과 재현 검증(제3자가 동일 명부·동일 시드로 확인)
  python3 raffle.py verify --input 응모현황.xlsx --audit 결과/audit.json
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import hmac
import json
import os
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

VERSION = "1.0.0"
KST = timezone(timedelta(hours=9))


# --------------------------------------------------------------------------
# 표 읽기 (CSV / TSV / XLSX)
# --------------------------------------------------------------------------

def _read_csv(path: str) -> List[Dict[str, str]]:
    delim = "\t" if path.lower().endswith((".tsv", ".tab")) else ","
    # 엑셀에서 저장한 한글 CSV는 utf-8-sig 또는 cp949인 경우가 많다.
    for enc in ("utf-8-sig", "utf-8", "cp949", "euc-kr"):
        try:
            with open(path, "r", encoding=enc, newline="") as fp:
                rows = list(csv.DictReader(fp, delimiter=delim))
            return [{(k or "").strip(): (v or "") for k, v in row.items()} for row in rows]
        except UnicodeDecodeError:
            continue
    raise SystemExit(f"[오류] 인코딩을 판별할 수 없습니다: {path}")


def _xlsx_serial_to_text(value: str) -> str:
    """엑셀 날짜 일련번호를 사람이 읽는 문자열로 바꾼다(정수/실수 모두 처리)."""
    try:
        num = float(value)
    except ValueError:
        return value
    if not (20000 <= num <= 80000):  # 날짜로 보기 어려운 값은 그대로 둔다.
        return value
    base = datetime(1899, 12, 30)
    dt = base + timedelta(days=num)
    return dt.strftime("%Y-%m-%d %H:%M:%S")


def _read_xlsx(path: str) -> List[Dict[str, str]]:
    """openpyxl 없이 xlsx 첫 시트를 읽는다(단순 표 전용)."""
    import zipfile
    import xml.etree.ElementTree as ET

    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    rns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())

        shared: List[str] = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in root.findall(f"{ns}si"):
                shared.append("".join(t.text or "" for t in si.iter(f"{ns}t")))

        sheet_path = "xl/worksheets/sheet1.xml"
        if "xl/workbook.xml" in names and "xl/_rels/workbook.xml.rels" in names:
            wb = ET.fromstring(zf.read("xl/workbook.xml"))
            rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
            rel_map = {r.get("Id"): r.get("Target") for r in rels}
            first = wb.find(f"{ns}sheets/{ns}sheet")
            if first is not None:
                target = rel_map.get(first.get(f"{rns}id"))
                if target:
                    target = target.lstrip("/")
                    cand = target if target.startswith("xl/") else f"xl/{target}"
                    if cand in names:
                        sheet_path = cand
        if sheet_path not in names:
            raise SystemExit(f"[오류] 시트를 찾을 수 없습니다: {path}")

        sheet = ET.fromstring(zf.read(sheet_path))

    def col_index(ref: str) -> int:
        letters = "".join(ch for ch in ref if ch.isalpha())
        idx = 0
        for ch in letters:
            idx = idx * 26 + (ord(ch.upper()) - 64)
        return idx - 1

    grid: List[List[str]] = []
    for row in sheet.iter(f"{ns}row"):
        cells: Dict[int, str] = {}
        for c in row.findall(f"{ns}c"):
            ctype = c.get("t")
            if ctype == "inlineStr":
                text = "".join(t.text or "" for t in c.iter(f"{ns}t"))
            else:
                v = c.find(f"{ns}v")
                text = v.text or "" if v is not None else ""
                if ctype == "s" and text.isdigit() and int(text) < len(shared):
                    text = shared[int(text)]
                elif ctype is None and c.get("s") is not None:
                    text = _xlsx_serial_to_text(text)
            if text:
                cells[col_index(c.get("r") or "A1")] = text
        width = (max(cells) + 1) if cells else 0
        grid.append([cells.get(i, "") for i in range(width)])

    while grid and not any(x.strip() for x in grid[0]):
        grid.pop(0)
    if not grid:
        return []

    header = [h.strip() for h in grid[0]]
    out: List[Dict[str, str]] = []
    for row in grid[1:]:
        if not any(x.strip() for x in row):
            continue
        out.append({header[i]: (row[i] if i < len(row) else "") for i in range(len(header))})
    return out


def read_table(path: str) -> List[Dict[str, str]]:
    if not os.path.exists(path):
        raise SystemExit(f"[오류] 파일이 없습니다: {path}")
    if path.lower().endswith((".xlsx", ".xlsm")):
        return _read_xlsx(path)
    if path.lower().endswith(".xls"):
        raise SystemExit("[오류] 구형 .xls는 지원하지 않습니다. .xlsx 또는 CSV로 저장해 주세요.")
    return _read_csv(path)


# --------------------------------------------------------------------------
# 정규화 (동일인 판별의 근거)
# --------------------------------------------------------------------------

def norm_text(value: Any) -> str:
    s = unicodedata.normalize("NFKC", str(value or ""))
    return re.sub(r"\s+", " ", s).strip()


def display_text(value: Any) -> str:
    """표시·출력용. 앞뒤 공백만 정리하고 원문 글자는 그대로 둔다(ㅋㅋㅋ가 ᄏᄏᄏ로 보이지 않도록)."""
    return re.sub(r"[ \t\r\n]+", " ", str(value or "")).strip()


def norm_name(value: Any) -> str:
    """이름은 공백을 모두 제거해 '홍 길동'과 '홍길동'을 같게 본다(식별 키로는 쓰지 않음)."""
    return re.sub(r"\s+", "", norm_text(value))


def norm_phone(value: Any) -> str:
    """휴대폰번호를 숫자만 남긴 표준형으로. 국가번호 +82 → 0으로 환원."""
    s = norm_text(value)
    if not s:
        return ""
    s = re.sub(r"[^\d+]", "", s)
    if s.startswith("+82"):
        s = "0" + s[3:]
    elif s.startswith("0082"):
        s = "0" + s[4:]
    elif s.startswith("82") and len(s) >= 11:
        s = "0" + s[2:]
    s = re.sub(r"\D", "", s)
    if s and not s.startswith("0") and len(s) in (9, 10):
        s = "0" + s
    return s if 9 <= len(s) <= 11 else ""


def norm_email(value: Any, strict: bool = True) -> str:
    """이메일 표준화.

    strict=True(기본)이면 별칭으로 계정을 늘려 응모권을 늘리는 것을 막는다.
      - 소문자화
      - '+태그' 제거      (hong+1@gmail.com → hong@gmail.com)
      - gmail 계열은 점 제거 (h.o.ng@gmail.com → hong@gmail.com)
    """
    s = norm_text(value).lower().replace(" ", "")
    if s.count("@") != 1:
        return ""
    local, domain = s.split("@")
    if not local or "." not in domain:
        return ""
    if strict:
        local = local.split("+", 1)[0]
        if domain in ("gmail.com", "googlemail.com"):
            local = local.replace(".", "")
            domain = "gmail.com"
    return f"{local}@{domain}" if local else ""


def parse_dt(value: Any) -> Optional[datetime]:
    s = norm_text(value)
    if not s:
        return None
    s = s.replace("오전", "AM").replace("오후", "PM")
    fmts = (
        "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
        "%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y/%m/%d",
        "%Y.%m.%d %H:%M:%S", "%Y.%m.%d %H:%M", "%Y.%m.%d",
        "%Y-%m-%dT%H:%M:%S", "%m/%d/%Y %H:%M:%S", "%m/%d/%Y",
    )
    for fmt in fmts:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


# --------------------------------------------------------------------------
# 열(컬럼) 자동 인식
# --------------------------------------------------------------------------

FIELD_KEYWORDS: Dict[str, Sequence[str]] = {
    # 우선순위 순서대로 검사한다. '네이밍'이 '이름'보다 먼저 와야 한다.
    "naming": ("네이밍", "제안명", "응모작", "작품명", "제출명", "아이디어", "제안이름", "naming", "proposal"),
    "submitted_at": ("제출일시", "접수일시", "응모일시", "타임스탬프", "timestamp", "제출시간", "등록일"),
    "email": ("이메일", "메일", "email", "e-mail", "mail"),
    "phone": ("연락처", "휴대폰", "핸드폰", "전화", "폰번호", "phone", "mobile", "tel", "hp"),
    "org": ("소속", "기관", "회사", "학교", "부서", "단체", "org", "company"),
    "extra": ("생년월일", "생일", "사번", "학번", "식별", "birth", "추가정보"),
    "name": ("성명", "이름", "응모자", "참가자", "name"),
}


def detect_columns(header: Sequence[str], overrides: Dict[str, str]) -> Dict[str, Optional[str]]:
    cols: Dict[str, Optional[str]] = {k: None for k in FIELD_KEYWORDS}
    used: set = set()

    for field_name, col in overrides.items():
        if field_name not in cols:
            raise SystemExit(f"[오류] 알 수 없는 열 지정: {field_name}")
        if col not in header:
            raise SystemExit(f"[오류] '{col}' 열이 파일에 없습니다. 실제 열: {list(header)}")
        cols[field_name] = col
        used.add(col)

    norm_header = {h: norm_text(h).lower().replace(" ", "") for h in header}
    for field_name, keywords in FIELD_KEYWORDS.items():
        if cols[field_name]:
            continue
        for kw in keywords:
            for h in header:
                if h in used:
                    continue
                if kw in norm_header[h]:
                    cols[field_name] = h
                    used.add(h)
                    break
            if cols[field_name]:
                break
    return cols


# --------------------------------------------------------------------------
# 응모 건 / 사람
# --------------------------------------------------------------------------

@dataclass
class Entry:
    row_no: int
    naming: str = ""
    name: str = ""
    email: str = ""
    phone: str = ""
    org: str = ""
    extra: str = ""
    submitted_at: Optional[datetime] = None
    raw: Dict[str, str] = field(default_factory=dict)
    issues: List[str] = field(default_factory=list)

    @property
    def valid(self) -> bool:
        return not self.issues

    @property
    def sort_key(self) -> Tuple[int, float, int]:
        if self.submitted_at:
            return (0, self.submitted_at.timestamp(), self.row_no)
        return (1, 0.0, self.row_no)


@dataclass
class Person:
    pid: int
    entries: List[Entry] = field(default_factory=list)
    emails: List[str] = field(default_factory=list)
    phones: List[str] = field(default_factory=list)

    @property
    def primary(self) -> Entry:
        return min(self.entries, key=lambda e: e.sort_key)

    @property
    def name(self) -> str:
        return self.primary.name

    @property
    def org(self) -> str:
        return self.primary.org

    @property
    def submissions(self) -> int:
        return len(self.entries)

    @property
    def names(self) -> List[str]:
        seen: List[str] = []
        for e in sorted(self.entries, key=lambda x: x.sort_key):
            if e.name and e.name not in seen:
                seen.append(e.name)
        return seen

    @property
    def identity_key(self) -> str:
        """추첨 해시의 입력값. 사람이 가진 모든 식별자를 정렬해 결합한 결정적 문자열."""
        parts = [f"p:{p}" for p in sorted(set(self.phones))]
        parts += [f"e:{e}" for e in sorted(set(self.emails))]
        if not parts:  # 이론상 도달하지 않음(식별자 없는 응모는 무효 처리)
            parts = [f"r:{self.primary.row_no}"]
        return "|".join(parts)

    def matches(self, token: str, strict_email: bool = True) -> bool:
        """제외 대상 지정값(이메일/전화번호/이름)이 이 사람과 일치하는지."""
        tok = norm_text(token)
        if not tok:
            return False
        email = norm_email(tok, strict_email)
        if email and email in self.emails:
            return True
        phone = norm_phone(tok)
        if phone and phone in self.phones:
            return True
        return norm_name(tok) in [norm_name(n) for n in self.names]


def build_entries(rows: List[Dict[str, str]], cols: Dict[str, Optional[str]],
                  strict_email: bool = True) -> List[Entry]:
    entries: List[Entry] = []
    for i, row in enumerate(rows, start=2):  # 헤더가 1행이라 데이터는 2행부터
        def get(fieldname: str) -> str:
            col = cols.get(fieldname)
            return display_text(row.get(col, "")) if col else ""

        e = Entry(
            row_no=i,
            naming=get("naming"),
            name=get("name"),
            email=norm_email(get("email"), strict_email),
            phone=norm_phone(get("phone")),
            org=get("org"),
            extra=get("extra"),
            submitted_at=parse_dt(get("submitted_at")),
            raw=row,
        )
        if not e.naming:
            e.issues.append("네이밍 미기재")
        if not e.name:
            e.issues.append("성명 미기재")
        if not e.email and not e.phone:
            e.issues.append("연락처·이메일 모두 없음(당첨 연락 및 본인 확인 불가)")
        if get("email") and not e.email:
            e.issues.append("이메일 형식 오류")
        if get("phone") and not e.phone:
            e.issues.append("연락처 형식 오류")
        entries.append(e)
    return entries


def build_persons(entries: Iterable[Entry]) -> List[Person]:
    """이메일 또는 휴대폰번호가 하나라도 겹치면 같은 사람으로 병합(Union-Find).

    이름은 병합 근거로 쓰지 않는다 → 동명이인이 한 사람으로 합쳐지지 않는다.
    """
    valid = [e for e in entries if e.valid]
    parent: Dict[int, int] = {}

    def find(x: int) -> int:
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a: int, b: int) -> None:
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[max(ra, rb)] = min(ra, rb)

    by_email: Dict[str, int] = {}
    by_phone: Dict[str, int] = {}
    for idx, e in enumerate(valid):
        parent[idx] = idx
        if e.email:
            if e.email in by_email:
                union(idx, by_email[e.email])
            else:
                by_email[e.email] = idx
        if e.phone:
            if e.phone in by_phone:
                union(idx, by_phone[e.phone])
            else:
                by_phone[e.phone] = idx

    groups: Dict[int, List[Entry]] = {}
    for idx, e in enumerate(valid):
        groups.setdefault(find(idx), []).append(e)

    persons: List[Person] = []
    for root in sorted(groups):
        members = groups[root]
        p = Person(pid=len(persons) + 1, entries=members)
        p.emails = sorted({e.email for e in members if e.email})
        p.phones = sorted({e.phone for e in members if e.phone})
        persons.append(p)
    persons.sort(key=lambda p: p.primary.sort_key)
    for i, p in enumerate(persons, start=1):
        p.pid = i
    return persons


# --------------------------------------------------------------------------
# 이상 응모 점검
# --------------------------------------------------------------------------

# 호환 자모(ㄱ, ㅏ)와 NFKC 정규화 후의 조합용 자모(U+1100~U+11FF)를 모두 잡는다.
HANGUL_JAMO = re.compile(r"^[\u3131-\u318E\u1100-\u11FF]+$")
HAS_LETTER = re.compile(r"[0-9A-Za-z가-힣\u3131-\u318E\u1100-\u11FF]")


def naming_flags(naming: str, person_name: str = "") -> List[str]:
    """자동 실격이 아니라 '담당자 확인 필요' 표시용."""
    flags: List[str] = []
    s = norm_text(naming)
    body = re.sub(r"\s", "", s)
    if len(body) < 2:
        flags.append("2자 미만")
    if body and HANGUL_JAMO.match(body):
        flags.append("자음/모음만 입력")
    if body and not HAS_LETTER.search(body):
        flags.append("문자·숫자 없음")
    if body and len(set(body)) == 1 and len(body) >= 2:
        flags.append("같은 글자 반복")
    if person_name and norm_name(s) == norm_name(person_name):
        flags.append("본인 성명과 동일")
    return flags


def person_flags(p: Person, bulk_threshold: int) -> List[str]:
    flags: List[str] = []
    if p.submissions >= bulk_threshold:
        flags.append(f"다량 제출({p.submissions}건)")
    if len(p.names) > 1:
        flags.append("동일 연락처에 다른 성명(" + ", ".join(p.names) + ")")
    junk = sum(1 for e in p.entries if naming_flags(e.naming, e.name))
    if junk:
        flags.append(f"형식 미흡 제안 {junk}건")
    return flags


def duplicate_name_groups(persons: Sequence[Person]) -> Dict[str, List[Person]]:
    """동명이인 현황(이름은 같지만 식별자가 달라 서로 다른 사람으로 판정된 경우)."""
    by_name: Dict[str, List[Person]] = {}
    for p in persons:
        key = norm_name(p.name)
        if key:
            by_name.setdefault(key, []).append(p)
    return {k: v for k, v in by_name.items() if len(v) > 1}


# --------------------------------------------------------------------------
# 추첨 (결정적 · 검증 가능)
# --------------------------------------------------------------------------

def ticket_hash(seed: str, identity_key: str) -> str:
    return hmac.new(seed.encode("utf-8"), identity_key.encode("utf-8"), hashlib.sha256).hexdigest()


def roster_hash(persons: Sequence[Person]) -> str:
    """추첨 대상 명부의 지문. 명부가 한 글자라도 바뀌면 값이 달라진다."""
    h = hashlib.sha256()
    for key in sorted(p.identity_key for p in persons):
        h.update(key.encode("utf-8"))
        h.update(b"\n")
    return h.hexdigest()


def draw_winners(persons: Sequence[Person], seed: str, winners: int,
                 reserve: int = 0) -> Tuple[List[Person], List[Person]]:
    ranked = sorted(persons, key=lambda p: (ticket_hash(seed, p.identity_key), p.identity_key))
    return list(ranked[:winners]), list(ranked[winners:winners + reserve])


# --------------------------------------------------------------------------
# 출력
# --------------------------------------------------------------------------

def write_csv(path: str, header: Sequence[str], rows: Iterable[Sequence[Any]]) -> None:
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    # utf-8-sig: 한글 윈도우 엑셀에서 바로 열어도 깨지지 않는다.
    with open(path, "w", encoding="utf-8-sig", newline="") as fp:
        w = csv.writer(fp)
        w.writerow(header)
        for row in rows:
            w.writerow(row)


def mask_email(email: str) -> str:
    if "@" not in email:
        return email
    local, domain = email.split("@", 1)
    keep = local[:2] if len(local) > 2 else local[:1]
    return f"{keep}{'*' * max(len(local) - len(keep), 1)}@{domain}"


def mask_phone(phone: str) -> str:
    return f"{phone[:3]}****{phone[-4:]}" if len(phone) >= 8 else phone


def mask_name(name: str) -> str:
    if len(name) <= 1:
        return name
    if len(name) == 2:
        return name[0] + "*"
    return name[0] + "*" * (len(name) - 2) + name[-1]


def sha256_text(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def file_sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fp:
        for chunk in iter(lambda: fp.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def person_label(p: Person) -> str:
    ident = p.emails[0] if p.emails else (p.phones[0] if p.phones else "")
    ident = mask_email(ident) if "@" in ident else mask_phone(ident)
    org = f" / {p.org}" if p.org else ""
    return f"{mask_name(p.name)}({ident}){org}"


# --------------------------------------------------------------------------
# 공통 로딩
# --------------------------------------------------------------------------

@dataclass
class Loaded:
    rows: List[Dict[str, str]]
    cols: Dict[str, Optional[str]]
    entries: List[Entry]
    persons: List[Person]

    @property
    def invalid(self) -> List[Entry]:
        return [e for e in self.entries if not e.valid]


def load(args: argparse.Namespace) -> Loaded:
    rows = read_table(args.input)
    if not rows:
        raise SystemExit(f"[오류] 데이터가 없습니다: {args.input}")
    overrides: Dict[str, str] = {}
    for item in getattr(args, "col", []) or []:
        if "=" not in item:
            raise SystemExit(f"[오류] --col 형식은 필드=열이름 입니다: {item}")
        k, v = item.split("=", 1)
        overrides[k.strip()] = v.strip()

    cols = detect_columns(list(rows[0].keys()), overrides)
    missing = [k for k in ("naming", "name") if not cols[k]]
    if not cols["email"] and not cols["phone"]:
        missing.append("email 또는 phone")
    if missing:
        raise SystemExit(
            "[오류] 다음 열을 찾지 못했습니다: " + ", ".join(missing) + "\n"
            f"       파일의 열: {list(rows[0].keys())}\n"
            "       --col naming=제안명 --col email=이메일 처럼 직접 지정하세요."
        )
    entries = build_entries(rows, cols, strict_email=not args.loose_email)
    persons = build_persons(entries)
    return Loaded(rows=rows, cols=cols, entries=entries, persons=persons)


def resolve_exclusions(persons: Sequence[Person], tokens: Sequence[str],
                       strict_email: bool, kind: str) -> List[Tuple[str, Person]]:
    resolved: List[Tuple[str, Person]] = []
    for token in tokens:
        hits = [p for p in persons if p.matches(token, strict_email)]
        if not hits:
            raise SystemExit(f"[오류] {kind} '{token}' 에 해당하는 응모자를 찾지 못했습니다.")
        if len(hits) > 1:
            detail = "\n".join(f"        - {person_label(p)} / 제출 {p.submissions}건" for p in hits)
            raise SystemExit(
                f"[오류] {kind} '{token}' 이(가) {len(hits)}명과 일치합니다(동명이인 가능성).\n"
                f"{detail}\n        이메일 또는 휴대폰번호로 다시 지정하세요."
            )
        resolved.append((token, hits[0]))
    return resolved


# --------------------------------------------------------------------------
# inspect: 추첨 전 현황 점검
# --------------------------------------------------------------------------

def cmd_inspect(args: argparse.Namespace) -> int:
    data = load(args)
    persons = data.persons
    multi = sorted([p for p in persons if p.submissions > 1],
                   key=lambda p: (-p.submissions, p.pid))
    dup_names = duplicate_name_groups(persons)

    print("=" * 72)
    print(" 응모 현황 점검")
    print("=" * 72)
    print(f" 입력 파일        : {args.input}")
    print(f" 인식된 열        : " + ", ".join(f"{k}={v}" for k, v in data.cols.items() if v))
    print(f" 총 응모 건수     : {len(data.entries)}건")
    print(f" 유효 응모 건수   : {len(data.entries) - len(data.invalid)}건")
    print(f" 무효 응모 건수   : {len(data.invalid)}건")
    print(f" 추첨 대상 인원   : {len(persons)}명  ← 응모권은 1인 1장")
    print(f" 2건 이상 제출자  : {len(multi)}명")
    if persons:
        top = max(persons, key=lambda p: p.submissions)
        print(f" 최다 제출        : {top.submissions}건 ({person_label(top)})")
    print(f" 동명이인 그룹    : {len(dup_names)}건")

    if multi:
        print("\n[다건 제출자] — 몇 건을 내든 당첨 확률은 1명분으로 동일합니다.")
        for p in multi[: args.top]:
            flags = person_flags(p, args.bulk_threshold)
            note = f"  ⚑ {', '.join(flags)}" if flags else ""
            print(f"  - {p.submissions:>3}건  {person_label(p)}{note}")
        if len(multi) > args.top:
            print(f"  ... 외 {len(multi) - args.top}명 (--top 으로 더 보기)")

    if dup_names:
        print("\n[동명이인 확인] — 이름이 같지만 연락처가 달라 별개 인원으로 판정했습니다.")
        for name, group in list(dup_names.items())[: args.top]:
            print(f"  - {mask_name(group[0].name)} : {len(group)}명")
            for p in group:
                print(f"      · {person_label(p)} / 제출 {p.submissions}건")

    if data.invalid:
        print("\n[무효 응모] — 연락 불가·필수항목 누락 건")
        for e in data.invalid[: args.top]:
            print(f"  - {e.row_no}행: {', '.join(e.issues)} / 제안='{e.naming[:20]}'")
        if len(data.invalid) > args.top:
            print(f"  ... 외 {len(data.invalid) - args.top}건")

    review = [(e, naming_flags(e.naming, e.name)) for e in data.entries if e.valid]
    review = [(e, f) for e, f in review if f]
    if review:
        print(f"\n[형식 미흡 제안 {len(review)}건] — 자동 실격시키지 않았습니다. 담당자 확인 필요")
        for e, flags in review[: args.top]:
            print(f"  - {e.row_no}행: '{e.naming[:24]}' ({', '.join(flags)})")
        if len(review) > args.top:
            print(f"  ... 외 {len(review) - args.top}건")

    print("\n* 콘솔 출력의 개인정보는 마스킹되어 있습니다. 전체 목록은 draw 실행 시 CSV로 저장됩니다.")
    return 0


# --------------------------------------------------------------------------
# draw: 참가상 추첨
# --------------------------------------------------------------------------

def cmd_draw(args: argparse.Namespace) -> int:
    data = load(args)
    persons = data.persons
    strict_email = not args.loose_email

    winners_excl = resolve_exclusions(persons, args.exclude_winner, strict_email, "수상자 제외 지정")
    other_excl = resolve_exclusions(persons, args.exclude, strict_email, "기타 제외 지정")

    if len(winners_excl) > args.max_top_winners:
        raise SystemExit(
            f"[오류] 상위 수상자 제외는 최대 {args.max_top_winners}명입니다"
            f"(현재 {len(winners_excl)}명). --max-top-winners 로 조정하세요."
        )

    excluded: Dict[int, Tuple[Person, str, str]] = {}
    for token, p in winners_excl:
        excluded[p.pid] = (p, "상위 수상(최우수·우수)", token)
    for token, p in other_excl:
        excluded.setdefault(p.pid, (p, "기타 제외", token))

    roster = [p for p in persons if p.pid not in excluded]
    if len(roster) < args.winners:
        raise SystemExit(f"[오류] 추첨 대상 {len(roster)}명 < 당첨 인원 {args.winners}명")

    seed = args.seed.strip()
    if not seed:
        raise SystemExit("[오류] --seed 는 비워 둘 수 없습니다.")
    r_hash = roster_hash(roster)
    won, reserve = draw_winners(roster, seed, args.winners, args.reserve)

    outdir = args.out
    os.makedirs(outdir, exist_ok=True)
    now = datetime.now(KST)

    # 1) 당첨자 명단 (실제 연락용 · 개인정보 포함)
    write_csv(
        os.path.join(outdir, "01_당첨자.csv"),
        ["구분", "순번", "성명", "이메일", "연락처", "소속", "제안명(대표)", "제출건수", "추첨해시"],
        [
            [kind, i, p.name, p.emails[0] if p.emails else "", p.phones[0] if p.phones else "",
             p.org, p.primary.naming, p.submissions, ticket_hash(seed, p.identity_key)[:16]]
            for kind, group in (("참가상", won), ("예비", reserve))
            for i, p in enumerate(group, start=1)
        ],
    )

    # 2) 추첨 대상 명부 (1인 1행)
    write_csv(
        os.path.join(outdir, "02_추첨대상_명부.csv"),
        ["번호", "성명", "이메일", "연락처", "소속", "제출건수", "대표 제안명", "추첨해시", "비고"],
        [
            [p.pid, p.name, ";".join(p.emails), ";".join(p.phones), p.org, p.submissions,
             p.primary.naming, ticket_hash(seed, p.identity_key)[:16],
             " / ".join(person_flags(p, args.bulk_threshold))]
            for p in roster
        ],
    )

    # 3) 제외자
    write_csv(
        os.path.join(outdir, "03_제외자.csv"),
        ["성명", "이메일", "연락처", "소속", "제출건수", "제외 사유", "지정값"],
        [
            [p.name, ";".join(p.emails), ";".join(p.phones), p.org, p.submissions, reason, token]
            for p, reason, token in excluded.values()
        ],
    )

    # 4) 무효 응모
    write_csv(
        os.path.join(outdir, "04_무효응모.csv"),
        ["원본 행", "네이밍", "성명", "이메일", "연락처", "소속", "사유"],
        [
            [e.row_no, e.naming, e.name, e.email, e.phone, e.org, " / ".join(e.issues)]
            for e in data.invalid
        ],
    )

    # 5) 확인 필요 제안
    write_csv(
        os.path.join(outdir, "05_확인필요_제안.csv"),
        ["원본 행", "네이밍", "성명", "소속", "확인 사유"],
        [
            [e.row_no, e.naming, e.name, e.org, " / ".join(naming_flags(e.naming, e.name))]
            for e in data.entries if e.valid and naming_flags(e.naming, e.name)
        ],
    )

    # 6) 감사 기록 (개인정보 없이 재현 검증에 필요한 값만)
    audit = {
        "program": "naming-contest-raffle",
        "version": VERSION,
        "drawn_at_kst": now.strftime("%Y-%m-%d %H:%M:%S"),
        "input_file": os.path.basename(args.input),
        "input_sha256": file_sha256(args.input),
        "seed": seed,
        "params": {
            "winners": args.winners,
            "reserve": args.reserve,
            "strict_email_normalize": strict_email,
            "bulk_threshold": args.bulk_threshold,
        },
        "counts": {
            "entries_total": len(data.entries),
            "entries_invalid": len(data.invalid),
            "persons_total": len(persons),
            "excluded": len(excluded),
            "roster": len(roster),
        },
        "roster_sha256": r_hash,
        "excluded_identity_sha256": sorted(sha256_text(p.identity_key) for p, _, _ in excluded.values()),
        "winners": [
            {"rank": i, "identity_sha256": sha256_text(p.identity_key),
             "ticket_hash": ticket_hash(seed, p.identity_key)}
            for i, p in enumerate(won, start=1)
        ],
        "reserve_list": [
            {"rank": i, "identity_sha256": sha256_text(p.identity_key),
             "ticket_hash": ticket_hash(seed, p.identity_key)}
            for i, p in enumerate(reserve, start=1)
        ],
    }
    with open(os.path.join(outdir, "audit.json"), "w", encoding="utf-8") as fp:
        json.dump(audit, fp, ensure_ascii=False, indent=2)

    # 7) 보고용 요약 (개인정보 마스킹)
    lines = [
        "# 네이밍 공모전 참가상 추첨 결과",
        "",
        f"- 추첨 일시 : {now.strftime('%Y-%m-%d %H:%M:%S')} (KST)",
        f"- 대상 파일 : {os.path.basename(args.input)} (SHA-256 {audit['input_sha256'][:16]}…)",
        f"- 공개 시드 : `{seed}`",
        f"- 명부 지문 : `{r_hash[:32]}…`",
        "",
        "## 1. 집계",
        "",
        "| 구분 | 수량 |",
        "| --- | ---: |",
        f"| 총 응모 건수 | {len(data.entries)}건 |",
        f"| 무효 응모 | {len(data.invalid)}건 |",
        f"| 응모자(사람 단위) | {len(persons)}명 |",
        f"| 상위 수상 등 제외 | {len(excluded)}명 |",
        f"| 추첨 대상 | {len(roster)}명 |",
        f"| 당첨(참가상) | {len(won)}명 |",
        f"| 예비 당첨 | {len(reserve)}명 |",
        "",
        "## 2. 당첨자 (개인정보 마스킹)",
        "",
        "| 순번 | 응모자 | 제출건수 | 추첨해시(앞 16자리) |",
        "| ---: | --- | ---: | --- |",
    ]
    for i, p in enumerate(won, start=1):
        lines.append(f"| {i} | {person_label(p)} | {p.submissions} | `{ticket_hash(seed, p.identity_key)[:16]}` |")
    if reserve:
        lines += ["", "## 3. 예비 당첨자", "", "| 순위 | 응모자 |", "| ---: | --- |"]
        lines += [f"| {i} | {person_label(p)} |" for i, p in enumerate(reserve, start=1)]
    lines += [
        "",
        "## 4. 공정성 확보 조치",
        "",
        "- (1인 1표) 동일인이 여러 건을 제출해도 응모권은 1장으로 집계 → 다량 제출에 의한 확률 상승 원천 차단",
        "- (동일인 판별) 이메일·휴대폰번호 정규화 후 하나라도 일치하면 동일인으로 병합. 성명은 병합 근거로 쓰지 않아 동명이인은 별개 인원으로 유지",
        "- (중복 수상 방지) 최우수·우수 수상자는 사람 단위로 제외되어 해당자의 다른 제출건도 함께 제외",
        "- (검증 가능) 공개 시드와 명부 지문으로 제3자가 `verify` 명령을 통해 동일 결과 재현 가능",
        "",
        "*본 요약의 개인정보는 마스킹되어 있으며, 연락용 전체 명단은 `01_당첨자.csv` 참조.*",
    ]
    with open(os.path.join(outdir, "report.md"), "w", encoding="utf-8") as fp:
        fp.write("\n".join(lines) + "\n")

    print("=" * 72)
    print(" 참가상 추첨 결과")
    print("=" * 72)
    print(f" 총 응모 {len(data.entries)}건 / 무효 {len(data.invalid)}건")
    print(f" 응모자 {len(persons)}명 → 제외 {len(excluded)}명 → 추첨 대상 {len(roster)}명")
    print(f" 공개 시드 : {seed}")
    print(f" 명부 지문 : {r_hash}")
    print("-" * 72)
    for i, p in enumerate(won, start=1):
        print(f" {i:>2}. {person_label(p)}  (제출 {p.submissions}건)")
    if reserve:
        print("-" * 72)
        for i, p in enumerate(reserve, start=1):
            print(f" 예비 {i}. {person_label(p)}")
    print("-" * 72)
    print(f" 결과 파일 : {os.path.abspath(outdir)}")
    print(" (01_당첨자.csv / 02_추첨대상_명부.csv / 03_제외자.csv / 04_무효응모.csv /")
    print("  05_확인필요_제안.csv / audit.json / report.md)")
    return 0


# --------------------------------------------------------------------------
# verify: 결과 재현 검증
# --------------------------------------------------------------------------

def cmd_verify(args: argparse.Namespace) -> int:
    with open(args.audit, "r", encoding="utf-8") as fp:
        audit = json.load(fp)

    args.loose_email = not audit["params"].get("strict_email_normalize", True)
    args.col = getattr(args, "col", []) or []
    data = load(args)

    checks: List[Tuple[str, bool, str]] = []

    actual_file = file_sha256(args.input)
    checks.append(("응모 명부 파일 동일성", actual_file == audit["input_sha256"],
                   f"기록 {audit['input_sha256'][:16]}… / 실제 {actual_file[:16]}…"))

    excluded_set = set(audit.get("excluded_identity_sha256", []))
    roster = [p for p in data.persons if sha256_text(p.identity_key) not in excluded_set]
    checks.append(("제외 인원 수", len(data.persons) - len(roster) == audit["counts"]["excluded"],
                   f"기록 {audit['counts']['excluded']}명 / 실제 {len(data.persons) - len(roster)}명"))
    checks.append(("추첨 대상 인원", len(roster) == audit["counts"]["roster"],
                   f"기록 {audit['counts']['roster']}명 / 실제 {len(roster)}명"))

    r_hash = roster_hash(roster)
    checks.append(("명부 지문(SHA-256)", r_hash == audit["roster_sha256"],
                   f"기록 {audit['roster_sha256'][:16]}… / 실제 {r_hash[:16]}…"))

    seed = audit["seed"]
    won, reserve = draw_winners(roster, seed, audit["params"]["winners"], audit["params"]["reserve"])
    recomputed = [{"rank": i, "identity_sha256": sha256_text(p.identity_key),
                   "ticket_hash": ticket_hash(seed, p.identity_key)}
                  for i, p in enumerate(won, start=1)]
    checks.append(("당첨자 재현", recomputed == audit["winners"], f"{len(won)}명 비교"))

    recomputed_reserve = [{"rank": i, "identity_sha256": sha256_text(p.identity_key),
                           "ticket_hash": ticket_hash(seed, p.identity_key)}
                          for i, p in enumerate(reserve, start=1)]
    checks.append(("예비 당첨자 재현", recomputed_reserve == audit.get("reserve_list", []),
                   f"{len(reserve)}명 비교"))

    print("=" * 72)
    print(" 추첨 결과 검증")
    print("=" * 72)
    print(f" 감사 기록 : {args.audit}  (추첨 {audit['drawn_at_kst']} KST, v{audit['version']})")
    print(f" 공개 시드 : {seed}")
    print("-" * 72)
    ok = True
    for label, passed, detail in checks:
        print(f" [{'통과' if passed else '불일치'}] {label:<22} {detail}")
        ok = ok and passed
    print("-" * 72)
    if ok:
        print(" 결과: 감사 기록과 완전히 일치합니다. 추첨 결과가 재현되었습니다.")
        return 0
    print(" 결과: 불일치 항목이 있습니다. 명부가 변경되었거나 다른 조건으로 추첨된 기록입니다.")
    return 1


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="raffle.py",
        description="네이밍 공모전 참가상 추첨 (1인 1표 · 동명이인 구분 · 검증 가능)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {VERSION}")
    sub = parser.add_subparsers(dest="command", required=True)

    def common(p: argparse.ArgumentParser) -> None:
        p.add_argument("--input", "-i", required=True, help="응모 현황 파일 (.xlsx/.csv/.tsv)")
        p.add_argument("--col", action="append", default=[], metavar="필드=열이름",
                       help="열 자동 인식이 틀릴 때 직접 지정 "
                            "(필드: naming/name/email/phone/org/extra/submitted_at)")
        p.add_argument("--bulk-threshold", type=int, default=5,
                       help="다량 제출 표시 기준 건수 (기본 5)")

    p_inspect = sub.add_parser("inspect", help="추첨 전 응모 현황·중복 제출 점검")
    common(p_inspect)
    p_inspect.add_argument("--loose-email", action="store_true",
                           help="이메일 별칭(+태그, gmail 점) 정규화를 하지 않음")
    p_inspect.add_argument("--top", type=int, default=20, help="항목별 출력 개수 (기본 20)")
    p_inspect.set_defaults(func=cmd_inspect)

    p_draw = sub.add_parser("draw", help="참가상 추첨 실행")
    common(p_draw)
    p_draw.add_argument("--loose-email", action="store_true",
                        help="이메일 별칭(+태그, gmail 점) 정규화를 하지 않음")
    p_draw.add_argument("--winners", "-n", type=int, default=10, help="당첨 인원 (기본 10)")
    p_draw.add_argument("--reserve", "-r", type=int, default=5,
                        help="예비 당첨 인원 (연락 두절 대비, 기본 5)")
    p_draw.add_argument("--seed", required=True,
                        help="공개 시드. 추첨 전에 공표하거나 입회자와 함께 정한 문자열 "
                             "(예: '2026-09-15 추첨 입회 김OO')")
    p_draw.add_argument("--exclude-winner", action="append", default=[], metavar="이메일|전화|성명",
                        help="최우수·우수 수상자 제외 (반복 지정, 사람 단위로 제외됨)")
    p_draw.add_argument("--exclude", action="append", default=[], metavar="이메일|전화|성명",
                        help="기타 제외 대상 (부적격·본인 사퇴 등)")
    p_draw.add_argument("--max-top-winners", type=int, default=3,
                        help="상위 수상자 제외 허용 최대 인원 (기본 3 = 최우수1+우수2)")
    p_draw.add_argument("--out", "-o", default="추첨결과", help="결과 저장 폴더 (기본 '추첨결과')")
    p_draw.set_defaults(func=cmd_draw)

    p_verify = sub.add_parser("verify", help="audit.json으로 추첨 결과 재현 검증")
    p_verify.add_argument("--input", "-i", required=True, help="추첨에 사용한 응모 현황 파일")
    p_verify.add_argument("--audit", "-a", required=True, help="추첨 시 생성된 audit.json")
    p_verify.add_argument("--col", action="append", default=[], metavar="필드=열이름")
    p_verify.add_argument("--bulk-threshold", type=int, default=5, help=argparse.SUPPRESS)
    p_verify.set_defaults(func=cmd_verify)
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        return int(args.func(args))
    except BrokenPipeError:  # `| head` 등으로 출력이 잘린 경우
        try:
            sys.stdout.close()
        except BrokenPipeError:
            pass
        return 0
    except KeyboardInterrupt:
        print("\n[중단] 사용자가 실행을 취소했습니다.", file=sys.stderr)
        return 130


if __name__ == "__main__":
    sys.exit(main())
