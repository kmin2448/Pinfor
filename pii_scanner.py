#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
개인정보 검출 로컬 스캐너 (PII Scanner)
지정한 폴더 안의 파일들을 읽어 개인정보(주민등록번호, 휴대폰, 이메일,
카드번호, 계좌번호, 외국인등록번호, 여권/운전면허번호 등)를 검출합니다.

- 모든 처리는 로컬에서만 수행되며 외부 네트워크 전송이 전혀 없습니다.
- 검출 결과 리포트에는 실제 값이 마스킹되어 저장됩니다 (예: 900101-1******).
  → 리포트 파일 자체가 개인정보 유출원이 되지 않도록 한 설계입니다.

사용법:
    python pii_scanner.py <폴더경로>
    python pii_scanner.py ./문서 --out report.csv --html report.html
    python pii_scanner.py ./문서 --types rrn,phone,card    # 특정 유형만
    python pii_scanner.py ./문서 --reveal                  # (주의) 마스킹 해제

필요 라이브러리(읽을 파일 형식에 따라):
    pip install python-docx openpyxl pdfplumber
    pip install olefile        # .hwp 지원 시
"""

import os
import re
import sys
import csv
import shutil
import argparse
from datetime import datetime
from html import escape

QUARANTINE_PREFIX = "_PII_격리_"

# ───────────────────────────────────────────────────────────
# 선택적 라이브러리 (없으면 해당 형식만 건너뜀)
# ───────────────────────────────────────────────────────────
try:
    import docx  # python-docx
except ImportError:
    docx = None
try:
    import openpyxl
except ImportError:
    openpyxl = None
try:
    import pdfplumber
except ImportError:
    pdfplumber = None
try:
    import olefile
except ImportError:
    olefile = None


# ───────────────────────────────────────────────────────────
# 1. 검증 함수 (오탐 줄이기)
# ───────────────────────────────────────────────────────────
def validate_rrn(digits: str) -> bool:
    """주민등록번호 13자리 체크섬 검증."""
    if len(digits) != 13 or not digits.isdigit():
        return False
    # 생년월일 형식 1차 점검
    mm, dd = int(digits[2:4]), int(digits[4:6])
    if not (1 <= mm <= 12 and 1 <= dd <= 31):
        return False
    weights = [2, 3, 4, 5, 6, 7, 8, 9, 2, 3, 4, 5]
    s = sum(int(digits[i]) * weights[i] for i in range(12))
    check = (11 - (s % 11)) % 10
    return check == int(digits[12])


def luhn(num: str) -> bool:
    """신용카드 Luhn 알고리즘 검증."""
    num = re.sub(r"\D", "", num)
    if not (13 <= len(num) <= 19):
        return False
    digits = [int(d) for d in num]
    odd = digits[-1::-2]
    even = digits[-2::-2]
    total = sum(odd) + sum(sum(divmod(d * 2, 10)) for d in even)
    return total % 10 == 0


# ───────────────────────────────────────────────────────────
# 2. 검출기 정의
#    severity: 3=치명적, 2=높음, 1=중간
# ───────────────────────────────────────────────────────────
DETECTORS = {
    "rrn": {
        "label": "주민/외국인등록번호",
        "severity": 3,
        "pattern": re.compile(r"\d{6}[-\s]?[1-8]\d{6}"),
        "validate": lambda m: validate_rrn(re.sub(r"\D", "", m)),
    },
    "card": {
        "label": "신용카드번호",
        "severity": 3,
        "pattern": re.compile(r"\b(?:\d[ -]?){12,18}\d\b"),
        "validate": lambda m: luhn(m),
    },
    "phone": {
        "label": "휴대폰번호",
        "severity": 2,
        "pattern": re.compile(r"01[016789][-.\s]?\d{3,4}[-.\s]?\d{4}"),
        "validate": None,
    },
    "account": {
        "label": "계좌번호(추정)",
        "severity": 2,
        # 은행 계좌: 2~6 / 2~6 / 2~7 형태의 하이픈 구분 숫자
        "pattern": re.compile(r"\b\d{2,6}-\d{2,6}-\d{2,7}\b"),
        # YYYY-MM-DD 같은 날짜 형식은 제외
        "validate": lambda m: not re.fullmatch(r"(19|20)\d{2}-\d{1,2}-\d{1,2}", m),
    },
    "email": {
        "label": "이메일",
        "severity": 1,
        "pattern": re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"),
        "validate": None,
    },
    "passport": {
        "label": "여권번호(추정)",
        "severity": 2,
        "pattern": re.compile(r"\b[MSRODmsrod]\d{8}\b"),
        "validate": None,
    },
    "driver": {
        "label": "운전면허번호(추정)",
        "severity": 2,
        "pattern": re.compile(r"\b\d{2}-?\d{2}-?\d{6}-?\d{2}\b"),
        "validate": None,
    },
}


# 위험도(severity) → 표시 이름
SEV_NAME = {3: "치명적", 2: "높음", 1: "중간"}


def mask(value: str) -> str:
    """검출값 마스킹: 앞부분만 남기고 나머지는 * 처리."""
    keep = max(1, len(value) // 3)
    return value[:keep] + "*" * (len(value) - keep)


# ───────────────────────────────────────────────────────────
# 3. 파일 → 텍스트 추출 (형식별)
#    반환: [(줄번호, 텍스트), ...]
# ───────────────────────────────────────────────────────────
def read_plain(path):
    for enc in ("utf-8", "cp949", "euc-kr"):
        try:
            with open(path, "r", encoding=enc) as f:
                return [(i + 1, line) for i, line in enumerate(f)]
        except (UnicodeDecodeError, UnicodeError):
            continue
    return []


def read_docx(path):
    if docx is None:
        return None
    d = docx.Document(path)
    lines = [(i + 1, p.text) for i, p in enumerate(d.paragraphs) if p.text.strip()]
    # 표 안의 텍스트도 검사
    n = len(lines)
    for table in d.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    n += 1
                    lines.append((n, cell.text))
    return lines


def read_xlsx(path):
    if openpyxl is None:
        return None
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    lines = []
    for ws in wb.worksheets:
        for r, row in enumerate(ws.iter_rows(values_only=True), start=1):
            for c, val in enumerate(row):
                if val is not None and str(val).strip():
                    lines.append((f"{ws.title}!행{r}열{c+1}", str(val)))
    wb.close()
    return lines


def read_pdf(path):
    if pdfplumber is None:
        return None
    lines = []
    with pdfplumber.open(path) as pdf:
        for pno, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            for ln in text.splitlines():
                lines.append((f"p{pno}", ln))
    return lines


def read_hwp(path):
    """HWP 5.0 본문 텍스트 추출(간이). 정확도는 제한적."""
    if olefile is None:
        return None
    import zlib, struct
    try:
        ole = olefile.OleFileIO(path)
        dirs = ole.listdir()
        text = []
        for entry in dirs:
            if entry[0] == "BodyText":
                data = ole.openstream(entry).read()
                try:
                    data = zlib.decompress(data, -15)
                except zlib.error:
                    pass
                # UTF-16LE 추정 텍스트만 추려냄
                try:
                    chunk = data.decode("utf-16le", errors="ignore")
                    text.append(re.sub(r"[\x00-\x1f]", " ", chunk))
                except Exception:
                    pass
        ole.close()
        return [(i + 1, ln) for i, ln in enumerate("\n".join(text).splitlines()) if ln.strip()]
    except Exception:
        return None


READERS = {
    ".txt": read_plain, ".csv": read_plain, ".md": read_plain,
    ".log": read_plain, ".json": read_plain, ".tsv": read_plain,
    ".docx": read_docx,
    ".xlsx": read_xlsx, ".xlsm": read_xlsx,
    ".pdf": read_pdf,
    ".hwp": read_hwp,
}


# ───────────────────────────────────────────────────────────
# 4. 스캔 로직
# ───────────────────────────────────────────────────────────
# 겹치는 구간 정리 우선순위 (구체적 → 일반적)
PRIORITY = ["rrn", "card", "phone", "passport", "driver", "account", "email"]


def scan_text(lines, active_types):
    """추출된 (줄번호, 텍스트) 목록에서 PII 검출. 겹치는 구간은 우선순위로 정리."""
    findings = []
    ordered = [k for k in PRIORITY if k in active_types]
    for lineno, text in lines:
        claimed = []  # 이미 검출된 (start, end) 구간
        for key in ordered:
            det = DETECTORS[key]
            for m in det["pattern"].finditer(text):
                value = m.group()
                if det["validate"] and not det["validate"](value):
                    continue  # 검증 실패 → 오탐으로 간주, 제외
                s, e = m.start(), m.end()
                if any(s < ce and e > cs for cs, ce in claimed):
                    continue  # 더 우선순위 높은 유형이 이미 차지한 구간
                claimed.append((s, e))
                findings.append({
                    "type": det["label"],
                    "type_key": key,
                    "severity": det["severity"],
                    "line": lineno,
                    "value": value.strip(),
                })
    return findings


def scan_file(path, active_types):
    ext = os.path.splitext(path)[1].lower()
    reader = READERS.get(ext)
    if reader is None:
        return None, "지원하지 않는 형식"
    try:
        lines = reader(path)
    except Exception as e:
        return None, f"읽기 오류: {e}"
    if lines is None:
        return None, "필요 라이브러리 미설치"
    return scan_text(lines, active_types), None


def walk_folder(folder, active_types, max_mb=50, progress=None):
    results = []   # 검출된 파일
    skipped = []   # 건너뛴 파일 (사유)
    # 진행률 표시를 위해 대상 파일을 먼저 수집 (격리 폴더는 제외)
    all_files = []
    for root, dirs, files in os.walk(folder):
        dirs[:] = [d for d in dirs if not d.startswith(QUARANTINE_PREFIX)]
        for name in files:
            all_files.append(os.path.join(root, name))
    total = len(all_files)
    for idx, path in enumerate(all_files, start=1):
        if progress:
            progress(idx, total, path)
        try:
            size_mb = os.path.getsize(path) / (1024 * 1024)
        except OSError:
            continue
        if size_mb > max_mb:
            skipped.append((path, f"용량 초과({size_mb:.1f}MB)"))
            continue
        findings, err = scan_file(path, active_types)
        if err:
            if err not in ("지원하지 않는 형식",):
                skipped.append((path, err))
            continue
        if findings:
            results.append((path, findings))
    return results, skipped


# ───────────────────────────────────────────────────────────
# 4-1. 격리 / 삭제 (파괴적 작업 — 호출부에서 반드시 확인 후 사용)
# ───────────────────────────────────────────────────────────
def quarantine_files(paths, base_folder, quarantine_root=None):
    """선택한 파일을 격리 폴더로 '이동'한다 (복구 가능).
    반환: (moved[(원본,격리경로)], errors[(경로,사유)], quarantine_root)"""
    if quarantine_root is None:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        quarantine_root = os.path.join(base_folder, f"{QUARANTINE_PREFIX}{stamp}")
    os.makedirs(quarantine_root, exist_ok=True)
    moved, errors = [], []
    for p in paths:
        try:
            rel = os.path.relpath(p, base_folder)
            if rel.startswith(".."):
                rel = os.path.basename(p)
        except ValueError:
            rel = os.path.basename(p)
        dest = os.path.join(quarantine_root, rel)
        # 이름 충돌 방지
        base, ext = os.path.splitext(dest)
        n = 1
        while os.path.exists(dest):
            dest = f"{base}({n}){ext}"
            n += 1
        try:
            os.makedirs(os.path.dirname(dest), exist_ok=True)
            shutil.move(p, dest)
            moved.append((p, dest))
        except Exception as e:
            errors.append((p, str(e)))
    return moved, errors, quarantine_root


def delete_files(paths, use_trash=True):
    """선택한 파일을 삭제한다.
    use_trash=True 이고 send2trash 가 설치돼 있으면 OS 휴지통으로(복구 가능),
    아니면 영구 삭제한다.
    반환: (deleted[경로], errors[(경로,사유)], to_trash[bool])"""
    trash = None
    if use_trash:
        try:
            from send2trash import send2trash as trash
        except ImportError:
            trash = None
    deleted, errors = [], []
    for p in paths:
        try:
            if trash:
                trash(p)
            else:
                os.remove(p)
            deleted.append(p)
        except Exception as e:
            errors.append((p, str(e)))
    return deleted, errors, (trash is not None)



def print_console(results, skipped, reveal):
    total = sum(len(f) for _, f in results)
    print("\n" + "=" * 60)
    print(f"  개인정보 검출 결과 — 파일 {len(results)}개에서 {total}건 발견")
    print("=" * 60)
    for path, findings in sorted(results, key=lambda x: -max(f["severity"] for f in x[1])):
        print(f"\n📄 {path}")
        by_type = {}
        for f in findings:
            by_type.setdefault(f["type"], []).append(f)
        for typ, items in sorted(by_type.items(), key=lambda x: -x[1][0]["severity"]):
            sev = SEV_NAME[items[0]["severity"]]
            print(f"   [{sev}] {typ}: {len(items)}건")
            for it in items[:5]:
                shown = it["value"] if reveal else mask(it["value"])
                print(f"        - {it['line']}: {shown}")
            if len(items) > 5:
                print(f"        ... 외 {len(items) - 5}건")
    if skipped:
        print(f"\n⚠️  건너뛴 파일 {len(skipped)}개:")
        for path, reason in skipped[:10]:
            print(f"   - {os.path.basename(path)}: {reason}")
    print()


def write_csv(results, out, reveal):
    with open(out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["파일경로", "위험도", "유형", "위치", "검출값(마스킹)"])
        for path, findings in results:
            for it in findings:
                shown = it["value"] if reveal else mask(it["value"])
                w.writerow([path, SEV_NAME[it["severity"]], it["type"], it["line"], shown])
    print(f"✅ CSV 리포트 저장: {out}")


def write_html(results, skipped, out, reveal):
    total = sum(len(f) for _, f in results)
    rows = []
    for path, findings in sorted(results, key=lambda x: -max(f["severity"] for f in x[1])):
        for it in findings:
            shown = it["value"] if reveal else mask(it["value"])
            color = {3: "#e74c3c", 2: "#e67e22", 1: "#f1c40f"}[it["severity"]]
            rows.append(
                f"<tr><td>{escape(path)}</td>"
                f"<td style='color:{color};font-weight:600'>{SEV_NAME[it['severity']]}</td>"
                f"<td>{escape(it['type'])}</td><td>{escape(str(it['line']))}</td>"
                f"<td><code>{escape(shown)}</code></td></tr>"
            )
    html = f"""<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8">
<title>개인정보 검출 리포트</title>
<style>
body{{font-family:'Malgun Gothic',sans-serif;background:#e0e5ec;color:#3a3f47;padding:32px}}
h1{{font-size:22px}} .meta{{color:#7a828c;margin-bottom:20px}}
table{{width:100%;border-collapse:collapse;background:#e0e5ec;border-radius:14px;
box-shadow:6px 6px 12px #bcc1c9,-6px -6px 12px #ffffff;overflow:hidden}}
th,td{{padding:10px 14px;text-align:left;font-size:13px;border-bottom:1px solid #d1d6dd}}
th{{background:#d8dde4}} code{{background:#d1d6dd;padding:2px 6px;border-radius:6px}}
</style></head><body>
<h1>🔒 개인정보 검출 리포트</h1>
<div class="meta">생성: {datetime.now():%Y-%m-%d %H:%M} ·
파일 {len(results)}개 · 총 {total}건 · 건너뜀 {len(skipped)}개</div>
<table><thead><tr><th>파일</th><th>위험도</th><th>유형</th><th>위치</th><th>검출값</th></tr></thead>
<tbody>{''.join(rows)}</tbody></table>
</body></html>"""
    with open(out, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"✅ HTML 리포트 저장: {out}")


# ───────────────────────────────────────────────────────────
# 6. CLI
# ───────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser(description="개인정보 검출 로컬 스캐너")
    ap.add_argument("folder", help="검사할 폴더 경로")
    ap.add_argument("--types", help=f"검출 유형(쉼표구분). 기본=전체. 선택지: {','.join(DETECTORS)}")
    ap.add_argument("--out", help="CSV 리포트 저장 경로")
    ap.add_argument("--html", help="HTML 리포트 저장 경로")
    ap.add_argument("--reveal", action="store_true", help="(주의) 마스킹 해제하고 실제 값 표시")
    ap.add_argument("--max-mb", type=float, default=50, help="검사 최대 파일 용량 MB (기본 50)")
    args = ap.parse_args()

    if not os.path.isdir(args.folder):
        print(f"❌ 폴더를 찾을 수 없습니다: {args.folder}")
        sys.exit(1)

    if args.types:
        active = [t.strip() for t in args.types.split(",") if t.strip() in DETECTORS]
        if not active:
            print(f"❌ 유효한 유형 없음. 선택지: {','.join(DETECTORS)}")
            sys.exit(1)
    else:
        active = list(DETECTORS)

    print(f"🔍 검사 시작: {args.folder}")
    print(f"   유형: {', '.join(DETECTORS[t]['label'] for t in active)}")
    results, skipped = walk_folder(args.folder, active, args.max_mb)

    print_console(results, skipped, args.reveal)
    if args.out:
        write_csv(results, args.out, args.reveal)
    if args.html:
        write_html(results, skipped, args.html, args.reveal)


if __name__ == "__main__":
    main()
