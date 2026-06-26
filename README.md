# 🔒 개인정보 검출 스캐너 (PII Scanner)

지정한 폴더 안의 파일을 읽어 **개인정보**(주민등록번호, 휴대폰, 이메일, 카드번호,
계좌번호, 여권·운전면허번호 등)를 검출하고, 해당 파일을 **격리**하거나 **삭제**할 수
있는 로컬 전용 프로그램입니다.

- ✅ 모든 처리는 **로컬에서만** 수행되며 외부 네트워크 전송이 전혀 없습니다.
- ✅ 리포트의 검출값은 기본적으로 **마스킹**되어 저장됩니다 (예: `900101-1******`).
- ✅ 검출된 파일을 **체크박스로 선택**(전체 선택 버튼 제공)해 **별도 폴더로 이동**하거나 삭제할 수 있습니다.
- ✅ 결과 목록에서 **파일명을 클릭**하면 기본 프로그램으로 바로 열 수 있고, **📂 폴더** 버튼으로 위치를 열 수 있습니다.
- ✅ 이동은 원본을 **옮기는**(복구 가능) 방식이고, 삭제는 가능하면 **휴지통**으로 보냅니다.

지원 형식: `.txt .csv .md .log .json .tsv .docx .xlsx .xlsm .pdf .hwp`

---

## 📥 EXE 다운로드 (설치 없이 실행)

Python 설치 없이 Windows에서 바로 실행할 수 있는 `.exe` 파일을 제공합니다.

### 방법 A — GitHub Actions 아티팩트에서 받기
1. 이 저장소의 **Actions** 탭으로 이동합니다.
2. **Build Windows EXE** 워크플로의 최신 성공 실행을 엽니다.
   (수동 실행: 워크플로 화면의 **Run workflow** 버튼)
3. 하단 **Artifacts** 의 `개인정보검출스캐너-windows` 를 내려받아 압축을 풉니다.
4. `개인정보검출스캐너.exe` 더블클릭 → 끝.

### 방법 B — Releases 에서 받기 (버전 태그 배포 시)
`v1.0.0` 같은 태그를 푸시하면 자동으로 **Releases** 페이지에 `.exe` 가 첨부됩니다.

```bash
git tag v1.0.0
git push origin v1.0.0
```

> 💡 `.exe` 는 Windows 전용입니다. macOS/Linux 에서는 아래 "직접 실행"을 사용하세요.

---

## 🛠️ 직접 EXE 빌드하기 (Windows)

저장소를 받은 뒤 Windows에서:

```bat
build.bat
```

또는 수동으로:

```bat
pip install -r requirements.txt pyinstaller
pyinstaller pii_scanner.spec --noconfirm
```

빌드 결과: `dist\개인정보검출스캐너.exe`

---

## 🐍 Python 으로 직접 실행 (Windows / macOS / Linux)

```bash
pip install -r requirements.txt

# GUI 실행
python pii_scanner_gui.py

# CLI 실행
python pii_scanner.py ./검사할폴더
python pii_scanner.py ./문서 --out report.csv --html report.html
python pii_scanner.py ./문서 --types rrn,phone,card   # 특정 유형만
```

---

## 📂 파일 구성

| 파일 | 설명 |
| --- | --- |
| `pii_scanner.py` | 검출 엔진 + CLI |
| `pii_scanner_gui.py` | 뉴모피즘 GUI (customtkinter) |
| `pii_scanner.spec` | PyInstaller 빌드 설정 |
| `requirements.txt` | 의존 라이브러리 |
| `build.bat` | Windows 원클릭 빌드 스크립트 |
| `.github/workflows/build-exe.yml` | EXE 자동 빌드/배포 (GitHub Actions) |

---

## ⚠️ 주의

- "실제 값 표시(마스킹 해제)" 를 켜면 화면·리포트에 실제 개인정보가 그대로 노출됩니다.
- **삭제**는 되돌리기 어려울 수 있으니, 가급적 **격리**를 먼저 사용하세요.
