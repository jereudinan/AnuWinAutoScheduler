# ANU Auto Scheduler (v1.0.1)

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)
![Python](https://img.shields.io/badge/python-3.9+-blue.svg)

**ANU Auto Scheduler**는 정해진 시간에 웹사이트를 자동으로 열어주는 윈도우 전용 스케줄링 도구입니다. 세련된 Fluent Design UI를 제공하며, 반복적인 웹 접속 업무를 자동화하는 데 최적화되어 있습니다.

## ✨ 주요 기능

- 🕒 **정밀한 스케줄링**: 시/분 단위로 원하는 작업을 예약할 수 있습니다.
- 🌐 **브라우저 선택**: Chrome 또는 Edge 브라우저를 선택하여 실행할 수 있습니다.
- 🚀 **부팅 시 자동 실행**: 윈도우 시작 시 프로그램이 자동으로 실행되도록 설정 가능합니다.
- 📥 **시스템 트레이**: 프로그램 종료 시 닫히지 않고 트레이 아이콘으로 최소화되어 백그라운드에서 동작합니다.
- 🌓 **테마 지원**: 사용자의 취향에 따라 라이트 모드와 다크 모드를 지원합니다.
- 🔔 **알람 기능**: 스케줄 실행 시 알람 소리와 함께 실행 여부를 묻는 팝업을 띄워줍니다.
- 🚫 **중복 실행 방지**: 프로그램이 이미 실행 중인 경우 기존 창을 활성화합니다.

## 🚀 설치 및 실행 방법

### 1. 소스 코드로 실행 (개발용)

Python이 설치되어 있어야 하며, 다음 라이브러리 설치가 필요합니다.

```bash
pip install PyQt6 qfluentwidgets
```

그 후 `main.py`를 실행합니다.

```bash
python main.py
```

### 2. 빌드된 파일 실행 (사용자용)

`dist/` 폴더 내의 실행 파일을 실행하거나 배포된 설치 프로그램을 사용하세요.

## 🛠 빌드 방법 (PyInstaller)

프로그램을 실행 파일로 빌드하려면 다음 명령어를 참고하세요.

```bash
pyinstaller --noconfirm --onedir --windowed --add-data "sounds;sounds" main.py
```

## ⚙️ 설정 관리

설정 정보는 프로그램 실행 경로의 `config.json` 파일에 저장됩니다.
- `run_at_startup`: 부팅 시 자동 실행 여부
- `theme`: UI 테마 (Light/Dark)
- `browser`: 사용할 브라우저 (Chrome/Edge)
- `schedules`: 등록된 스케줄 리스트

## 🐞 버그 제보 및 문의

프로그램 사용 중 발생하는 오작동이나 버그는 아래 이메일로 제보해 주세요.

- **이메일**: phb@somunnanshop.com
- **버전**: v1.0.1

---
## 📝 업데이트 이력 (Update History)

### 📅 2026-03-31 (v1.0.1)
- **✨ 신규 기능**: 기본 설정 메뉴에 **"바탕화면 바로가기 만들기"** 버튼 추가.
- **🛠 안정성 개선**:
  - 64비트 환경 및 권한 문제 해결을 위해 **윈도우 부팅 시 자동 실행** 로직 강화.
  - 프로그램 위치 이동 시에도 자동 실행 경로가 최신화되도록 개선.
- **🎨 UI/UX 개선**:
  - 작업 표시줄 및 트레이 아이콘을 **노란색 초시계(History) 아이콘**으로 변경.
  - 작업 표시줄 아이콘에 마우스 오버 시 프로그램 이름(`AnuAutoScheduler`)이 표시되도록 개선.
  - 트레이 아이콘에 툴팁 추가 및 윈도우 작업 표시줄 아이콘 고정 문제 해결.

---
© 2026 AnuAutoScheduler. All rights reserved.
