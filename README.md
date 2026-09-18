# SpDocumentsConverter

성풍 출고장과 토글 주문을 이카운트·위하고 등의 Excel 업로드 양식으로 변환하는 데스크톱 프로그램입니다.
AMD64 Windows와 **macOS Sequoia 15.5 이상인 Apple Silicon Mac**을 대상으로 같은 소스와 PyInstaller 설정을 사용합니다.
이 브랜치는 기존 화면, 기능, 입출력 및 변환 규칙을 유지하면서 운영체제 호환성을 정리합니다.

## 실행 환경

- Python 3.14와 Tcl/Tk. `tkinter`는 표준 라이브러리지만 Python 설치에 Tcl/Tk가 포함되어 있어야 합니다.
- 저장된 Excel 파일을 읽는 기능은 `openpyxl`을 사용합니다.
- **현재 시트 / 선택 영역** 기능은 해당 컴퓨터에 설치된 Microsoft Excel을 `xlwings`로 제어합니다.
  `pip`가 Windows에서는 `pywin32`, macOS에서는 `appscript`와 `psutil`을 자동 설치합니다.
  Excel 추가 기능(add-in)은 필요하지 않습니다.

Mac의 Excel 연동은 처음 사용할 때 자동화 권한을 허용해야 합니다.
거부했다면 시스템 설정 → 개인정보 보호 및 보안 → 자동화에서 이 앱(소스 실행 시 터미널/Python)의 Excel 제어를 허용합니다.

Mac의 Excel은 샌드박스 앱이라 파일을 저장할 때 "파일 접근 권한 부여" 창이 뜨고, 취소하면 저장이 실패합니다.
그래서 Mac에서는 **현재 시트** 기능이 Excel에 파일 저장을 요청하지 않고, 자동화로 셀 값만 읽어
메모리에서 사본을 만들기 때문에 권한 창이 나타나지 않습니다.
Windows에서는 지금까지처럼 시스템 임시 폴더에 임시 사본을 저장해 사용합니다.

## 개발 환경 준비

Windows에서는 AMD64(64비트) Python을 설치하고 Tcl/Tk를 포함합니다. PowerShell에서:

```powershell
py -3.14 -m venv .venv
.venv\Scripts\python -m pip install -r requirements-build.txt
.venv\Scripts\python main.py
```

Mac 배포 빌드에는 [python.org의 Python 3.14 macOS 설치본](https://www.python.org/downloads/macos/)처럼
15.5 이하를 지원하는 Python과 Tcl/Tk를 사용합니다. 공식 설치본에는 Tcl/Tk가 포함됩니다.

```sh
/Library/Frameworks/Python.framework/Versions/3.14/bin/python3.14 -m venv .venv
.venv/bin/python -m pip install -r requirements-build.txt
.venv/bin/python main.py
```

Homebrew Python을 사용하는 로컬 개발 환경은 다음과 같이 준비할 수 있습니다:

```sh
brew install python@3.14 python-tk@3.14
python3.14 -m venv .venv
.venv/bin/python -m pip install -r requirements-build.txt
.venv/bin/python main.py
```

`python -m tkinter`로 Tcl/Tk 설치를 확인할 수 있습니다.
아래 명령의 `python`은 위에서 준비한 `.venv`의 Python을 사용합니다.
Python 설치본을 바꿀 때는 기존 가상환경을 재사용하지 않고 새로 만듭니다.

## 사용

- **첫 번째 출고장 탭**: Excel에서 문서를 열고 현재 시트 또는 선택 영역을 변환합니다.
  선택 영역 방식에는 제목 한 행을 가리키는 `TITLES` 이름 범위가 필요합니다.
- **두 번째 출고장 탭**: 파일과 시트를 선택한 뒤 실행합니다.
- **토글**: Excel의 현재 시트로 실행합니다. 기존 파일/시트 선택 UI의 동작도 유지합니다.

파일 선택 방식은 Excel이 마지막 저장한 수식 계산값을 읽으므로 원본의 변경 내용을 먼저 저장합니다.
기존 파일 선택창에는 `.xls`도 표시되지만 `openpyxl`로 직접 읽을 수는 없습니다.
구형 `.xls` 파일은 Excel에서 열어 현재 시트 방식으로 변환하거나 `.xlsx`로 저장합니다.
변환 결과는 운영체제의 임시 폴더에 생성되고 기본 연결 프로그램으로 열립니다. 보관할 파일은 원하는 위치에 저장합니다.

## 테스트와 빌드

```sh
python -m unittest discover -s tests -v
python main.py --smoke-test
python -m PyInstaller --noconfirm main.spec
```

- AMD64 Windows 결과: `dist/SpDocumentsConverter/SpDocumentsConverter.exe`.
  배포할 때는 같은 폴더의 `_internal` 등도 함께 전달합니다.
- Mac 결과: `dist/SpDocumentsConverter.app` (arm64, 최소 macOS 15.5).
  Tcl/Tk 및 품목 조회표 두 개(`soo_lu.xlsx`, `happy_lu.xlsx`)가 포함되며 Excel 자동화 권한 설명도 설정됩니다.
  번들에 넣는 파일 이름은 ASCII만 사용합니다. 한글 파일명은 받는 쪽 압축 해제 도구의 유니코드 정규화(NFC/NFD)로
  이름이 바뀌어 서명 봉인이 깨지고 "a sealed resource is missing or invalid"로 실행이 차단될 수 있습니다.
- 빌드는 각 운영체제에서 실행합니다. Windows에서 Mac 앱을 빌드하거나 반대로 빌드하지 않습니다.
- Python과 라이브러리는 결과물에 포함되므로 배포받는 사람은 별도로 설치하지 않아도 됩니다.
  현재 시트 / 선택 영역 기능을 쓰는 사람에게는 Microsoft Excel이 필요합니다.

GitHub Actions는 Windows x64와 Mac arm64에서 회귀 테스트, GUI 시작 검사, 빌드 및 패키징 후 시작 검사를 수행하도록 구성되어 있습니다.
Mac 빌드는 `macos-15` 러너에서 수행하며, 앱의 최소 버전을 15.5로 선언합니다.
빌드가 끝날 때 앱에 포함된 모든 Mach-O 파일의 아키텍처와 최소 OS를 검사하고,
15.5보다 새 OS를 요구하는 라이브러리가 있으면 빌드를 실패시킵니다.
현재 OS를 대상으로 컴파일한 Homebrew Python/Tk가 원인이 될 수 있습니다.
`MACOSX_DEPLOYMENT_TARGET=15.5`나 `Info.plist` 설정만으로 이미 컴파일된 라이브러리의 호환성을 낮출 수는 없습니다.

러너의 `macos-15` 라벨은 정확히 15.5를 고정하지 않습니다. 위 검사는 바이너리에 선언된 최소 버전을 확인하며,
실제 배포 검증에는 macOS 15.5에서 GUI 시작, Excel 현재 시트·선택 영역 변환, 결과 파일 열기도 포함합니다.
Mac 배포 압축에는 `.app`의 실행 권한과 심볼릭 링크를 보존합니다.
기본 빌드와 GitHub Actions 결과물은 ad-hoc 서명이며, 아래 절차로 Developer ID 서명과 공증을 추가할 수 있습니다.

자동 테스트는 Excel을 제어하지 않습니다. 실제 Excel에서 현재 시트 복사, 선택 영역 읽기,
생성된 파일 열기와 Mac의 최초 권한 요청은 각 운영체제에서 별도로 확인합니다.

## Mac 외부 배포: Developer ID 서명과 공증

유료 Apple Developer 계정의 **Developer ID Application** 인증서와 개인 키가 이 Mac의 키체인에 있어야 합니다.
Xcode → Settings → Apple Accounts(또는 Accounts)에서 유료 개발자 팀을 선택하고,
Manage Certificates → + → Developer ID Application으로 발급합니다.
`Apple Development`는 개발용 인증서이므로 이 배포 절차에 사용하지 않습니다.
설치된 서명 인증서의 정확한 이름은 다음 명령으로 확인합니다.

```sh
security find-identity -v -p codesigning
```

앞서 준비한 macOS 15.5 호환 Python 환경에서 인증서 이름을 지정해 빌드합니다.
PyInstaller가 포함된 라이브러리와 앱을 함께 서명하고 Hardened Runtime을 적용합니다.
결과는 기본 빌드와 구분하여 `dist/release`에 생성합니다.

```sh
MACOS_CODESIGN_IDENTITY='Developer ID Application: YOUR_NAME (TEAM_ID)' \
MACOSX_DEPLOYMENT_TARGET=15.5 \
python -m PyInstaller --clean --noconfirm \
  --distpath dist/release --workpath build/release main.spec

codesign --verify --deep --strict dist/release/SpDocumentsConverter.app
```

빌드와 서명 검사가 성공하면 **공증 없이 서명한 앱을 내부 공유**할 수도 있습니다.

```sh
ditto -c -k --sequesterRsrc --keepParent \
  dist/release/SpDocumentsConverter.app dist/SpDocumentsConverter-macos-arm64-signed.zip
```

이 ZIP을 받은 사람은 압축을 풀고 앱을 응용 프로그램 폴더로 옮깁니다.
공증이 없으므로 첫 실행이 차단될 수 있습니다. 직접 전달받은 이 앱을 실행하려면,
한 번 실행을 시도한 다음 시스템 설정 → 개인정보 보호 및 보안 → **그래도 열기**로 허용합니다.
이 앱에 대한 예외가 저장되면 이후에는 두 번 클릭해 실행할 수 있습니다.
회사 관리 정책이 예외 허용을 제한하는 Mac에서는 관리자 확인이 필요합니다.
자세한 절차는 [Apple 앱 실행 안내](https://support.apple.com/ko-kr/102445)를 참고합니다.

**공증 검증까지 통과하는 배포본**을 만들려면 다음 절차를 이어서 수행합니다.
공증은 Apple이 앱의 악성 코드와 서명 문제를 자동 검사하는 과정이며, 제출에는 별도 로그인이 필요합니다.
[Apple 계정에서 앱 암호를 생성](https://support.apple.com/ko-kr/102654)한 뒤,
아래 명령의 이메일과 Team ID를 해당 개발자 계정 값으로 바꾸어 본인 터미널에서 한 번 실행합니다.
앱 암호는 명령 실행 후 나타나는 보안 입력란에 입력하며, 소스나 명령 인자에 저장하지 않습니다.

```sh
xcrun notarytool store-credentials "SpDocumentsConverter" \
  --apple-id 'APPLE_ID_EMAIL' --team-id 'TEAM_ID'
```

공증용 ZIP을 만들어 Apple에 제출합니다.

```sh
ditto -c -k --sequesterRsrc --keepParent \
  dist/release/SpDocumentsConverter.app build/SpDocumentsConverter-notary.zip
xcrun notarytool submit build/SpDocumentsConverter-notary.zip \
  --keychain-profile "SpDocumentsConverter" --wait
```

결과가 **Accepted**일 때 앱에 공증 티켓을 붙이고 검증합니다.

```sh
xcrun stapler staple dist/release/SpDocumentsConverter.app &&
xcrun stapler validate dist/release/SpDocumentsConverter.app &&
codesign --verify --deep --strict dist/release/SpDocumentsConverter.app &&
spctl --assess --type execute --verbose=4 dist/release/SpDocumentsConverter.app &&
ditto -c -k --sequesterRsrc --keepParent \
  dist/release/SpDocumentsConverter.app dist/SpDocumentsConverter-macos-arm64-notarized.zip
```

전달할 파일은 마지막에 만든 `dist/SpDocumentsConverter-macos-arm64-notarized.zip`입니다.
공증 제출용 ZIP에는 아직 티켓이 붙지 않았으므로, 티켓을 붙인 앱을 다시 압축해 전달합니다.
서명·공증 후에도 실제 대상 Mac에서 앱 실행과 Excel 자동화 권한·변환을 확인합니다.

## 참고

- [Python tkinter](https://docs.python.org/3/library/tkinter.html)
- [xlwings 설치 및 운영체제별 의존성](https://docs.xlwings.org/en/stable/installation.html)
- [PyInstaller macOS 앱 빌드](https://pyinstaller.org/en/stable/usage.html#building-macos-app-bundles)
- [PyInstaller macOS 하위 버전 호환성](https://pyinstaller.org/en/stable/usage.html#making-macos-apps-forward-compatible)
- [Apple Developer ID 인증서](https://developer.apple.com/help/account/certificates/create-developer-id-certificates/)
- [Apple 공증 절차](https://developer.apple.com/documentation/security/customizing-the-notarization-workflow)
