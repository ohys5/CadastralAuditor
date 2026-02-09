# QField Auto Setup Plugin for QGIS

이 플러그인은 QField를 이용한 현장조사를 위해 QGIS 프로젝트를 자동으로 설정해줍니다.

## 주요 기능
- **자동 레이어 생성**: GeoPackage 기반의 포인트 레이어를 즉시 생성합니다.
- **표준 필드 구성**: `ID`, `날짜`, `메모`, `사진` 필드를 자동으로 추가합니다.
- **위젯 자동 설정**:
    - **사진 필드**: 현장에서 즉시 촬영 및 확인이 가능한 '첨부(Attachment)' 위젯 설정.
    - **날짜 필드**: 현재 시간이 자동으로 입력되도록 설정.
- **프로젝트 패키징 준비**: 생성된 레이어와 프로젝트를 지정된 경로에 저장하여 QField로 전송할 준비를 마칩니다.

## 설치 방법
1. `/home/ubuntu/qfield_auto_setup` 폴더를 복사합니다.
2. QGIS 플러그인 경로에 붙여넣습니다.
   - **Windows**: `%AppData%\QGIS\QGIS3\profiles\default\python\plugins`
   - **Linux/macOS**: `~/.local/share/QGIS/QGIS3/profiles/default/python/plugins`
3. QGIS를 재시작한 후 `플러그인 관리 및 설치` 메뉴에서 `QField Auto Setup`을 활성화합니다.

## 사용 방법
1. 툴바 또는 메뉴에서 `QField 자동 설정` 아이콘을 클릭합니다.
2. 조사 프로젝트 이름, 저장 경로, 좌표계(CRS)를 입력합니다.
3. `확인`을 누르면 레이어 생성부터 위젯 설정까지 자동으로 완료됩니다.
