# Globe XML Mapper - 작업 가이드

## 프로젝트 개요
국조 서식(.xlsx) → OECD GloBE XML 변환 WinForms 프로그램.
**단일 xlsx 파일** + **Control Panel** 방식 — Excel COM Interop으로 엑셀 직접 제어.

## UX 흐름
```
1. 메인 화면: [파일 열기] / [새 파일 만들기] / [템플릿 다운로드]
2. 파일 열기/생성 → Excel COM으로 xlsx 열기
3. Control Panel 표시 (TopMost, 드래그 이동, 접기/펼치기)
   ┌─────────────────────────────┐
   │ Globe XML Mapper       [─]  │
   ├─────────────────────────────┤
   │ 최종모기업 (UPE)  2개  [+][-]│
   │ 구성기업 (CE)     3개  [+][-]│
   │ 제외기업           1개  [+][-]│
   ├─────────────────────────────┤
   │      [XML 변환하기]          │
   └─────────────────────────────┘
4. [+] 시트 복제 / [-] 마지막 시트 삭제
5. [XML 변환하기] → 저장 후 ClosedXML로 매핑 + 검증
6. 엑셀 종료 → 메인 화면 복귀
7. 윈폼 종료 → 엑셀 저장 확인 후 함께 종료
```

## template.xlsx 시트 구조
- 국조53부표2 (1): 1.1~1.2 + 1.3.1 (기본정보 + 첫 번째 UPE)
- 국조53부표2 (2): 1.3.2.1 (첫 번째 CE)
- 국조53부표2 (3): 1.3.2.2 (첫 번째 제외기업)
- (4)~(25): 빈 시트 (추가 시 복제 대상)

## 숨김시트 (_META)
섹션→시트 매핑 메타정보. ExcelController가 자동 생성/관리.
```
section    | sheetName
1.1~1.2    | 국조53부표2 (1)
1.3.1      | 국조53부표2 (1)      ← (1)에 포함
1.3.2.1    | 국조53부표2 (2)
1.3.2.2    | 국조53부표2 (3)
```
시트 추가/삭제 시 메타도 자동 갱신.
XML 변환 시 MappingOrchestrator가 메타 참조.

## 프로젝트 구조
```
mapper/
├── GlobeMapper.csproj
├── Program.cs
├── MainForm.cs                  # 메인 화면 (3버튼)
├── ControlPanelForm.cs          # Control Panel (TopMost, 시트 추가/삭제, XML 변환)
├── TermsDialog.cs
├── Services/
│   ├── ExcelController.cs        # Excel COM Interop 래퍼
│   ├── MappingBase.cs            # 공통 유틸
│   ├── Mapping_1.1~1.2.cs       # 기본정보 매퍼
│   ├── Mapping_1.3.1.cs         # UPE 매퍼
│   ├── Mapping_1.3.2.1.cs       # CE 매퍼 (별첨 시트 포함)
│   ├── Mapping_1.3.2.2.cs       # 제외기업 매퍼
│   ├── MappingOrchestrator.cs   # 메타 기반 매핑 + FillMessageSpec
│   ├── ValidationUtil.cs
│   └── XmlExportService.cs
├── Resources/
│   ├── Globe.cs                  # XSD 생성 클래스
│   ├── XSD/
│   ├── mappings/                 # 섹션별 매핑 JSON
│   ├── template.xlsx             # 원본 템플릿 (빌드 출력 포함)
│   ├── terms.txt
│   ├── Globe_Enum_정리.csv
│   └── GIR_XML_에러코드_매뉴얼.md
└── Tools/                        # Python 스크립트
```

## 핵심 클래스

### ExcelController (Services/ExcelController.cs)
- `Open(path)` / `CreateNew(template, savePath)` — 엑셀 열기
- `AddSectionSheet(section)` — 템플릿 시트 복제 + 메타 등록
- `RemoveSectionSheet(section)` — 마지막 시트 삭제 + 메타 제거
- `GetSectionSheets(section)` / `GetSectionCounts()` — 메타 조회
- `GetFilePathForMapping()` — 저장 후 파일 경로 반환 (ClosedXML용)
- `CloseWithSavePrompt()` — 저장 확인 후 종료
- `WorkbookClosed` 이벤트 — 엑셀 종료 감지

### MappingOrchestrator
- `MapWorkbook(filePath, globe)` — _META 기반 매핑 (Control Panel 방식)
- `MapFolder(rootPath, globe)` — 디렉토리 기반 매핑 (하위 호환)

### CE 행 블록 반복 (시트2)
- 시트2에서 4~21행이 CE 1건 블록. [+]로 블록 복제, 각 블록의 O14셀에 "별첨N" 자동 갱신
- 별첨 시트(`부표2 (2) 별첨`)에 별첨1, 별첨2... 섹션이 자동 추가/삭제
- Mapping_1.3.2.1이 blockCount 기반으로 N개 CE 순회, 별첨N에서 Ownership 읽기
- 별첨 시트에서 주주 행 추가/삭제는 Control Panel에서 별첨 번호 선택 후 [+][-]

## 새 섹션 추가 절차
1. mapping JSON: `Resources/mappings/mapping_{섹션}.json`
2. 매퍼 클래스: `Services/Mapping_{섹션}.cs` (MappingBase 상속)
3. MappingOrchestrator.MapperFactory에 등록
4. ExcelController.SectionTemplateIndex에 템플릿 시트 인덱스 등록
5. ControlPanelForm.Sections에 표시명 추가
6. ValidationUtil 검증 추가
7. CLAUDE.md 업데이트

## 구현 진행 상황
- [x] 1.1~1.2: 신고구성기업, 사업연도, 회계정보
- [x] 1.3.1: 최종모기업 (UPE) — 행 블록 반복
- [x] 1.3.2.1: 구성기업 (CE) — 행 블록 반복 + 첨부 시트 소유지분
- [x] 1.3.2.2: 제외기업 — 행 블록 반복
- [x] 1.3.3: 기업구조 변동 — 단순 행 추가 (XML 미포함)
- [x] 1.4: 정보 요약 (Summary) — 단순 행 추가
- [x] 2: 국가별 적용면제 — 52행 복합 블록 반복 + Summary 상세
- [ ] 3.1~3.2.3.2: 글로벌최저한세 계산 (JurisdictionSection/ETR)
- [ ] 2: Summary
- [ ] 3: JurisdictionSection / ETR
- [ ] 4: UTPRAttribution
