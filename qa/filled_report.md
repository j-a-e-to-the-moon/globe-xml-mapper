# E2E 검증 리포트

- 입력 파일: `sample/main_template_samle_filled.xlsx`
- XSD: `Resources/XSD/GLOBEXML_v1.0_KR.xsd`
- 실행 시각: 2026-04-27 16:25:58

## 1. 매핑 오류

총 **0** 건

## 2. ValidationUtil 오류

총 **0** 건

## 3. XML 직렬화

- 저장: [qa/filled_output.xml](filled_output.xml)
- 크기: 58,196 bytes

## 4. XSD 검증 오류

총 **0** 건

## 종합

- 매핑 오류: 0
- Validation 오류: 0
- XSD 검증 오류: 0
- **총 0건**

---

## 해석 가이드

- **매핑 오류**: 서식 셀에 필수 데이터가 비어 있음 (사용자 입력 필요)
- **Validation 오류**: 코드의 70xxx 비즈니스 룰 위반 (ValidationUtil 구현 범위)
- **XSD 검증 오류**: 생성된 XML 구조가 스키마 위반 — 종류별:
  - `has invalid child element 'X'. List of possible elements expected: 'Y'` → Y([R])가 비어 있어 emit되지 않음 → **서식 샘플 미완성**
  - `is invalid. The value 'Z' is not valid` → 값 타입/포맷 위반 → **코드 또는 서식 포맷 버그**
  - `Schema location` 관련 → XSD 참조 설정 이슈 (`XmlUrlResolver` 사용 중)
