# API 키 관리

**위치**: 리본 `EGTools++` 탭 → `EGTools` 그룹 → 스플릿 버튼 드롭다운 메뉴

공공 API 함수에 사용할 서비스별 API 키를 등록/삭제하는 대화상자입니다.

| 서비스 | 용도 | 관련 함수 |
|---|---|---|
| juso | 도로명주소 검색 | [SEARCHADDRESS](../SEARCHADDRESS.md) |
| odcloud | 사업자등록 상태 조회 | [BRNSTATUS](../BRNSTATUS.md) |
| vworld | 지도/좌표 | [GEOSEARCH](../GEOSEARCH.md) · [GEOCODER](../GEOCODER.md) · [GEOADDRESS](../GEOADDRESS.md) · [ADDRESSMAP](../ADDRESSMAP.md) |
| data.go.kr | 공휴일 | [KOREANHOLIDAYS](../KOREANHOLIDAYS.md) · [달력/일정표](calendar.md#공휴일-안내)(임시공휴일 조회용) |

- 서비스를 선택하면 키 등록 여부가 **마스킹(●●●●●●)으로만** 표시됩니다.
  저장된 키 원문은 다시 보여주지 않습니다.
- 키는 `HKCU\Software\EGTools\ApiKeys` 레지스트리에 저장되며 **통합문서나
  수식에는 남지 않습니다**.
- 함수의 **마지막 인수로 키를 직접 전달**하는 기존 방식과 공존합니다
  — 인수로 전달한 키는 자동으로 같은 저장소에 저장되어 이후에는 생략할 수 있습니다.
- **키는 EGTools++와 함께 배포되지 않습니다.** 각 서비스(juso.go.kr,
  odcloud/공공데이터포털, vworld.kr, data.go.kr)에서 사용자가 직접 발급받아야 합니다.
