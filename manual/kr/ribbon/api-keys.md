# API 키 관리

**위치**: 리본 `EGTools++` 탭 → `EGTools` 그룹 → 스플릿 버튼 드롭다운 메뉴

공공 API 함수·리본 기능에 사용할 API 키를 등록/삭제하는 대화상자입니다.

목록에는 서비스가 **3개** 있습니다.

| 서비스 | 용도 | 관련 기능 |
|---|---|---|
| juso | 도로명주소 검색 | [SEARCHADDRESS](../SEARCHADDRESS.md) |
| data.go.kr | 공휴일·사업자등록 상태 | [KOREANHOLIDAYS](../KOREANHOLIDAYS.md) · [달력/일정표](calendar.md#공휴일-안내) · [BRNSTATUS](../BRNSTATUS.md) |
| vworld | 지도/좌표 | [GEOSEARCH](../GEOSEARCH.md) · [GEOCODER](../GEOCODER.md) · [GEOADDRESS](../GEOADDRESS.md) · [ADDRESSMAP](../ADDRESSMAP.md) |

- 서비스를 선택하면 키 등록 여부가 **마스킹(●●●●●●)으로만** 표시됩니다.
  저장된 키 원문은 다시 보여주지 않습니다.
- 키는 **사용자 계정에 저장**되며 **통합문서나 수식에는 남지 않습니다**.
- 함수의 **마지막 인수로 키를 직접 전달**하는 기존 방식과 공존합니다
  — 인수로 전달한 키는 자동으로 같은 저장소에 저장되어 이후에는 생략할 수 있습니다.
- **키는 EGTools++와 함께 배포되지 않습니다.** 각 서비스(juso.go.kr,
  data.go.kr, vworld.kr)에서 사용자가 직접 발급받아야 합니다.

## data.go.kr 키는 하나입니다

공공데이터포털(data.go.kr)은 **계정당 인증키를 하나만** 발급하고, 활용신청한
모든 서비스에 그 키를 함께 씁니다. 그래서 EGTools++도 서비스별로 나누지 않고
`data.go.kr` 한 칸으로 관리합니다.

- 이 한 개의 키를 [KOREANHOLIDAYS](../KOREANHOLIDAYS.md)(특일정보),
  [달력/일정표](calendar.md#공휴일-안내)의 임시공휴일 병합,
  [BRNSTATUS](../BRNSTATUS.md)(국세청 사업자등록 상태)가 **모두 공유**합니다.
- 예전 버전에서 **사업자등록용으로 따로 등록해 둔 키**는 처음 사용할 때
  통합된 칸으로 **자동 이관**됩니다 — 사용자가 다시 등록하거나 정리할 필요가
  없습니다.
- 포털이 보여주는 인코딩(Encoding)/디코딩(Decoding) 키 중 어느 쪽을 붙여넣어도
  됩니다. 필요한 경우 내부에서 인코딩해 사용합니다.
- 함수가 조회에 실패해 **키가 거부된 것으로 확인되면 저장된 키를 자동으로
  삭제**하고 재발급 안내를 돌려줍니다. 새 키를 발급받아 다시 등록하세요.
