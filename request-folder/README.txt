request-folder — 신청서를 넣는 폴더

사용법
1) settings 시트의 request_folder 값에 이 폴더(또는 이 폴더를 복사한 경로)를 지정하거나
   매크로 SelectRequestFolder 로 이 폴더를 선택합니다.
2) 팀별 하위폴더(<팀명>_<문서번호>) 안에 신청서 .xlsx 를 넣습니다.
   - 폴더명 예: 정보보호센터_1234  ->  request_team=정보보호센터, request_doc_no=1234
   - 마지막 '_' 기준으로 팀/문서번호가 갈립니다(팀명에 '_'가 있어도 됨).
   - '_'가 없으면 폴더명 전체가 팀명, 문서번호는 빈값.
3) Excel 에서 매크로 MergeFirewallRequestFolder 실행(Alt+F8) -> 통합 + 방화벽 대역 분석.

신청서 .xlsx 양식
- 첫 시트에 'No'(또는 '번호') 가 있는 행이 헤더 행입니다(이 폴더 파일은 B열에 No).
- 필요한 컬럼: 출발지IP 출발지 목적지IP 목적지 프로토콜 포트 방향 용도 시작일 종료일 비고
- 열 순서는 달라도 되고, 비표준 헤더는 settings!header_alias 로 매핑할 수 있습니다.
- 출발지IP/목적지IP 는 단일 IP, CIDR 대역, 세미콜론 구분 다중주소 모두 가능합니다.
- 새 신청서는 _빈양식/신청서_빈양식.xlsx 를 복사해서 작성하세요.

이 폴더에 들어 있는 신청 건은 워크북의 기본 firewall_ranges와 매칭되므로
그대로 통합/분석하면 적용대상방화벽(target_firewalls)이 채워집니다.
