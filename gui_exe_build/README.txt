YullinMoon AI GUI 실행기

1. dist\YullinMoon_AI_Server 폴더 전체를 함께 사용합니다.
2. YullinMoon_AI_Server.exe를 실행합니다.
3. 서버 실행 버튼을 누릅니다.
4. 실행 완료 후 웹 열기 또는 서버 끄기를 선택할 수 있습니다.
5. 창 오른쪽 위 X를 누르면 서버와 실행기가 모두 종료됩니다.

메일 계정과 API 설정:
- .env.example 파일을 .env로 복사하고 계정 정보를 입력하면 적용됩니다.
- .env에는 비밀번호가 있으므로 다른 사람에게 배포할 때 주의하세요.
- APPROVAL_TEST_MODE=true이면 고객 대신 APPROVAL_TEST_RECIPIENT로만 발송됩니다.
- 실제 고객 발송 전에는 테스트 모드를 유지하세요.

기존 견적서 위치:
- backend\data\quotation_files 폴더에 기존 견적서 파일을 넣으세요.
- 과거 DB에 예전 절대경로가 남아 있어도 동일한 파일명을 이 폴더에서 자동 검색합니다.
