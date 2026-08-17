# 첨부파일 읽기 (Attachment Reading)

메일 첨부파일을 텍스트로 뽑아내고, 채팅 에이전트가 그 내용을 근거로 답할 수 있게
하는 부분에 대한 설명. 관련 파일: `attachment_service.py`, `agent_tools.py`,
`agent_service.py`, `llm_service.py`.

## 큰 그림

에이전트(채팅 AI)는 원래 메일 본문과 견적 데이터만 볼 수 있고 첨부파일은 못 봤다.
`agent_tools.py`에 **`read_mail_attachments`라는 Tool**을 추가해서, 사용자가
"이 첨부파일에 뭐라고 써있어?" 같은 질문을 하면 에이전트가 스스로 이 Tool을
호출해 첨부파일 내용을 확인하고 답하도록 만들었다.

## 두 갈래: 텍스트형 vs 이미지형

첨부파일은 크게 두 가지 방식으로 나눠서 처리한다.

### (A) 이미 글자로 되어있는 파일 — PDF / 엑셀 / 한글 / PPT

메일을 처음 가져올 때(`mail_service.py`의 import 시점) `attachment_service.process_attachment()`가
미리 텍스트를 뽑아서 `Attachment.extracted_text`에 저장해둔다. 에이전트가 물어보면
DB에서 바로 꺼내서 돌려준다 (빠르고 비용 없음).

### (B) 사진·시안·스캔본처럼 텍스트가 없는 "그림" 파일

텍스트가 없으니 OpenAI Vision에게 "이 사진 좀 봐줘"라고 시켜야 한다. 매번 분석하면
느리고 비용도 들기 때문에, **사용자가 실제로 물어봤을 때 그 순간에 한 번만 분석**하고
결과를 `Attachment.analysis_summary`에 캐싱한다. 다음에 또 물어보면 재분석하지 않고
저장된 결과를 그대로 쓴다. (`agent_tools._ensure_attachment_text` / `_analyze_attachment_image`)

## 지원 확장자

| 확장자 | 처리 방식 | 함수 |
|---|---|---|
| `.pdf` (텍스트형) | 텍스트 추출 | `extract_pdf_text` |
| `.pdf` (스캔본, 텍스트 없음) | Vision 분석 대상으로 전환 | — |
| `.xlsx`, `.xlsm` | 셀 값 텍스트로 직렬화 | `extract_excel_text` |
| `.hwpx` | zip 내부 XML에서 텍스트 추출 | `extract_hwpx_text` |
| `.pptx` | zip 내부 슬라이드 XML에서 텍스트 추출 | `extract_pptx_text` |
| `.hwp` (구 한글, 바이너리) | OLE `PrvText`(미리보기) 스트림 추출 | `extract_hwp_preview_text` |
| `.txt`, `.csv`, `.log`, `.md` | 인코딩 추정 후 그대로 읽기 | `_read_text_file` |
| `.png`, `.jpg`, `.jpeg`, `.webp`, `.bmp`, `.gif` | Vision 분석 대상 | — |
| 확장자 없는 파일 | Pillow로 내용을 열어봐서 실제 이미지인지 판별 후 이미지면 Vision 대상으로 구제 | `_sniff_image_format` |

### `.pptx` / `.hwpx`가 새 의존성 없이 되는 이유

pptx와 hwpx는 사실 겉모습만 "오피스 파일"이지 내부적으로는 그냥 **zip 압축 파일**이다.
압축을 풀면 슬라이드/문단별 XML이 들어있고, 그 안에 `<a:t>텍스트</a:t>` 같은 태그로
글자가 들어있다. `zipfile` + `xml.etree`로 직접 열어서 텍스트 태그만 뽑아내면 되므로
별도 라이브러리가 필요 없다.

### `.hwp`가 `olefile`을 쓰는 이유

구 `.hwp`(한글 5.0)는 zip이 아니라 **OLE 복합 문서**(옛 `.doc`와 같은 계열)라서 본문을
완전히 파싱하려면 전용 파서가 필요하다. 대신 한글 프로그램은 파일 미리보기·검색
색인용으로 `PrvText`라는 스트림에 본문을 UTF-16으로 풀어서 같이 저장해둔다
(탐색기에서 hwp 파일에 마우스를 올리면 미리보기가 뜨는 것도 이 스트림 덕분).
`olefile`로 이 스트림만 읽어오는 방식으로 완전한 파서 없이도 근사 텍스트를 얻는다.

**한계:** `PrvText`는 "미리보기용 요약"이라 표/도형 안의 텍스트나 아주 긴 문서의
뒷부분은 빠질 수 있다. 그래서 상태는 `EXTRACTED`로 두되 `error_message`에
"미리보기 기반이라 불완전할 수 있음"이라는 캐비엇을 남겨두고, 에이전트 프롬프트에도
이 한계를 사용자에게 알리도록 규칙을 넣어뒀다.

### 진짜로 못 읽는 것들

- `.zip` — 압축 안의 파일들을 재귀적으로 다시 판별해야 해서 아직 미지원.
- 확장자는 `.pdf`인데 실제 PDF 헤더(`%PDF-`)가 아닌 파일 — 문서보안(DRM) 솔루션으로
  암호화된 경우가 많다. 파싱을 시도하지 않고 바로 `FAILED` 처리하며, 발신자에게
  원본을 다시 요청해야 한다는 안내를 남긴다.

## 흐름 요약

```
메일 import
  └─ process_attachment() 실행 → 텍스트형은 즉시 추출, 이미지형은 IMAGE_PENDING으로만 표시

사용자가 채팅에서 첨부파일/사진/시안 관련 질문
  └─ 에이전트가 read_mail_attachments Tool 호출
       ├─ extracted_text가 이미 있으면 그대로 반환
       └─ 없고 이미지/스캔본이면 그 자리에서 Vision 분석 → analysis_summary에 캐싱 후 반환
```

메일 최초 분석(`llm_service.analyze_mail`, 품목 자동 추출용)도 별도로 이미지 첨부를
Vision에 태워서 분석하지만, 이건 견적 품목을 뽑아내기 위한 일회성 호출이고
`read_mail_attachments`처럼 언제든 다시 물어볼 수 있는 대화형 조회는 아니다.
