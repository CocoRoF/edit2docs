# edit2docs

**AI 에이전트 네이티브 문서 엔진: DOCX, XLSX, PPTX를 Python 라이브러리, 에이전트 도구, MCP 서버, 호스팅 서비스로 생성하고 채팅으로 편집합니다. English-first, 한국어는 완전한 1급 지원.**

[![PyPI](https://img.shields.io/pypi/v/edit2docs)](https://pypi.org/project/edit2docs/)
[![Python](https://img.shields.io/badge/python-3.12%2B-blue)](https://pypi.org/project/edit2docs/)
[![License: Apache-2.0](https://img.shields.io/badge/license-Apache--2.0-green)](./LICENSE)

[English README](./README.md)

## 무엇인가

LLM 에이전트가 Word/Excel/PowerPoint 파일을 "편집"할 때는 보통 파일 전체를
다시 생성하거나 그림으로 납작하게 만들어서 차트, 표, 서식, 수식이 사라집니다.
`edit2docs`는 에이전트(또는 스크립트)에게 실제 OOXML 위에서 동작하는 작은
**확장자 디스패치 동사** 묶음을 줍니다.

* **생성**: 한 줄 의도로 완성 문서를 만들거나(LLM), 직접 쓴 스펙으로
  결정적으로 만듭니다(키 불필요).
* **편집**: 채팅으로(LLM) 혹은 정확한 주소로(결정적) 기존 파일을 고칩니다.
  건드리지 않은 패키지 파트는 **바이트 단위로 동일**하게 남아, 차트·스파크라인·
  이미지·캐시된 수식 값이 부수 피해를 입지 않습니다.
* **조회·렌더**: 주소화된 아웃라인, 페이지 PNG/PDF/SVG, 마크다운.

결과물은 항상 네이티브 편집 가능(진짜 문단, 셀, 차트)합니다. 하나의 엔진,
네 가지 표면: Python 라이브러리, 에이전트 도구(function calling), 로컬 stdio
MCP 서버, 호스팅 FastAPI 서비스.

## 설치

Python 3.12 이상이 필요합니다.

```bash
pip install edit2docs              # 라이브러리 + 에이전트 도구 + 로컬 MCP 서버
pip install "edit2docs[server]"    # + 호스팅 멀티테넌트 서비스
```

선택 extras: `images`(Gemini/OpenAI 이미지 백엔드), `svg-fallback`(cairosvg),
`dev`(테스트/린트, `server` 포함). 저장소 `main`이 PyPI보다 앞서 있을 수 있습니다
([버전 이력](#버전-이력) 참고). 최신 main 설치:
`pip install "edit2docs @ git+https://github.com/CocoRoF/edit2docs"`.

## 빠른 시작

결정적 동사는 API 키가 필요 없습니다. 아래 코드는 그대로 실행됩니다.

```python
from edit2docs import build_doc, analyze_doc, set_doc_text, render_doc

build_doc({"slides": [
    {"layout": "title", "title": "3분기 리뷰"},
    {"layout": "content", "title": "핵심 성과", "bullets": ["매출 +12%", "이탈률 감소"]},
]}, "deck.pptx", lang="ko-KR")

info = analyze_doc("deck.pptx")      # 정확한 주소가 담긴 아웃라인
title = next(t for t in info["slides"][1]["texts"] if t["text"] == "핵심 성과")
# title == {"para": 0, "text": "핵심 성과", "shape_id": 2}

r = set_doc_text("deck.pptx", [
    {"slide": 1, "shape_id": title["shape_id"], "para": title["para"], "new_text": "주요 결과"},
])
print(r.path, r.applied)             # deck_edited.pptx 1  (입력 파일은 덮어쓰지 않음)

render_doc("deck_edited.pptx", to="png", out_dir="pages")   # 페이지 PNG, LibreOffice 불필요
```

생성형 동사는 Anthropic 키(`api_key=` 또는 `ANTHROPIC_API_KEY`)가 필요합니다.

```python
from edit2docs import generate_doc, edit_doc

generate_doc("3분기 영업 실적 임원 보고", output="deck.pptx", lang="ko-KR")
r = edit_doc("deck.pptx", "3번 슬라이드 제목을 더 단정적으로 바꿔줘", lang="ko-KR")
print(r.reply)        # 편집기가 무엇을 바꿨는지 설명
print(r.path)         # deck_edited.pptx (output=...으로 지정 가능)
```

## 동사

파일 확장자가 엔진을 고릅니다. 아래는 모두 `edit2docs`의 라이브러리 함수입니다.
`det`는 결정적·키 불필요, `LLM`은 BYOK입니다.

| 동사 | 하는 일 | |
|---|---|---|
| `build_doc(spec, output)` | 직접 쓴 스펙으로 문서 생성: 마크다운(docx), `{"sheets": [...]}`(xlsx), `{"slides": [...], "theme": {...}}`(pptx, 테마 덱 포함) | det |
| `generate_doc(intent, *, output, ...)` | 의도(+ 선택: `sources`, PPTX `template`) → 디자인된 문서 | LLM |
| `edit_doc(doc, instruction, ...)` | 자연어 편집 1턴; 건드리지 않은 내용은 바이트 단위로 동일 | LLM |
| `analyze_doc(doc)` | 편집 동사가 쓰는 정확한 주소가 담긴 구조 아웃라인 + `charts` 목록 | det |
| `set_doc_text(doc, edits, *, output=None)` | 표적 편집: docx 문단/표 셀, xlsx 셀/행/시트, pptx 도형 문단/표 셀, 네이티브 차트 제목/데이터(`chart` 인덱스가 있는 편집) | det |
| `edit_chart(doc, edits)` / `list_charts(doc)` | 네이티브 차트의 데이터·제목 편집(차트 XML과 임베디드 워크북 동기화), docx/xlsx/pptx | det |
| `arrange_doc(doc, ops, *, output=None)` | 구조: 슬라이드(pptx)·시트(xlsx) 통째로 복제 / 이동 / 삭제, 시트 이름 변경 | det |
| `list_doc_parts(doc)` / `get_doc_xml(doc, part)` / `set_doc_xml(doc, part, ...)` | 패키지 파트 목록 또는 한 파트의 XML 읽기; 파트 패치·생성·삭제(색, 폰트, 위치 등 무엇이든) | det |
| `preview_doc(doc, *, out_dir=None)` | pptx → 슬라이드별 SVG, docx/xlsx → 마크다운 | det |
| `render_doc(doc, *, to="png", out_dir=None, dpi=144.0)` | `to` = `png` / `pdf` / `svg` / `md`; resvg + PyMuPDF, LibreOffice 불필요 | det |
| `doc_guide(topic=None)` | 에이전트용 점진 공개 가이드(토픽: build, generate, edit, edit.text, edit.chart, arrange, edit.xml, recipes.slides, recipes.colors, render) | det |

비동기 변형: `async_generate_doc`, `async_edit_doc`. PPTX 전용 별칭(`generate_pptx`,
`edit_pptx`, `preview_pptx`, `set_pptx_text`, `analyze_pptx`)도 있습니다.

결과 객체: `GenerateResult(path, page_count, design_spec, warnings)`, `EditResult`
(`.reply`, `.operations`, `.path`), `TextEditsResult(path, applied, results)`
(편집별 `status`: `applied | stale | not_found | invalid`), `ArrangeResult`,
`RenderResult(paths, page_count, format, to)`. 파생 파일을 쓰는 동사는 기본적으로
`<stem>_edited.<ext>`로 저장합니다(`edit_chart`는 `_chart`, `arrange_doc`은
`_arranged`). `output=`으로 지정하세요.

### 생성

```python
generate_doc("3분기 실적 보고서", output="report.docx", lang="ko-KR")
generate_doc("분기 매출 요약", output="sales.xlsx", sources=["raw.pdf"])
generate_doc("3분기 영업 실적 임원 보고", output="deck.pptx",
             template="brand.pptx",          # 선택: 사용자 PPTX 템플릿
             deck_mode="template_restyle",   # "new" | "template_restyle" | "template_extend"
             pages=(8, 12))                  # 목표 페이지 범위 (pptx)
```

`sources`는 PDF / DOCX / DOC / PPTX / XLSX / HTML / EPUB / IPYNB 경로를 받으며,
각각 마크다운으로 변환되어 작성기의 참고 자료가 됩니다. 전체 시그니처:
`generate_doc(intent, *, output, api_key=None, sources=None, template=None,
deck_mode="new", pages=(8, 12), lang="en-US", model=...)`.

### 채팅 편집

```python
r = edit_doc("report.docx", "진행 현황 섹션에 '배포 완료' 항목을 추가해줘", lang="ko-KR")
print(r.reply, r.operations)

r = edit_doc("deck.pptx", "이 문서 내용을 반영해서 3번 슬라이드를 고쳐줘",
             sources=["notes.pdf"], lang="ko-KR",
             chat_history=[{"role": "user", "content": "..."},
                           {"role": "assistant", "content": "..."}])
```

플래너는 문서의 번호 매긴 아웃라인을 보고 최소한의 연산을 계획하며, 결정적
엔진이 이를 적용합니다. 계획에 실패하면 조용히 아무 일도 안 하는 대신 그렇다고
답합니다. 덱의 표적 편집은 실제 도형에 대해 제자리·바이트 보존 연산으로
실행됩니다(25개 연산, 예: `set_text`, `set_runs`, `set_table_cell`, `insert_row`,
`set_chart_data`, `add_slide`, `set_notes`, `set_hyperlink`, `set_theme_color`).
필요한 새 슬라이드·전면 재디자인 턴은 슬라이드를 SVG로 재생성하는 경로로
넘어갑니다. 연산 카탈로그는 코드로 조회할 수 있습니다(아래).

### 결정적 편집

```python
set_doc_text("report.docx", [
    {"action": "replace", "para": 0, "new_text": "3분기 최종 보고서"},
    {"action": "insert_after", "para": 0, "markdown": "새 문단"},
    {"action": "replace", "table": 0, "row": 1, "col": 2, "new_text": "142"},
    {"action": "delete", "para": 5},
])
set_doc_text("sales.xlsx", [
    {"action": "set_cell", "sheet": "Sales", "cell": "B3", "value": 142},
    {"action": "append_rows", "sheet": "Sales", "rows": [["Q4", 160]]},
    {"action": "add_sheet", "sheet": "Notes", "headers": ["a"], "rows": [["x"]]},
])
set_doc_text("deck.pptx", [{"slide": 1, "shape_id": 5, "para": 0, "new_text": "새 제목"}])

edit_chart("deck.pptx", [
    {"chart": 0, "title": "3분기 매출"},
    {"chart": 0, "categories": ["Q1", "Q2", "Q3"],
     "series": [{"name": "매출", "values": [120, 135, 150]}]},
])

arrange_doc("deck.pptx", [{"op": "duplicate", "target": 0}, {"op": "move", "target": 2, "to": 0}])
```

선택 `old_text` / `old_value` 가드는 오래된(stale) 편집을 거부합니다. 주소는
`analyze_doc`에서 가져오세요. 전체 형식은 `doc_guide("edit.text")`로 확인합니다.

PPTX 수술형 엔진은 직접 호출도 가능합니다.

```python
from edit2docs.documents.pptx_engine import apply_pptx_edits, describe_ops, find_shapes

describe_ops()                                    # 기계가 읽는 연산 카탈로그 (연산별 주소 + 페이로드)
find_shapes(pptx_bytes, text="매출")               # 텍스트/종류/이름/슬라이드로 도형 조회
new_bytes, results = apply_pptx_edits(pptx_bytes, edits, dry_run=True)   # 쓰지 않고 검증
# atomic=True -> 전부 적용 또는 전부 롤백
```

## 에이전트 도구 (function calling)

`doc_guide`, `analyze_doc`, `render_doc`, `set_doc_text`, `arrange_doc`,
`read_doc_xml`, `set_doc_xml`, `build_doc`, `generate_doc`, `edit_doc`가 도구
스키마와 디스패처로 노출됩니다.

```python
import anthropic
from edit2docs.agent_tools import ANTHROPIC_TOOLS, run_tool

client = anthropic.Anthropic()
msg = client.messages.create(
    model="claude-sonnet-5",
    max_tokens=2048,
    tools=ANTHROPIC_TOOLS,
    messages=[{"role": "user", "content": "deck.pptx의 3번 슬라이드 제목을 고쳐줘"}],
)
for block in msg.content:
    if block.type == "tool_use":
        result = run_tool(block.name, block.input)   # run_tool_async도 있음
```

`OPENAI_TOOLS`는 OpenAI 방식 function calling 형태입니다.
`tool_specs(provider, extension="xlsx")`는 해당 파일 형식에 맞는 동사만(형식별로
특화된 설명과 함께) 돌려주며, 확장자를 주지 않으면 전체 집합을 돌려줍니다. 이
표면에서 차트 편집은 `set_doc_text`(`chart` 인덱스가 있는 편집)로 합니다.

## 로컬 MCP 서버 (인프라 불필요)

`pip install edit2docs`로 설치되는 `edit2docs-mcp` stdio 서버가 같은 열 개
동사를 로컬 파일에 대해 노출합니다.

```jsonc
// Claude Desktop / Claude Code / Cursor
{
  "mcpServers": {
    "edit2docs": {
      "command": "edit2docs-mcp",
      "env": { "ANTHROPIC_API_KEY": "sk-ant-..." }   // 생성형 도구만 필요
    }
  }
}
```

설치 없이: `uvx --from edit2docs edit2docs-mcp`. 클라이언트별 설정은
[docs/mcp-clients.md](./docs/mcp-clients.md)를 보세요(그 문서는 호스팅 서버의
도구 목록을 설명하며 로컬 서버와 다릅니다).

## 호스팅 서비스

```bash
pip install "edit2docs[server]"
edit2docs serve [--host H] [--port P] [--reload]    # FastAPI, 기본 포트 8000
```

단독 모드는 외부 인프라가 필요 없습니다: `EDIT2DOCS_DATA_DIR` 아래 SQLite + 로컬
파일시스템 저장소 + 인라인 작업 큐. Postgres, Redis(arq 워커 큐), S3 호환 저장소는
환경 변수로 켭니다.

| 엔드포인트 | 용도 |
|---|---|
| `POST /v1/assets`, `GET/DELETE /v1/assets/{id}` | 문서 업로드 / 조회 / 삭제 (기본 200 MB 상한) |
| `POST /v1/jobs/generate-deck` | 생성 작업 큐잉; `output_format` = `pptx` / `docx` / `xlsx` |
| `POST /v1/jobs/edit-deck` | 채팅 편집 작업 큐잉 |
| `GET /v1/jobs/{id}`, `GET /v1/jobs/{id}/events` | 상태; SSE 진행 스트림(단계 + 주소화된 대상이 있는 연산별 실시간 편집 이벤트) |
| `POST /v1/preview` | pptx → 슬라이드별 SVG, docx/xlsx → 주소화된 HTML(`data-e2d-*`) |
| `POST /v1/text-edits` | 결정적 표적 편집 |
| `GET /v1/models` | 실시간 모델 목록(호출자의 키로 Anthropic Models API 중계, 키 없으면 내장 목록) |
| `GET /health` | 생존 확인 및 모드 보고 |
| `/mcp`, `/mcp-sse` | Streamable HTTP / SSE 위의 MCP (덱 작업 중심 도구: `generate_deck`, `edit_deck`, `upload_source`, `request_upload_url`, `get_asset`, `download_url`, `list_templates`, `list_voices`, `hello`) |

Anthropic 키는 요청마다 BYOK입니다(`X-Anthropic-API-Key` 헤더). 오류는 이중 언어:
`message`가 `Accept-Language`를 따르고 영어/한국어 변형도 함께 제공됩니다.

### 설정

환경 변수, 접두사 `EDIT2DOCS_` (`src/edit2docs/config.py` 참고):

| 변수 | 기본값 | 비고 |
|---|---|---|
| `EDIT2DOCS_DEFAULT_LANG` | `en-US` | `ko-KR`로 두면 배포가 한국어 기본 |
| `EDIT2DOCS_HOST` / `EDIT2DOCS_PORT` | `0.0.0.0` / `8000` | `serve` 플래그로 덮어쓰기 가능 |
| `EDIT2DOCS_DATA_DIR` | `/data/edit2docs` | SQLite + 파일 저장소 루트 |
| `EDIT2DOCS_DATABASE_URL` | 데이터 디렉터리 아래 SQLite | 예: `postgresql+asyncpg://...` |
| `EDIT2DOCS_REDIS_URL` | 미설정(인라인 큐) | arq 워커 큐 활성화 |
| `EDIT2DOCS_S3_ENDPOINT_URL` + `EDIT2DOCS_S3_BUCKET` | 미설정(로컬 fs) | `S3_REGION`, `S3_ACCESS_KEY_ID`, `S3_SECRET_ACCESS_KEY`, `S3_PUBLIC_BASE_URL`과 함께 |
| `EDIT2DOCS_RAW_SIGNING_KEY` | 시작할 때마다 무작위 | 운영에서는 지정해야 서명된 로컬 URL이 재시작 후에도 유효 |
| `EDIT2DOCS_AUTH_DEV_API_KEY` | 미설정(익명) | 소규모 배포용 단일 bearer 토큰 |
| `EDIT2DOCS_MAX_UPLOAD_SIZE_BYTES` | 200 MB | 리버스 프록시와 맞추세요 |
| `EDIT2DOCS_MODEL_{PLANNER,WRITER,STRATEGIST,EXECUTOR}` | 요청 모델 | 역할별 모델 덮어쓰기 |
| `EDIT2DOCS_STRATEGIST_SOURCE_CHAR_CAP` | `60000` | 덱 전략가에 넣는 소스별 상한 (0 = 무제한) |

이미지 생성 백엔드(`images` extra)는 `IMAGE_BACKEND`, `OPENAI_API_KEY`,
`GEMINI_API_KEY` 같은 제공자 변수를 읽습니다. `.env.example`을 참고하세요.

### 웹 스튜디오

[**edit2docs-web**](https://github.com/CocoRoF/edit2docs-web)은 공식 프론트엔드입니다:
업로드, 단계별 SSE 진행이 있는 생성, 각 연산이 건드리는 문단 / 셀 / 슬라이드를
캔버스에 하이라이트하는 채팅 편집 스튜디오. 엔진 연결은
`EDIT2DOCS_SERVER_INTERNAL_URL`과 `EDIT2DOCS_SERVER_API_KEY`로 합니다.

## 포맷별 동작

* **DOCX**: 작성 LLM이 제약된 마크다운을 내고, 결정적 python-docx 렌더러가 파일을
  만듭니다. 편집은 문단/표 셀 주소 기반입니다. 프리뷰는 주소화된 네이티브 HTML
  (`data-e2d-para`, `data-e2d-table`, `data-e2d-cell`)입니다.
* **XLSX**: 디자이너 LLM이 YAML 시트 스펙을 내고 openpyxl이 렌더합니다. 편집은
  stale 가드가 있는 `set_cell` / `append_rows` / `add_sheet`입니다. 프리뷰는
  `data-e2d-cell="B3"` 주소와 캐시된 수식 값이 있는 그리드입니다.
* **PPTX**: strategist → 페이지별 SVG → 네이티브 DrawingML, 사용자 템플릿
  restyle/extend, 선택적 Edge-TTS 내레이션. 채팅 편집은 위에서 설명한 제자리
  수술형 엔진을 씁니다.

무손실 편집은 [contextifier](https://github.com/CocoRoF/Contextifier)의 raw OOXML
레이어(`contextifier>=0.8.0`) 위에 있습니다: 편집이 건드린 파트만 다시 씁니다.

### 네이티브 차트와 표 (PPTX)

`data-pptx-native="chart|table"`로 표시한 SVG 그룹은 실제 PowerPoint 차트(차트 파트 +
임베디드 워크북) 또는 네이티브 `<a:tbl>` 표로 내보내집니다.

```xml
<g id="sales_chart" data-pptx-native="chart">
  <metadata data-pptx-native="chart">
    { "name": "sales_chart", "x": 125, "y": 141, "width": 1000, "height": 440,
      "type": "bar", "categories": ["Q1", "Q2", "Q3"],
      "series": [{ "name": "Sales", "values": [120, 135, 150] }] }
  </metadata>
  <!-- 네이티브 내보내기를 끄면 쓰이는 대체 도형 -->
</g>
```

`ExportRequest(native_objects=True)`로 켭니다(`tools/export.py`).

## 언어

기본은 영어(`lang="en-US"`)이고, 한국어는 1급 지원입니다: 한글 폭 인식, 실제
문자에서 감지한 런별 OOXML `lang`, 한국어 폰트 스택, 완전한 메시지 카탈로그,
현지화된 응답. 호출별(`lang="ko-KR"`), 요청별(`Accept-Language`), 배포별
(`EDIT2DOCS_DEFAULT_LANG=ko-KR`)로 전환합니다. zh-CN / zh-TW / ja-JP도 같은
문자 감지와 폰트 스택 처리를 받습니다.

## 생태계

| 저장소 | 설명 |
|---|---|
| [edit2docs-web](https://github.com/CocoRoF/edit2docs-web) | 호스팅 서비스용 웹 스튜디오 (Next.js) |
| [contextifier](https://github.com/CocoRoF/Contextifier) | 무손실 편집에 쓰는 raw OOXML 레이어 |
| [ppt-master](https://github.com/hugohe3/ppt-master) | `src/edit2docs/core/` PPTX 코어의 upstream (MIT) |
| [edit2ppt](https://github.com/CocoRoF/edit2ppt) | 자매 프로젝트; 덱 파이프라인의 출처 |

## 프로젝트 구조

`src/edit2docs/`: `simple.py`(라이브러리 동사) · `agent_tools.py`,
`agent_guide.py`, `tool_matrix.py`(에이전트 표면) · `documents/`(docx/xlsx/pptx
엔진, arrange, 차트·XML 편집) · `tools/`(LLM 파이프라인: 생성, 편집, 프리뷰,
내보내기) · `core/`(ppt-master 파생 덱 엔진) · `mcp/`(로컬 stdio + 호스팅 서버) ·
`api/`(FastAPI) · `render/`, `i18n/`, `llm/`, `db/`, `storage/`, `workers/`,
`services/`. 설계 노트는 [docs/](./docs/)와
[PPTX_EDITING_UPGRADE_PLAN.md](./PPTX_EDITING_UPGRADE_PLAN.md)에 있습니다.

## 개발

```bash
git clone https://github.com/CocoRoF/edit2docs && cd edit2docs
uv venv .venv && uv pip install -e ".[server,dev]"
.venv/bin/python -m pytest tests/          # 0.19.0 기준 949 passed, 1 skipped
.venv/bin/python -m ruff check src/edit2docs
```

`scripts/lint_ascii_paths.py`가 ASCII 전용 경로를 강제합니다(테스트에서 실행).

## 버전 이력

저장소 현재 버전: **0.19.0**. 2026-10-01 기준 PyPI에는 0.16.1까지 있고,
0.17.0~0.19.0은 `main`에만 있습니다.

| 버전 | 요약 |
|---|---|
| 0.19.0 | 레이아웃 인식 네이티브 `add_slide`; 연산 `set_notes`, `set_bullet`, `set_hyperlink`, `set_z_order`, `set_legend`, `set_series_color`, `set_theme_color`, `set_theme_font` (PPTX 연산 25개); `contextifier>=0.8.0` 필요 |
| 0.18.0 | 자기 서술형 PPTX 엔진: `OP_CATALOG` / `describe_ops()`, `dry_run`, `atomic`, `find_shapes`; 연산 `set_runs`, `add_textbox`, `delete_shape`, `duplicate_shape`, `insert_column`, `delete_column`, `merge_cells` |
| 0.17.x | 구조적 PPTX 편집: 표적 턴에서 슬라이드 전체 SVG 재생성 대신 주소화된 바이트 보존 제자리 편집 |
| 0.16.x | 확장자 한정 도구(`tool_specs(extension=...)`, `tool_matrix`); `GET /v1/models`; 슬라이드 편집 재시도·실패 작업 처리 보강 |
| 0.15.1 | `arrange_doc` (0.15.0은 구현이 빠진 채 배포됨; 0.15.1 이상 사용) |
| 0.14.1 | `mcp<2.0` 상한 (2.0이 `mcp.server.fastmcp`를 제거) |
| 0.10-0.14 | `build_doc`, `read_doc_xml` / `set_doc_xml`, 계층형 `doc_guide`, 테마 덱 |
| 0.9.0 | 토큰 최적화: 프롬프트 캐시 재구성, 입력 상한, 역할별 모델 티어링 |
| 0.8.0 | contextifier 기반 무손실 편집; `edit_chart` |
| 0.7.x | ppt-master v3.1 동기화, 네이티브 차트/표 내보내기, English-first, Apache-2.0 전환(0.7.1) |
| 0.1-0.6 | 멀티포맷 엔진, 호스팅 API, 실시간 편집 스트리밍, 주소화 프리뷰, `render_doc` |

## 라이선스

[Apache License 2.0](./LICENSE). `src/edit2docs/core/`의 PPTX 코어는
[ppt-master](https://github.com/hugohe3/ppt-master)(MIT, Copyright (c) Hugo He)에서
파생되었으며, 해당 MIT 조건은 [NOTICE](./NOTICE)와
[LICENSE.ppt-master.MIT](./LICENSE.ppt-master.MIT)에 보존되어 있습니다.
