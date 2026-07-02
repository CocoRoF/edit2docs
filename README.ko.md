# edit2docs

**AI Agent 네이티브 문서 엔진 — DOCX · XLSX · PPTX. 한국어 우선.**

[English README](./README.md) · [edit2ppt](https://github.com/CocoRoF/edit2ppt)의
자매 프로젝트 (PPTX 파이프라인 상속) · [ppt-master](https://github.com/hugohe3/ppt-master) (MIT) 기반

---

한 줄 의도로 완성된 오피스 문서를 생성하고, 기존 파일을 채팅으로 편집합니다 —
Word 보고서, Excel 워크북, PowerPoint 덱 전부 네이티브 편집 가능한 OOXML로.
하나의 엔진, 네 가지 표면:

```bash
pip install edit2docs              # 라이브러리 + 에이전트 도구 + 로컬 MCP
pip install "edit2docs[server]"    # + 호스팅 서비스
```

## 다섯 가지 동사 (모든 표면 공통, 확장자로 엔진 선택)

| 동사 | 기능 | LLM |
|---|---|---|
| `generate_doc` | 의도 (+소스/PPTX 템플릿) → 완성 문서 | ✳ |
| `edit_doc` | 자연어 편집 1턴; 안 건드린 내용은 바이트 동일 유지 | ✳ |
| `preview_doc` | .pptx → 슬라이드별 SVG · .docx/.xlsx → 마크다운 | — |
| `analyze_doc` | 구조 아웃라인 + `set_doc_text`용 정확한 주소 | — |
| `set_doc_text` | 결정론적 정밀 편집 (문단/셀/슬라이드 텍스트) | — |

```python
from edit2docs import generate_doc, edit_doc, analyze_doc, set_doc_text

generate_doc("3분기 실적 보고서", output="report.docx")
generate_doc("분기별 매출 정리", output="sales.xlsx", sources=["raw.pdf"])
r = edit_doc("report.docx", "진행 사항 섹션에 배포 완료 항목 추가해줘")
set_doc_text("sales.xlsx", [{"sheet": "매출", "cell": "B3", "value": 142}])
```

- **에이전트 도구**: `from edit2docs.agent_tools import ANTHROPIC_TOOLS, run_tool`
- **로컬 MCP**: `edit2docs-mcp` (stdio) — Claude Desktop/Code/Cursor 설정 한 줄
- **호스팅**: `edit2docs serve` (edit2ppt에서 상속한 FastAPI 서비스)

BYOK: `api_key=` 또는 `ANTHROPIC_API_KEY`. preview/analyze/set_text는 키 불필요.

## 포맷별 동작

- **DOCX** — 작성 LLM이 제약된 마크다운을 쓰고 결정론적 렌더러(python-docx)가
  Word로 변환. 편집은 문단 주소 기반 연산(replace/insert_after/delete, 표 셀은
  row/col)이라 안 건드린 문단의 서식이 보존됩니다.
- **XLSX** — 설계 LLM이 YAML 시트 스펙(시트/헤더/행/숫자서식/수식)을 쓰고
  openpyxl이 스타일된 워크북 렌더. 편집은 set_cell/append_rows/add_sheet + stale 가드.
- **PPTX** — edit2ppt의 전체 파이프라인 상속: 전략가 → 페이지별 SVG → 네이티브
  DrawingML, 사용자 PPTX 템플릿(restyle/extend), recompose 채팅편집, 표 셀 포함
  문단 단위 텍스트 편집.

모든 LLM 플래너는 동일 계약: `reply`+`edit_plan` 펜스 블록, 포맷 리마인더 1회
재시도, 실패 시 침묵 대신 정직한 답변. 라이선스: MIT.
