# PPTX 편집 초고도화 설계안 (Structural PPTX Editing)

작성: 2026-08-06 · 대상: `edit2docs` PPTX 채팅 편집 파이프라인

---

## 0. 한 줄 요약

현재 PPTX 편집은 **슬라이드 전체를 SVG로 재생성**하는 손실적·단일체 방식이라 표·차트·헤더·스타일·레이아웃을 구조적으로 잃는다. 이미 DOCX/XLSX가 쓰고 있는 **"주소 기반 정밀 연산 + 바이트 보존 적용"** 모델을 PPTX로 이식한다. 핵심 프리미티브 상당수는 **이미 contextifier raw 레이어에 존재**하며, PPTX 도구로 **배선만 하면 된다**.

---

## 1. 근본 원인 (코드 근거)

`tools/edit_deck.py`의 한 턴은 4단계다:

1. `render_preview` → 모든 슬라이드를 **flat SVG**로 렌더
2. `editor-planner` → 텍스트 아웃라인으로 슬라이드 op(`edit`/`add`/`delete`) 계획
3. `editor-slide` → op마다 **슬라이드 전체 SVG를 재생성**
4. `recompose_pptx` → 패키지 재조립

이 구조가 사용자가 지적한 모든 증상의 원인이다:

| 증상 | 근본 원인 (코드) |
|---|---|
| **표 편집 불가·평면화** | 렌더 시 `tbl_to_svg.convert_tbl`이 `<a:tbl>`을 `<rect>`+`<line>`+`<text>`로 분해(구조 소실). 역변환 없음. `preserve_native`가 **원본 표 프레임을 deepcopy로 그대로 재삽입**(`pptx_edit._carry_native_frames:385`) → 모델이 셀을 바꿔도 **버려짐** |
| **차트 편집 불가** | `chart_to_svg`는 "preview fidelity only" — 시리즈/데이터 소실. 역변환 없음. `preserve_native`가 원본 차트 파트+워크북을 그대로 재삽입 → 데이터 편집 불가 |
| **스타일 헤더 삭제** | `editor-slide` 프롬프트가 **전체 슬라이드 재현**을 요구("anything you omit disappears"). 테마색·그라디언트는 렌더 시 리터럴 hex로 baking되어 테마 링크 소실. 모델이 빠뜨린 그룹/헤더는 그대로 사라짐 |
| **레이아웃 미충족·위치 오류** | brief에 좌표 없음 → 모델이 좌표를 **재발명**. `convert_svg_to_slide_shapes`는 SVG x/y를 EMU로 **그대로** 매핑(레이아웃 엔진 없음). 플레이스홀더 미활용 |

정리: **손실 렌더 → 전체 재작성 → 손실 역변환**의 3중 손실 파이프라인. native 보존은 "얼려서 편집 불가"라는 부작용을 낳는다.

---

## 2. 올바른 방향 — DOCX/XLSX 모델의 PPTX 이식

DOCX/XLSX는 이미 정답 구조다:

- **주소가 달린 요소**(문단 N, 표 셀 (r,c), A1 셀)에 대한 **수술적 연산**을
- **contextifier raw 레이어**로 **바이트 보존** 적용(건드린 파트만 재직렬화, 나머지 원본 그대로)
- 모델에는 **경량 텍스트 아웃라인**(주소 포함, 윈도잉)만 제공 → 렌더 왕복 없음

PPTX도 동일하게: **슬라이드를 주소 가능한 구조로 표현 → 정밀 op을 바이트 보존 적용**. SVG 재작성은 "자유 재디자인" 전용 opt-in 폴백으로 강등.

### 결정적 사실: 프리미티브가 이미 있다

contextifier raw 레이어에 **이미 존재하지만 PPTX 도구로 미배선**:

| 프리미티브 | 위치 | 상태 |
|---|---|---|
| `RawSlide.set_text(shape_id, text, para)` (서식 보존) | `contextifier/raw/pptx.py:332` | ✅ 존재, DOCX가 유사 사용 |
| `RawTable.cell(r,c).set_text()` / `insert_row` / `delete_row` | `pptx.py:164-268` | ✅ 존재, DOCX 표에서 검증됨 |
| `ChartModel.set_data(categories, series)` / `set_title()` (워크북 동기화 포함) | `chart.py:416,484` | ✅ 존재, **`edit_chart` verb로 이미 노출** |
| 슬라이드 구조 duplicate/move/delete | `arrange.py` → raw | ✅ 존재, `arrange_doc`로 노출 |
| 주소 체계 `data-e2p-shape/table/cell/para` (cNvPr id) | `slide_to_svg.py:1062`, `tbl_to_svg.py:193`, `txbody_to_svg.py:1136` | ✅ 프리뷰에 이미 찍힘 |
| shape 인벤토리 `RawShapeInfo(id,name,kind,text)`, `slide.tables`, `slide.charts` | `pptx.py:302-371` | ✅ 존재(읽기) |

**미존재(3단계에서 contextifier에 신규 구현):** shape fill/font/position(`a:xfrm`) 뮤테이터, 표 셀 스타일, 차트 라벨/범례/축/색, 플레이스홀더 열거 + 레이아웃 기반 삽입.

---

## 3. 단계별 실행 계획

### Stage 0 — 주소 가능한 슬라이드 모델 (읽기)
- `documents/pptx_engine.py`에 `pptx_outline(pptx)` 신설 (docx_outline 대응). 슬라이드마다:
  - shapes: `{shape_id, kind, name, text, paras[], pos(l,t,w,h), style요약}`
  - tables: `{shape_id, n_rows, n_cols, cells[r][c]}`
  - charts: `{shape_id, kind, title, categories, series}`
  - placeholders: `{idx, type, name}`
- contextifier `RawSlide`의 기존 read API 활용(신규 코드 없음).
- 아웃라인 라인 포맷: `- shape 5 [title]: "..."`, `- table 7 cell (0,1): "..."`, `- chart 9 [bar] "매출": cats=[...] series=[...]` — 계획 op이 이 주소로 해석.

### Stage 1 — 수술적 op 엔진 (기존 프리미티브 배선, contextifier 무변경) ⚑ 최우선·독립 배포 가능
- `documents/pptx_engine.py`에 `PptxEdit` 스키마 + `apply_pptx_edits`(docx_engine 미러). 모두 raw 바이트 보존.
- op(전부 기존 프리미티브):
  - `set_shape_text {slide, shape_id, para, new_text, old_text?}` → `RawSlide.set_text`
  - `set_table_cell {slide, shape_id, row, col, new_text, old_text?}` → `RawTable.cell().set_text`
  - `table_insert_row` / `table_delete_row`
  - `set_chart_data {slide, shape_id, categories, series}` / `set_chart_title` → `ChartModel.set_data/set_title`
  - 슬라이드 구조: `arrange_doc` 재사용(duplicate/move/delete)
- 낙관적 동시성(`old_text` 가드), 결과 상태 `applied|stale|not_found|invalid` (docx 동일).
- **효과: 표·차트·텍스트가 제자리에서 편집되고 나머지는 전부 보존** → 불만의 다수를 이 단계에서 해소.

### Stage 2 — PPTX 정밀 planner + edit_deck 라우팅
- 신규 `pptx-editor-planner` 프롬프트: 아웃라인 대상 **주소 op**을 emit(프로즈 brief 아님). 모델 컨텍스트 = 경량 텍스트 아웃라인(윈도잉), SVG 왕복 제거.
- `edit_deck` 라우팅: 기본은 수술적 엔진. **SVG 재생성은 명시적 `redesign` op**(자유 창작 재배치)로만, 그때도 `preserve_native`로 표/차트 보존.
- 프론트: `data-e2p-*` 주소로 편집 대상 하이라이트(docx `data-e2d-*` 방식 재사용).

### Stage 3 — contextifier 프리미티브 확장 (의존성 신규 코드)
- `RawSlide`: `set_shape_fill/line/font`, `set_shape_position`(a:xfrm off/ext) — shape_id 주소.
- `RawTableCell`: fill/border/font/valign 스타일.
- `ChartModel`: data-label/legend/axis/series-color, chartEx 쓰기.
- **플레이스홀더 열거 + 레이아웃 기반 삽입**: `add_slide_from_layout`, 플레이스홀더 P 채우기 → "내용이 레이아웃을 못 채움" 근본 해결.
- pptx_engine 신규 op: `set_shape_style`, `set_shape_position`, `set_cell_style`, `add_shape`, `add_slide`(layout-aware).
- contextifier 릴리스 → edit2docs 반영(연쇄 릴리스 절차 준수).

### Stage 4 — 레이아웃 지능 + 검증
- 오토핏/리플로우: 텍스트 증가 시 텍스트박스 autofit, 콘텐츠 추가 시 플레이스홀더 경계 존중.
- **골든 덱 테스트**: 표·차트·스타일 헤더 포함 덱으로 (a) 셀/시리즈 편집 동작, (b) 미접촉 요소 바이트 보존, (c) 헤더/테마 생존, (d) 위치가 레이아웃 존중 — 회귀 고정.

---

## 4. 도구 철학 (사용자 요구: "경량·정밀 도구로 Agent가 정확히 판단")

- 소수의 **타입 있고 주소 달린 verb**만 노출. 에이전트는 **압축된 주소 아웃라인**(KB 단위)을 보고 **정밀 op**을 emit. 전체 재작성 없음.
- 각 op은 **결정론적·바이트 보존·되돌림 가능**. 확장자 기반 tool scoping(이미 구현됨)과 정합.
- 손실 SVG 왕복은 기본 경로에서 제거, 자유 재디자인 전용으로 격리.

---

## 5. 리스크·순서

- **Stage 1은 자립적**(contextifier 무변경, 기존 검증된 프리미티브) → 먼저 단독 배포해 즉시 가치. 낮은 리스크.
- Stage 3은 contextifier 수정 → 연쇄 릴리스 필요(더 큼). Stage 1~2 안정화 후 진행.
- 하위호환: 기존 `edit`(SVG) 경로는 `redesign`으로 보존, 제거 아님.

---

## ✅ 구현 상태 (2026-08-06)

전 구간 구현 완료 — **배포 전 검토 대기**. 두 repo에 브랜치로 커밋.

**contextifier** (raw 레이어 확장):
- `RawShapeInfo`에 `left/top/width/height`(EMU) 기하 읽기 추가.
- `RawSlide.set_shape_font` / `set_shape_fill` / `set_shape_position`(a:xfrm).
- `RawTableCell.set_style`(fill/font). 전부 바이트 보존·서식 보존. 단위 테스트 +8 (627 통과).

**edit2docs**:
- `documents/pptx_engine.py` — `pptx_outline`(주소 아웃라인, 위치 inch 포함) + `apply_pptx_edits`.
  ops: `set_text` · `set_shape_style` · `set_shape_position` · `set_table_cell` ·
  `set_cell_style` · `insert_row` · `delete_row` · `set_chart_data` · `set_chart_title`.
- `tools/edit_doc.py` — pptx를 수술 경로로 확장(아웃라인·apply·planner role·`needs_svg` 신호).
- `core/prompts/pptx-editor-planner.en.md` — 주소 op planner.
- `workers/executors/edit_deck.py` — 전 포맷 수술 경로 우선, pptx는 `needs_svg`일 때만 SVG 폴백.
- 프론트(`edit2docs-web`): `data-e2p-*` 라이브 하이라이트(SlideCanvas), op 라벨 확장.
- 테스트: `test_pptx_engine`(15) + `test_pptx_surgical_edit`(4). 전체 923 통과.

**해소된 사용자 불만**:
- 표 편집 불가 → ✅ `set_table_cell`/`insert_row`/`delete_row`/`set_cell_style` (실제 표 유지).
- 차트 편집 불가 → ✅ `set_chart_data`/`set_chart_title` (실제 차트 유지).
- 스타일 헤더 삭제 → ✅ 수술 경로는 슬라이드를 재작성하지 않음 → 미접촉 요소 바이트 보존(골든 테스트로 고정).
- 위치/스타일 → ✅ `set_shape_position`(아웃라인이 현재 위치 inch 노출) + `set_shape_style`.

**보류(다음 반복)**: 차트 축/범례/라벨·시리즈 색; chartEx 쓰기; 플레이스홀더 열거 + OPC 레이아웃 기반 신규 슬라이드 삽입(현재 신규 슬라이드는 SVG 경로가 처리); 텍스트 증가 시 오토핏/리플로우.

**배포 순서(검토·승인 후)**: contextifier 릴리스(신규 raw 메서드 의존) → edit2docs main 머지 → hr-web no-cache 재빌드(edit2docs-server + web).

## ✅ 2차 강화 (2026-08-06) — 프레임워크 인터페이스 + author-in-place

두 서브에이전트(적대적 재검토 + 격차분석) 결과를 반영해 **정확성 하드닝 + 강력한 프레임워크 인터페이스 + 저작 능력 확장**을 완료.

**정확성 수정(적대적 R2)**: 버블차트 set_data 손상 차단(`RawUnsupportedError`), 비유한/모호(inf·nan·"1,5") 차트값 거부, bool 주소 강제변환 차단, O(shapes²)→단일 walk(`paragraphs_by_shape`), 윈도잉 콘텐츠앵커(따옴표 텍스트 매칭), 생성분기 op캡, `_as_num` 안전화, None 카테고리→"", 빈 placeholder set_text가 a:p 생성.

**프레임워크 인터페이스(자기기술)**: `OP_CATALOG`(단일 진실원, 16 ops의 address/payload 계약) + `describe_ops()` 내성 + **`dry_run`**(무바이트 검증) + **`atomic`**(전부-or-무 롤백) + **`find_shapes`**(text/kind/name/slide 쿼리). VALID_PPTX_ACTIONS는 카탈로그에서 파생.

**신규 저작 op**: `set_runs`(문단 내 다중 서식 런 — "한 단어만 굵게"), `add_textbox`·`delete_shape`·`duplicate_shape`(shape 생명주기), `insert_column`·`delete_column`·`merge_cells`(표 완성). 전부 바이트 보존. contextifier에 대응 raw 프리미티브 추가.

## ✅ Phase 2 완료 (2026-08-06) — 네이티브 저작 + 풀서피스 op (25종)

- **네이티브 add_slide**(레이아웃 사용·플레이스홀더 채움 → "내용이 레이아웃 채움" 완전 해결; SVG 폴백 제거) + set_paragraphs.
- **set_notes**(노트슬라이드 파트 생성 포함), **set_bullet**, **set_hyperlink**, **set_z_order**.
- **차트 심화**: set_legend, set_series_color.
- **테마**: set_theme_color, set_theme_font(덱 전체).
- 전부 바이트 보존(contextifier 0.8.0 raw 프리미티브). 회귀 edit2docs 949·contextifier 640.

### 🔜 Phase 3 (남은 것)
이미지 추가/교체/크롭, 차트 축제목·데이터라벨·타입변경, 표 셀 분할·열너비/행높이, 그룹/언그룹, 구조적 op의 1급 직접호출 verb 노출.

## 🔜 (구) Phase 2 로드맵

SVG 폴백을 완전 제거하기 위한 잔여 능력 (격차분석 우선순위):
- **네이티브 레이아웃 슬라이드 삽입 + 플레이스홀더 채우기** (L, 최고가치 — "내용이 레이아웃 못 채움"의 완전 해결). contextifier에 슬라이드레이아웃/마스터/`p:ph` 리더 + `add_slide(layout)` 필요.
- **스피커 노트 쓰기** (notesSlide 파트 생성/rel/content-type).
- **하이퍼링크**(런/도형, `a:hlinkClick`+rel), **불릿/번호**(`a:buChar`/`a:buAutoNum`).
- **차트 심화**: 타입변경·축제목·범례·데이터라벨·시리즈색; chartEx 쓰기.
- **이미지**: 추가/교체/크롭(`p:pic`/`blipFill`/media 파트).
- **z-order/그룹·언그룹**, **테마/마스터**(clrScheme·fontScheme 스왑).
- **인터페이스**: 구조적 op을 1급 직접호출 verb로 노출(챗 planner는 한 클라이언트).

## 6. 착수 지점 제안

Stage 0+1을 하나의 자립 PR로: `pptx_outline` + `pptx_engine`(set_shape_text / set_table_cell / table_insert_row/delete_row / set_chart_data/title) + 최소 planner. 표·차트·텍스트 in-place 편집 + 전체 보존을 먼저 실증한다.
