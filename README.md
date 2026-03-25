# Univer 0.16.0 Preset 모드 공식 API 분석 가이드

> **공식 자료 기반** | 버전: `@univerjs/presets@0.16.0` (2026-02-28 릴리스)
>
> **참고 문서:**
> - [공식 설치 가이드](https://docs.univer.ai/guides/docs/getting-started/installation)
> - [Facade API 가이드](https://docs.univer.ai/en-US/guides/sheets/getting-started/facade)
> - [Core Features](https://docs.univer.ai/en-US/guides/sheets/features/core)
> - [Range & Selection & Cell](https://docs.univer.ai/guides/sheets/features/core/range-selection)
> - [Facade API Reference](https://reference.univer.ai/en-US)
> - [v0.16.0 Presets Release](https://github.com/dream-num/univer-presets/releases/tag/v0.16.0)
> - [v0.16.1 Core Release](https://github.com/dream-num/univer/releases/tag/v0.16.1)

---

## 목차

1. [버전 개요](#1-버전-개요)
2. [Preset 모드 vs Plugin 모드](#2-preset-모드-vs-plugin-모드)
3. [설치 및 기본 설정](#3-설치-및-기본-설정)
4. [createUniver 설정 옵션](#4-createuniver-설정-옵션)
5. [사용 가능한 Preset 패키지](#5-사용-가능한-preset-패키지)
6. [Facade API 핵심 구조](#6-facade-api-핵심-구조)
7. [Workbook 조작 API](#7-workbook-조작-api)
8. [Worksheet 조작 API](#8-worksheet-조작-api)
9. [Range (셀/범위) 조작 API](#9-range-셀범위-조작-api)
10. [Selection (선택 영역) API](#10-selection-선택-영역-api)
11. [이벤트 시스템](#11-이벤트-시스템)
12. [UI 커스터마이징](#12-ui-커스터마이징)
13. [라이프사이클 관리](#13-라이프사이클-관리)
14. [v0.16.x 주요 변경사항 및 Breaking Changes](#14-v016x-주요-변경사항-및-breaking-changes)
15. [실전 코드 예제 모음](#15-실전-코드-예제-모음)

---

## 1. 버전 개요

### Univer 0.16.0 릴리스 정보

| 항목 | 내용 |
|------|------|
| **릴리스 날짜** | 2026-02-28 |
| **Presets 패키지** | `@univerjs/presets@0.16.0` |
| **Core SDK** | `@univerjs/core@0.16.0` |
| **라이선스** | Apache-2.0 |

### v0.16.x 주요 신규 기능

- **Shape 삽입 지원 (Beta)**: `@univerjs-pro/sheets-shape`, `@univerjs-pro/sheets-shape-ui` 패키지 추가
- **History 기능 리팩토링**: 더 풍부한 히스토리 정보 표시 지원
- **이미지 플립 지원**: 플로팅 이미지의 좌우/상하 반전 기능
- **렌더링 성능 최적화**: 대량 병합 셀 + 숨겨진 행이 있는 워크시트 성능 개선
- **컨텍스트 메뉴 리팩토링**: 성능 최적화 및 기존 이슈 해결

---

## 2. Preset 모드 vs Plugin 모드

Univer는 두 가지 통합 방식을 제공합니다:

| 항목 | Plugin 모드 | Preset 모드 |
|------|------------|------------|
| Facade API 패키지 | 수동 import 필요 | 자동 포함 |
| 플러그인 등록 순서 | 직접 관리 필요 | 자동 처리 |
| 지연 로딩 | 플러그인 단위 지연 로딩 | Preset 단위 지연 로딩만 지원 |
| 복잡도 | 높음 (세밀한 제어) | 낮음 (빠른 시작) |

**Preset 모드의 핵심 장점:**
- `createUniver()` 호출 시 `univerAPI` (Facade API 인스턴스)가 자동 반환
- 별도의 Facade 패키지 import 불필요
- 기능별 Preset을 배열로 전달하여 간편하게 구성

---

## 3. 설치 및 기본 설정

### 패키지 설치

```bash
npm install @univerjs/presets@0.16.0 @univerjs/preset-sheets-core@0.16.0
```

### 최소 구성 코드 (Sheets)

```javascript
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US';
import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets';

import '@univerjs/preset-sheets-core/lib/index.css';

const { univerAPI } = createUniver({
  locale: LocaleType.EN_US,
  locales: {
    [LocaleType.EN_US]: mergeLocales(
      UniverPresetSheetsCoreEnUS,
    ),
  },
  presets: [
    UniverSheetsCorePreset({
      container: 'app',
    }),
  ],
});

univerAPI.createWorkbook({});
```

### HTML

```html
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Univer Sheets</title>
  <style>
    #app {
      width: 100vw;
      height: 100vh;
    }
  </style>
</head>
<body>
  <div id="app"></div>
  <script type="module" src="./index.js"></script>
</body>
</html>
```

---

## 4. createUniver 설정 옵션

### CreateUniverParameter 인터페이스

```javascript
const { univer, univerAPI } = createUniver({
  // 테마 설정 (선택)
  theme: defaultTheme,

  // 다크 모드 활성화 (선택, 기본값: false)
  darkMode: false,

  // 로케일 설정 (필수)
  locale: LocaleType.EN_US,

  // 로케일 객체 (필수)
  locales: {
    [LocaleType.EN_US]: mergeLocales(/* 로케일 패키지들 */),
  },

  // 로그 레벨 (선택)
  logLevel: LogLevel.WARN,

  // Preset 목록 (필수)
  presets: [
    UniverSheetsCorePreset({ container: 'app' }),
  ],

  // 추가 플러그인 (선택)
  plugins: [],
});
```

### UniverSheetsCorePreset 설정 옵션

```javascript
UniverSheetsCorePreset({
  // DOM 컨테이너 (string 셀렉터 또는 HTMLElement)
  container: 'app',

  // 헤더 표시 여부
  header: true,

  // 툴바 표시 여부
  toolbar: true,

  // 리본 타입: 'classic' | 'collapsed' | 'simple'
  // v0.17.0부터 'default'가 'collapsed'로 변경됨
  ribbonType: 'classic',

  // 푸터 설정
  footer: {
    sheetBar: true,      // 시트 탭 바
    statisticBar: true,   // 통계 바
    menus: true,          // 하단 메뉴
    zoomSlider: true,     // 줌 슬라이더
  },

  // 컨텍스트 메뉴 표시 여부
  contextMenu: true,

  // 자동 포커스 비활성화
  disableAutoFocus: false,

  // 메뉴 항목 설정
  menu: {
    'sheet.command.set-range-bold': { hidden: false },
  },

  // 시트 관련 설정
  sheets: {
    isRowStylePrecedeColumnStyle: false,
    autoHeightForMergedCells: false,
    freezeSync: true,
    clipboardConfig: {
      hidePasteOptions: false,
    },
    protectedRangeShadow: true,
    disableForceStringAlert: false,
    disableForceStringMark: false,
  },
});
```

---

## 5. 사용 가능한 Preset 패키지

### 핵심 Preset

| 패키지 | 설명 | 설치 |
|--------|------|------|
| `@univerjs/preset-sheets-core` | Sheets 코어 (편집, 수식, 렌더링) | 필수 |

### 기능별 Preset

| 패키지 | 설명 |
|--------|------|
| `@univerjs/preset-sheets-filter` | 필터 기능 |
| `@univerjs/preset-sheets-sort` | 정렬 기능 |
| `@univerjs/preset-sheets-conditional-formatting` | 조건부 서식 |
| `@univerjs/preset-sheets-data-validation` | 데이터 유효성 검사 |
| `@univerjs/preset-sheets-drawing` | 드로잉/이미지 |
| `@univerjs/preset-sheets-hyper-link` | 하이퍼링크 |
| `@univerjs/preset-sheets-find-replace` | 찾기/바꾸기 |
| `@univerjs/preset-sheets-thread-comment` | 스레드 댓글 |

### 풀 기능 구성 예제

```javascript
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import sheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US';
import { UniverSheetsFilterPreset } from '@univerjs/preset-sheets-filter';
import sheetsFilterEnUS from '@univerjs/preset-sheets-filter/locales/en-US';
import { UniverSheetsSortPreset } from '@univerjs/preset-sheets-sort';
import sheetsSortEnUS from '@univerjs/preset-sheets-sort/locales/en-US';
import { UniverSheetsConditionalFormattingPreset } from '@univerjs/preset-sheets-conditional-formatting';
import sheetsConditionalFormattingEnUS from '@univerjs/preset-sheets-conditional-formatting/locales/en-US';
import { UniverSheetsDataValidationPreset } from '@univerjs/preset-sheets-data-validation';
import sheetsDataValidationEnUS from '@univerjs/preset-sheets-data-validation/locales/en-US';
import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing';
import sheetsDrawingEnUS from '@univerjs/preset-sheets-drawing/locales/en-US';
import { UniverSheetsHyperLinkPreset } from '@univerjs/preset-sheets-hyper-link';
import sheetsHyperLinkEnUS from '@univerjs/preset-sheets-hyper-link/locales/en-US';
import { UniverSheetsFindReplacePreset } from '@univerjs/preset-sheets-find-replace';
import sheetsFindReplaceEnUS from '@univerjs/preset-sheets-find-replace/locales/en-US';
import { UniverSheetsThreadCommentPreset } from '@univerjs/preset-sheets-thread-comment';
import sheetsThreadCommentEnUS from '@univerjs/preset-sheets-thread-comment/locales/en-US';
import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets';

import '@univerjs/preset-sheets-core/lib/index.css';
import '@univerjs/preset-sheets-filter/lib/index.css';
import '@univerjs/preset-sheets-sort/lib/index.css';
import '@univerjs/preset-sheets-conditional-formatting/lib/index.css';
import '@univerjs/preset-sheets-data-validation/lib/index.css';
import '@univerjs/preset-sheets-drawing/lib/index.css';
import '@univerjs/preset-sheets-hyper-link/lib/index.css';
import '@univerjs/preset-sheets-find-replace/lib/index.css';
import '@univerjs/preset-sheets-thread-comment/lib/index.css';

const { univerAPI } = createUniver({
  locale: LocaleType.EN_US,
  locales: {
    [LocaleType.EN_US]: mergeLocales(
      sheetsCoreEnUS,
      sheetsFilterEnUS,
      sheetsSortEnUS,
      sheetsConditionalFormattingEnUS,
      sheetsDataValidationEnUS,
      sheetsDrawingEnUS,
      sheetsHyperLinkEnUS,
      sheetsFindReplaceEnUS,
      sheetsThreadCommentEnUS,
    ),
  },
  presets: [
    UniverSheetsCorePreset({ container: 'app' }),
    UniverSheetsFilterPreset(),
    UniverSheetsSortPreset(),
    UniverSheetsConditionalFormattingPreset(),
    UniverSheetsDataValidationPreset(),
    UniverSheetsDrawingPreset(),
    UniverSheetsHyperLinkPreset(),
    UniverSheetsFindReplacePreset(),
    UniverSheetsThreadCommentPreset(),
  ],
});

univerAPI.createWorkbook({});
```

---

## 6. Facade API 핵심 구조

Preset 모드에서 `createUniver()` 호출 시 반환되는 `univerAPI`는 `FUniver` 인스턴스입니다.

### Facade API 계층 구조

```
univerAPI (FUniver)
├── getActiveWorkbook()    → FWorkbook
│   ├── getActiveSheet()   → FWorksheet
│   │   ├── getRange()     → FRange
│   │   └── getSelection() → FSelection
│   └── create()           → FWorksheet (새 시트 생성)
├── Event                  → 이벤트 상수
├── Enum                   → 열거형 상수
└── addEvent()             → 이벤트 리스너 등록
```

### 비동기 API 주의사항

데이터를 수정하는 Facade API는 대부분 **비동기(Promise)**입니다. 수정 직후 데이터를 조회하려면 반드시 `await` 또는 `.then()`을 사용해야 합니다.

```javascript
// 올바른 사용법
await fRange.setValue('Hello');
const value = fRange.getValue(); // 'Hello'

// 잘못된 사용법 (값이 반영되지 않을 수 있음)
fRange.setValue('Hello');
const value = fRange.getValue(); // undefined 또는 이전 값
```

---

## 7. Workbook 조작 API

### Workbook 생성

```javascript
// 빈 Workbook 생성
univerAPI.createWorkbook({});

// 데이터가 포함된 Workbook 생성
univerAPI.createWorkbook({
  id: 'workbook-01',
  name: '내 스프레드시트',
  sheetOrder: ['sheet-01'],
  sheets: {
    'sheet-01': {
      id: 'sheet-01',
      name: 'Sheet1',
      rowCount: 100,
      columnCount: 20,
      cellData: {
        0: {
          0: { v: '이름' },
          1: { v: '나이' },
          2: { v: '점수' },
        },
        1: {
          0: { v: '홍길동' },
          1: { v: 25 },
          2: { v: 95 },
        },
      },
    },
  },
});
```

### Workbook 가져오기 및 조작

```javascript
// 현재 활성 Workbook
const fWorkbook = univerAPI.getActiveWorkbook();

// 편집 시작
fWorkbook.startEditing();

// 편집 종료 (true: 변경 커밋, false: 취소)
await fWorkbook.endEditingAsync(true);

// 선택 비활성화/활성화
fWorkbook.disableSelection();
fWorkbook.enableSelection();

// 선택 숨기기/표시
fWorkbook.transparentSelection();
fWorkbook.showSelection();
```

---

## 8. Worksheet 조작 API

### Worksheet 접근

```javascript
const fWorkbook = univerAPI.getActiveWorkbook();

// 활성 시트 가져오기
const fWorksheet = fWorkbook.getActiveSheet();

// 새 시트 생성
const newSheet = fWorkbook.create('새 시트', 100, 20);

// 활성 시트 삭제
fWorkbook.deleteActiveSheet();
```

### Worksheet 주요 메서드

```javascript
const fWorksheet = fWorkbook.getActiveSheet();

// 범위 가져오기
const range = fWorksheet.getRange(0, 0, 2, 2);     // (row, col, rows, cols)
const rangeA1 = fWorksheet.getRange('A1:B2');       // A1 표기법
const cell = fWorksheet.getRange('A1');             // 단일 셀
const col = fWorksheet.getRange('A:A');             // 전체 열
const row = fWorksheet.getRange('1:1');             // 전체 행

// 선택 영역 가져오기
const selection = fWorksheet.getSelection();

// 병합된 범위 가져오기
const mergedRanges = fWorksheet.getMergedRanges();

// 캔버스 새로고침
fWorksheet.refreshCanvas();
```

---

## 9. Range (셀/범위) 조작 API

### 범위 생성

```javascript
const fWorkbook = univerAPI.getActiveWorkbook();
const fWorksheet = fWorkbook.getActiveSheet();

// 다양한 방식으로 범위 생성
const rangeA1 = fWorksheet.getRange(0, 0);          // A1 셀
const rangeA1B2 = fWorksheet.getRange(0, 0, 2, 2);  // A1:B2
const rangeStr = fWorksheet.getRange('A1:B2');       // 문자열
const rangeSheet = fWorksheet.getRange('Sheet1!A1:B2'); // 시트 지정

console.log(rangeA1B2.getA1Notation()); // 'A1:B2'
```

### 값 읽기

```javascript
const fRange = fWorksheet.getRange('A1:B2');

// 단일 셀 값 (범위의 좌상단 셀)
console.log(fRange.getValue());         // 셀 값
console.log(fRange.getRawValue());      // 원시 값
console.log(fRange.getDisplayValue());  // 표시 값
console.log(fRange.getCellData());      // 셀 데이터 객체
console.log(fRange.getFormula());       // 수식

// 리치 텍스트
const richText = fRange.getRichTextValue();
console.log(richText.toPlainText());    // 일반 텍스트

// 범위 전체 값
console.log(fRange.getValues());        // 2D 배열
console.log(fRange.getRawValues());     // 원시 값 배열
console.log(fRange.getDisplayValues()); // 표시 값 배열
console.log(fRange.getCellDatas());     // 셀 데이터 객체 배열
console.log(fRange.getFormulas());      // 수식 배열
```

### 값 쓰기

```javascript
const fRange = fWorksheet.getRange('A1:B2');

// 단일 값 설정 (범위 전체에 적용)
fRange.setValue('Hello, Univer');

// 수식 설정 ('='로 시작)
fRange.setValue('=A1+B1');

// 셀 객체로 설정
fRange.setValue({
  v: 'Hello, Univer',
  custom: { key: 'value' },
});

// 좌상단 셀에만 값 설정
fRange.setValueForCell('Hello');

// 배열로 복수 값 설정 (범위 크기와 일치해야 함)
fRange.setValues([
  ['A1값', 'B1값'],
  ['A2값', 'B2값'],
]);

// 셀 객체 배열로 설정
fRange.setValues([
  [{ v: 'A1' }, { v: 'B1' }],
  [{ v: 'A2' }, { v: 'B2' }],
]);

// 객체로 설정 (행/열 인덱스 기반, 범위 크기 불일치 허용)
fRange.setValues({
  0: { 0: 'A1', 1: 'B1' },
  1: { 0: 'A2', 1: 'B2' },
});
```

### 값 삭제

```javascript
const fRange = fWorksheet.getRange('A1:D10');

// 내용 + 서식 모두 삭제
fRange.clear();

// 내용만 삭제
fRange.clearContent();

// 서식만 삭제
fRange.clearFormat();
```

### 스타일 조작

```javascript
const fRange = fWorksheet.getRange('A1:B2');

// 스타일 읽기
console.log(fRange.getCellStyleData());  // 좌상단 셀 스타일 데이터
console.log(fRange.getCellStyle());      // 좌상단 셀 스타일
console.log(fRange.getCellStyles());     // 전체 셀 스타일

// 스타일 설정 (체이닝 가능)
fRange
  .setValues([[1, 2], [3, 4]])
  .setFontWeight('bold')
  .setFontLine('underline')
  .setFontFamily('Arial')
  .setFontSize(24)
  .setFontColor('red');

// 스타일 초기화 (null 전달)
fRange
  .setFontWeight(null)
  .setFontLine(null)
  .setFontFamily(null)
  .setFontSize(null)
  .setFontColor(null);
```

### 셀 병합

```javascript
// 병합
const fRange = fWorksheet.getRange('B1:B2');
fRange.merge();
console.log(fRange.isMerged()); // true

// 병합 확인
const fRange2 = fWorksheet.getRange('A1:B2');
console.log(fRange2.isPartOfMerge()); // true

// 병합 해제
fRange2.breakApart();

// 가로 방향 병합
fRange2.mergeAcross();

// 세로 방향 병합
fRange2.mergeVertically();
```

### 셀 삽입/삭제

```javascript
// 빈 셀 삽입 (기존 데이터 우측으로 이동)
const fRange = fWorksheet.getRange('A1:B2');
fRange.insertCells(univerAPI.Enum.Dimension.COLUMNS);

// 빈 셀 삽입 (기존 데이터 하단으로 이동)
fRange.insertCells(univerAPI.Enum.Dimension.ROWS);

// 셀 삭제 (좌측 데이터 당겨오기)
fRange.deleteCells(univerAPI.Enum.Dimension.COLUMNS);

// 셀 삭제 (상단 데이터 당겨오기)
fRange.deleteCells(univerAPI.Enum.Dimension.ROWS);
```

### 범위 하이라이트

```javascript
// 기본 스타일로 하이라이트
const fRange = fWorksheet.getRange('C3:E5');
fRange.highlight();

// 커스텀 스타일 하이라이트
const fRange2 = fWorksheet.getRange('C7:E9');
const primaryCell = fWorksheet.getRange('D8').getRange();
const disposable = fRange2.highlight(
  { stroke: 'red', fill: 'yellow' },
  {
    ...primaryCell,
    actualRow: primaryCell.startRow,
    actualColumn: primaryCell.startColumn,
  },
);

// 5초 후 하이라이트 제거
setTimeout(() => disposable.dispose(), 5000);
```

### 텍스트 열 분할

```javascript
const fRange = fWorksheet.getRange('A1:A3');
fRange.setValues([
  ['이름,나이,점수'],
  ['홍길동,25,95'],
  ['김철수,30,88'],
]);

// 빈 셀 연속 구분자를 하나로 처리하여 분할
fRange.splitTextToColumns(true);

// 빈 셀 유지
fRange.splitTextToColumns(false);

// 커스텀 구분자
fRange.splitTextToColumns(
  false,
  univerAPI.Enum.SplitDelimiterType.Custom,
  '#'
);
```

### 범위 좌표 정보

```javascript
const fRange = fWorksheet.getRange('A1:B2');

// 좌표 정보 (width, height, left, right, top, bottom, x, y)
console.log(fRange.getCellRect());

// 병합 정보 + 좌표 동시 조회
console.log(fRange.getCell());
```

---

## 10. Selection (선택 영역) API

### 선택 영역 가져오기

```javascript
const fWorkbook = univerAPI.getActiveWorkbook();
const fWorksheet = fWorkbook.getActiveSheet();

// 범위 활성화 (선택)
const fRange = fWorksheet.getRange('A1:B2');
fRange.activate();

// 현재 선택 영역 가져오기
const fSelection = fWorksheet.getSelection();
const activeRange = fSelection.getActiveRange();
console.log(activeRange.getA1Notation()); // 'A1:B2'
```

### 선택 영역 설정

```javascript
// FRange로 설정
fWorksheet.getRange('A1:B2').activate();

// FWorksheet 메서드로 설정
fWorksheet.setActiveSelection(fWorksheet.getRange('C1:D2'));
```

### 현재 셀 가져오기

```javascript
const fRange = fWorksheet.getRange('A10:B11');
fRange.activate();

let fSelection = fWorksheet.getSelection();
const { actualRow, actualColumn } = fSelection.getCurrentCell();
console.log(
  fWorksheet.getRange(actualRow, actualColumn).getA1Notation()
); // 'A10'

// 활성 셀 변경
fSelection = fSelection.updatePrimaryCell(
  fWorksheet.getRange('B11')
);
```

---

## 11. 이벤트 시스템

### 이벤트 리스너 등록/해제

```javascript
// 이벤트 리스너 등록
const disposable = univerAPI.addEvent(
  univerAPI.Event.CellClicked,
  (params) => {
    const { worksheet, workbook, row, column } = params;
    console.log('클릭된 셀:', worksheet.getRange(row, column).getA1Notation());
  }
);

// 이벤트 리스너 해제
disposable.dispose();
```

### 셀 마우스 이벤트

| 이벤트 | 설명 |
|--------|------|
| `CellPointerMove` | 마우스가 셀 위로 이동 |
| `CellPointerDown` | 마우스 버튼 누름 |
| `CellPointerUp` | 마우스 버튼 놓음 |
| `CellHover` | 마우스 호버 |
| `CellClicked` | 셀 클릭 |
| `DragOver` | 드래그 중 셀 위를 지나감 |
| `Drop` | 셀에 드롭 |

```javascript
// 마우스 이벤트 예제
univerAPI.addEvent(univerAPI.Event.CellPointerMove, (params) => {
  const { worksheet, row, column } = params;
  console.log('현재 위치:', worksheet.getRange(row, column).getA1Notation());
});

univerAPI.addEvent(univerAPI.Event.CellClicked, (params) => {
  const { worksheet, row, column } = params;
  console.log('클릭:', worksheet.getRange(row, column).getA1Notation());
});
```

### 셀 편집 이벤트

| 이벤트 | 설명 | 취소 가능 |
|--------|------|-----------|
| `BeforeSheetEditStart` | 편집 시작 전 | O (`params.cancel = true`) |
| `SheetEditStarted` | 편집 시작 후 | X |
| `SheetEditChanging` | 편집 중 내용 변경 | X |
| `BeforeSheetEditEnd` | 편집 종료 전 | O (`params.cancel = true`) |
| `SheetEditEnded` | 편집 종료 후 | X |

```javascript
// 편집 시작 차단 (특정 조건)
univerAPI.addEvent(univerAPI.Event.BeforeSheetEditStart, (params) => {
  const { row, column } = params;
  if (row === 0) {
    params.cancel = true; // 첫 번째 행 편집 차단
  }
});

// 편집 종료 시 값 검증
univerAPI.addEvent(univerAPI.Event.BeforeSheetEditEnd, (params) => {
  const { row, column, value } = params;
  console.log(`셀(${row},${column}) 편집 완료:`, value);
});

// 편집 중 실시간 값 변경 감지
univerAPI.addEvent(univerAPI.Event.SheetEditChanging, (params) => {
  const { row, column, value } = params;
  console.log(`편집 중(${row},${column}):`, value);
});
```

### 선택 영역 변경 이벤트

```javascript
univerAPI.addEvent(univerAPI.Event.SelectionChanged, (params) => {
  const { worksheet, workbook, selections } = params;
  console.log('선택 영역 변경:', selections);
});
```

### 클립보드 이벤트

```javascript
// 복사 전
univerAPI.addEvent(univerAPI.Event.BeforeClipboardChange, (params) => {
  const { text, html } = params;
  console.log('복사할 내용:', text);
  // params.cancel = true; // 복사 취소
});

// 붙여넣기 전
univerAPI.addEvent(univerAPI.Event.BeforeClipboardPaste, (params) => {
  const { text, html } = params;
  console.log('붙여넣을 내용:', text);
  // params.cancel = true; // 붙여넣기 취소
});

// 복사 후
univerAPI.addEvent(univerAPI.Event.ClipboardChanged, (params) => {
  console.log('복사 완료:', params.text);
});

// 붙여넣기 후
univerAPI.addEvent(univerAPI.Event.ClipboardPasted, (params) => {
  console.log('붙여넣기 완료:', params.text);
});
```

---

## 12. UI 커스터마이징

### 메뉴 항목 숨기기/비활성화

```javascript
UniverSheetsCorePreset({
  container: 'app',
  menu: {
    'sheet.command.set-range-bold': {
      hidden: true,  // 굵기 버튼 숨기기
    },
    'sheet.command.set-range-italic': {
      hidden: true,
      disabled: true, // 기울기 버튼 숨기기 + 비활성화
    },
  },
});
```

### 주요 메뉴 항목 ID 목록

**툴바 메뉴:**

| ID | 기능 |
|----|------|
| `univer.command.undo` | 실행 취소 |
| `univer.command.redo` | 다시 실행 |
| `sheet.command.set-range-bold` | 굵게 |
| `sheet.command.set-range-italic` | 기울임꼴 |
| `sheet.command.set-range-underline` | 밑줄 |
| `sheet.command.set-range-fontsize` | 글꼴 크기 |
| `sheet.command.set-range-text-color` | 글꼴 색상 |
| `sheet.command.set-background-color` | 배경 색상 |
| `sheet.command.set-border-basic` | 테두리 |
| `sheet.command.set-horizontal-text-align` | 가로 정렬 |
| `sheet.command.set-vertical-text-align` | 세로 정렬 |
| `sheet.command.set-text-wrap` | 텍스트 줄바꿈 |
| `sheet.command.add-worksheet-merge` | 셀 병합 |
| `sheet.command.smart-toggle-filter` | 필터 토글 |

**컨텍스트 메뉴:**

| ID | 기능 |
|----|------|
| `sheet.command.copy` | 복사 |
| `sheet.command.cut` | 잘라내기 |
| `sheet.command.paste` | 붙여넣기 |
| `sheet.command.insert-row-before` | 행 삽입 (위) |
| `sheet.command.insert-col-before` | 열 삽입 (왼쪽) |
| `sheet.command.remove-row-confirm` | 행 삭제 |
| `sheet.command.remove-col-confirm` | 열 삭제 |
| `sheet.command.set-selection-frozen` | 틀 고정 |

### 리본 타입

```javascript
// 클래식 모드 (기본, 그룹 + 툴바 2행 구조)
UniverSheetsCorePreset({ ribbonType: 'classic' });

// 컴팩트 모드 (그룹 + 툴바 1행)
UniverSheetsCorePreset({ ribbonType: 'collapsed' });

// 단순 모드 (그룹 없이 평면 배치)
UniverSheetsCorePreset({ ribbonType: 'simple' });
```

### 푸터 커스터마이징

```javascript
// 푸터 완전 숨기기
UniverSheetsCorePreset({ footer: false });

// 푸터 부분 커스터마이징
UniverSheetsCorePreset({
  footer: {
    sheetBar: true,       // 시트 탭
    statisticBar: false,   // 통계 바 숨기기
    menus: true,           // 하단 메뉴
    zoomSlider: false,     // 줌 슬라이더 숨기기
  },
});
```

---

## 13. 라이프사이클 관리

### 라이프사이클 단계

일부 Facade API는 특정 라이프사이클 단계에서만 동작합니다.

```javascript
const disposable = univerAPI.addEvent(
  univerAPI.Event.LifeCycleChanged,
  ({ stage }) => {
    if (stage === univerAPI.Enum.LifecycleStages.Rendered) {
      // UI 렌더링 완료 시점
      console.log('렌더링 완료');
    }
    if (stage === univerAPI.Enum.LifecycleStages.Steady) {
      // 안정화 단계 (모든 초기화 완료)
      console.log('안정화 완료');
    }
  }
);
```

### Workbook 생성 후 API 호출 주의사항

```javascript
univerAPI.createWorkbook({});

// 즉시 API 호출 시 동작하지 않을 수 있음
// 아래와 같이 라이프사이클 이벤트에서 실행 권장
const disposable = univerAPI.addEvent(
  univerAPI.Event.LifeCycleChanged,
  ({ stage }) => {
    if (stage === univerAPI.Enum.LifecycleStages.Rendered) {
      const fWorkbook = univerAPI.getActiveWorkbook();
      const fWorksheet = fWorkbook.getActiveSheet();
      const fRange = fWorksheet.getRange('A1');
      fRange.setValue('초기 데이터');
      disposable.dispose();
    }
  }
);
```

---

## 14. v0.16.x 주요 변경사항 및 Breaking Changes

### v0.16.0 (Presets) 주요 변경

1. **Shape 기능 추가**: Advanced Preset에 Shape 기능 포함
2. **SDK 0.16.0 업데이트**: Core SDK 동기화

### v0.16.1 (Core) 주요 변경

1. **Shape 삽입 지원 (Beta)**
   - `@univerjs-pro/sheets-shape` + `@univerjs-pro/sheets-shape-ui`
   - 상단 툴바 > 삽입 메뉴에서 접근

2. **History 기능 리팩토링**: 더 풍부한 히스토리 정보 표시

3. **버그 수정**
   - 퍼센트 숫자 서식 편집 이슈 수정
   - 대량 병합 셀 + 숨겨진 행 렌더링 성능 개선
   - 셀 에디터 최대 높이 계산 이슈 수정
   - `customHeaders` 설정값 전달 이슈 수정
   - 컨텍스트 메뉴 리팩토링 및 성능 최적화

### v0.6.0 이후의 주요 Breaking Changes (참고)

> 0.6.0 이후 적용된 Breaking Changes로, 0.16.0에서도 유효합니다.

1. **`@univerjs/facade` 패키지 제거**: 새로운 Facade 사용 방식으로 마이그레이션 필요
   ```javascript
   // 이전 (더 이상 사용 불가)
   import { FUniver } from '@univerjs/core';
   // 이후 (Plugin 모드에서만 해당)
   import { FUniver } from '@univerjs/core/facade';
   ```

2. **React 16 호환성**: 추가 설정 필요

3. **Redi 뷰 관련 API 이동**:
   ```javascript
   // 이전
   import { useDependency } from '@univerjs/core';
   // 이후
   import { useDependency } from '@univerjs/ui';
   ```

---

## 15. 실전 코드 예제 모음

### 예제 1: 성적 관리 스프레드시트

```javascript
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US';
import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets';
import '@univerjs/preset-sheets-core/lib/index.css';

const { univerAPI } = createUniver({
  locale: LocaleType.EN_US,
  locales: {
    [LocaleType.EN_US]: mergeLocales(UniverPresetSheetsCoreEnUS),
  },
  presets: [
    UniverSheetsCorePreset({ container: 'app' }),
  ],
});

univerAPI.createWorkbook({
  name: '성적 관리',
  sheets: {
    'sheet-01': {
      id: 'sheet-01',
      name: '1학기',
      rowCount: 50,
      columnCount: 10,
      cellData: {
        0: {
          0: { v: '학번' },
          1: { v: '이름' },
          2: { v: '국어' },
          3: { v: '영어' },
          4: { v: '수학' },
          5: { v: '합계' },
          6: { v: '평균' },
        },
        1: {
          0: { v: '001' },
          1: { v: '홍길동' },
          2: { v: 85 },
          3: { v: 92 },
          4: { v: 78 },
          5: { f: '=SUM(C2:E2)' },
          6: { f: '=AVERAGE(C2:E2)' },
        },
        2: {
          0: { v: '002' },
          1: { v: '김영희' },
          2: { v: 95 },
          3: { v: 88 },
          4: { v: 91 },
          5: { f: '=SUM(C3:E3)' },
          6: { f: '=AVERAGE(C3:E3)' },
        },
      },
    },
  },
});
```

### 예제 2: 셀 편집 이벤트 모니터링

```javascript
// 편집 시작 전 권한 체크
univerAPI.addEvent(univerAPI.Event.BeforeSheetEditStart, (params) => {
  const { row, column } = params;
  // 헤더 행(0행) 편집 차단
  if (row === 0) {
    params.cancel = true;
    console.log('헤더 행은 편집할 수 없습니다.');
  }
});

// 편집 종료 시 유효성 검증
univerAPI.addEvent(univerAPI.Event.BeforeSheetEditEnd, (params) => {
  const { row, column, value } = params;
  // C~E열(2~4)에 숫자만 허용
  if (column >= 2 && column <= 4 && row > 0) {
    const numValue = Number(value?.body?.dataStream?.trim());
    if (isNaN(numValue) || numValue < 0 || numValue > 100) {
      params.cancel = true;
      console.log('0~100 사이의 숫자만 입력할 수 있습니다.');
    }
  }
});

// 실시간 편집 내용 추적
univerAPI.addEvent(univerAPI.Event.SheetEditChanging, (params) => {
  const { row, column, value } = params;
  console.log(`편집 중 (${row}, ${column}):`, value);
});
```

### 예제 3: 프로그래밍 방식으로 데이터 조작

```javascript
const fWorkbook = univerAPI.getActiveWorkbook();
const fWorksheet = fWorkbook.getActiveSheet();

// 범위에 값 일괄 설정
const headerRange = fWorksheet.getRange('A1:D1');
headerRange.setValues([['제품명', '단가', '수량', '합계']]);
headerRange
  .setFontWeight('bold')
  .setFontSize(12)
  .setFontColor('#FFFFFF');

// 데이터 입력
const dataRange = fWorksheet.getRange('A2:C4');
dataRange.setValues([
  ['노트북', 1200000, 5],
  ['마우스', 35000, 20],
  ['키보드', 89000, 10],
]);

// 수식 설정 (합계 = 단가 * 수량)
fWorksheet.getRange('D2').setValue('=B2*C2');
fWorksheet.getRange('D3').setValue('=B3*C3');
fWorksheet.getRange('D4').setValue('=B4*C4');

// 합계 행 추가
fWorksheet.getRange('A5').setValue('총합계');
fWorksheet.getRange('D5').setValue('=SUM(D2:D4)');
fWorksheet.getRange('A5:D5')
  .setFontWeight('bold')
  .setFontSize(11);
```

### 예제 4: 선택 영역 이벤트 활용

```javascript
// 선택 영역 변경 시 정보 표시
univerAPI.addEvent(univerAPI.Event.SelectionChanged, (params) => {
  const { worksheet, selections } = params;
  if (selections && selections.length > 0) {
    const range = selections[0];
    console.log(`선택 범위: ${range.getActiveRange().getA1Notation()}`);
  }
});

// 특정 범위 프로그래밍 방식으로 선택
const fWorksheet = univerAPI.getActiveWorkbook().getActiveSheet();
const targetRange = fWorksheet.getRange('B2:D5');
targetRange.activate();
```

### 예제 5: 워터마크 추가 플러그인 사용

```javascript
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US';
import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets';
import { UniverWatermarkPlugin } from '@univerjs/watermark';
import '@univerjs/preset-sheets-core/lib/index.css';

const { univerAPI } = createUniver({
  locale: LocaleType.EN_US,
  locales: {
    [LocaleType.EN_US]: mergeLocales(UniverPresetSheetsCoreEnUS),
  },
  presets: [
    UniverSheetsCorePreset({ container: 'app' }),
  ],
  plugins: [
    [UniverWatermarkPlugin, {
      textWatermarkSettings: {
        content: 'CONFIDENTIAL',
        fontSize: 16,
        color: 'rgb(0,0,0)',
        bold: false,
        italic: false,
        direction: 'ltr',
        x: 60,
        y: 36,
        repeat: true,
        spacingX: 200,
        spacingY: 100,
        rotate: -30,
        opacity: 0.1,
      },
    }],
  ],
});

univerAPI.createWorkbook({});
```

### 예제 6: 커스텀 셀 렌더링

```javascript
// 마커 기반 커스텀 렌더링
const fWorksheet = univerAPI.getActiveWorkbook().getActiveSheet();

// 셀에 마커 설정
fWorksheet.getRange('B2').setValue({
  v: '완료',
  custom: { status: 'done' },
});

fWorksheet.getRange('B3').setValue({
  v: '진행중',
  custom: { status: 'progress' },
});

// 커스텀 렌더러 등록
univerAPI.getSheetHooks().onCellRender([{
  drawWith: (ctx, info, skeleton, spreadsheets) => {
    const { data } = info;
    if (data?.custom?.status === 'done') {
      const { primaryWithCoord } = info;
      const { startX, startY } = primaryWithCoord;
      ctx.fillText('✅', startX, startY + 10);
    }
    if (data?.custom?.status === 'progress') {
      const { primaryWithCoord } = info;
      const { startX, startY } = primaryWithCoord;
      ctx.fillText('🔄', startX, startY + 10);
    }
  },
}]);

fWorksheet.refreshCanvas();
```

### 예제 7: 다크 모드 토글

```javascript
const { univerAPI } = createUniver({
  locale: LocaleType.EN_US,
  locales: { /* ... */ },
  darkMode: false,
  presets: [
    UniverSheetsCorePreset({ container: 'app' }),
  ],
});

// 다크 모드 동적 토글
univerAPI.toggleDarkMode(true);   // 다크 모드 활성화
univerAPI.toggleDarkMode(false);  // 라이트 모드로 복원
```

### 예제 8: 셀 에디터 실시간 데이터 조회

```javascript
// 셀 편집기의 실시간 내용을 가져오는 방법
// (Univer Sheets의 셀 편집기는 Univer Docs 엔진 기반)
univerAPI.onCommandExecuted((command) => {
  const { id } = command;
  if (
    id === 'doc.command.insert-text' ||
    id === 'doc.command.delete-text'
  ) {
    const doc = univerAPI.getActiveDocument();
    if (doc) {
      const snapshot = doc.getSnapshot();
      console.log('현재 입력 중:', snapshot.body?.dataStream);
    }
  }
});
```

### 예제 9: 외부 버튼에서 편집 종료 후 스냅샷 조회

```javascript
// 외부 버튼 클릭 시 편집 종료 후 데이터 가져오기
document.getElementById('saveBtn').addEventListener('click', async () => {
  const fWorkbook = univerAPI.getActiveWorkbook();

  // 편집 중인 셀의 데이터를 스냅샷에 반영시킴
  await fWorkbook.endEditingAsync(true);

  // 이제 최신 데이터 조회 가능
  const fWorksheet = fWorkbook.getActiveSheet();
  const values = fWorksheet.getRange('A1:D10').getValues();
  console.log('저장할 데이터:', values);
});
```

---

## 참고: 공식 자료 링크 모음

| 자료 | URL |
|------|-----|
| 설치 가이드 | https://docs.univer.ai/guides/docs/getting-started/installation |
| Core Features | https://docs.univer.ai/en-US/guides/sheets/features/core |
| Range & Selection | https://docs.univer.ai/guides/sheets/features/core/range-selection |
| Facade API 가이드 | https://docs.univer.ai/en-US/guides/sheets/getting-started/facade |
| Facade API Reference | https://reference.univer.ai/en-US |
| FWorkbook API | https://reference.univer.ai/en-US/classes/FWorkbook |
| FWorksheet API | https://reference.univer.ai/en-US/classes/FWorksheet |
| FRange API | https://reference.univer.ai/en-US/classes/FRange |
| Preset 패키지 | https://docs.univer.ai/reference/packages/presets/univerjs/preset-sheets-core |
| GitHub Releases | https://github.com/dream-num/univer/releases |
| GitHub Discussions | https://github.com/dream-num/univer/discussions |
| Playground | https://docs.univer.ai/en-US/playground/sheets/basic-via-preset |
