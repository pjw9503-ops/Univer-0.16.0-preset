import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US';
import {
  createUniver,
  LocaleType,
  mergeLocales,
} from '@univerjs/presets';

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
  sheets: {
    'sheet-01': {
      id: 'sheet-01',
      name: 'Sheet1',
      rowCount: 50,
      columnCount: 10,
      cellData: {
        0: {
          0: { v: '제품명' },
          1: { v: '가격' },
          2: { v: '수량' },
          3: { v: '합계' },
        },
        1: {
          0: { v: '노트북' },
          1: { v: 1200000 },
          2: { v: 5 },
          3: { f: '=B2*C2' },
        },
        2: {
          0: { v: '마우스' },
          1: { v: 35000 },
          2: { v: 20 },
          3: { f: '=B3*C3' },
        },
        3: {
          0: { v: '키보드' },
          1: { v: 89000 },
          2: { v: 10 },
          3: { f: '=B4*C4' },
        },
        4: {
          0: { v: '모니터' },
          1: { v: 450000 },
          2: { v: 3 },
          3: { f: '=B5*C5' },
        },
        5: {
          0: { v: '총합계' },
          3: { f: '=SUM(D2:D5)' },
        },
      },
    },
  },
});

setupAlertTooltip(univerAPI);

/**
 * 셀 클릭 시 셀 값을 알림 팝업으로 표시.
 * @param {object} api - createUniver()에서 반환된 Facade API 인스턴스
 */
function setupAlertTooltip(api) {
  let currentDisposable = null;

  api.addEvent(
    api.Event.CellClicked,
    (params) => {
      const { worksheet, row, column } = params;
      const fRange = worksheet.getRange(row, column);
      const cellValue = fRange.getValue();
      const cellNotation = fRange.getA1Notation();

      if (currentDisposable) {
        currentDisposable.dispose();
        currentDisposable = null;
      }

      if (cellValue == null || cellValue === '') return;

      currentDisposable = fRange.attachAlertPopup({
        title: cellNotation,
        message: String(cellValue),
        type: 0,
      });
    },
  );
}
