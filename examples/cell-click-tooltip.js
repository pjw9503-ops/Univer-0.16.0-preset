/**
 * Univer 0.16.0 Preset 모드 - 셀 클릭 시 값 기반 동적 툴팁
 *
 * 공식 API 기준:
 * - FRange.attachPopup(): 셀에 팝업(툴팁) 부착
 *   @see https://reference.univer.ai/en-US/classes/FRange#attachpopup
 * - FRange.attachAlertPopup(): 셀에 알림 팝업 부착
 *   @see https://reference.univer.ai/en-US/classes/FRange#attachalertpopup
 * - univerAPI.Event.CellClicked: 셀 클릭 이벤트
 *   @see https://docs.univer.ai/guides/sheets/features/core/range-selection
 * - univerAPI.registerComponent(): 커스텀 컴포넌트 등록
 *   @see https://docs.univer.ai/en-US/guides/docs/ui/components
 */

// =============================================================
// 방법 1: attachAlertPopup - 간단한 텍스트 툴팁 (가장 간단)
// =============================================================

/**
 * 셀 클릭 시 셀 값을 알림 팝업으로 표시하는 가장 간단한 방식.
 * 별도 컴포넌트 등록 없이 title/message만 전달하면 됩니다.
 *
 * @param {object} univerAPI - createUniver()에서 반환된 Facade API 인스턴스
 */
function setupAlertTooltip(univerAPI) {
  let currentDisposable = null;

  univerAPI.addEvent(
    univerAPI.Event.CellClicked,
    (params) => {
      const { worksheet, row, column } = params;
      const fRange = worksheet.getRange(row, column);
      const cellValue = fRange.getValue();
      const cellNotation = fRange.getA1Notation();

      // 이전 팝업 제거
      if (currentDisposable) {
        currentDisposable.dispose();
        currentDisposable = null;
      }

      // 값이 있는 셀만 툴팁 표시
      if (cellValue == null || cellValue === '') return;

      // type: 0 = info, 1 = warning, 2 = error
      currentDisposable = fRange.attachAlertPopup({
        title: cellNotation,
        message: String(cellValue),
        type: 0,
      });
    }
  );
}

// =============================================================
// 방법 2: attachPopup + registerComponent - 커스텀 툴팁 UI
// =============================================================

/**
 * 커스텀 컴포넌트를 등록하여 셀 클릭 시 스타일링된 툴팁을 표시.
 * React 기반 컴포넌트를 registerComponent로 등록 후,
 * attachPopup의 componentKey로 참조합니다.
 *
 * @param {object} univerAPI - createUniver()에서 반환된 Facade API 인스턴스
 */
function setupCustomTooltip(univerAPI) {
  // 1) 툴팁 컴포넌트 등록
  univerAPI.registerComponent(
    'CellTooltip',
    (props) => {
      const { cellAddress, cellValue, cellType } = props;

      const containerStyle = {
        background: '#ffffff',
        border: '1px solid #e0e0e0',
        borderRadius: '6px',
        padding: '10px 14px',
        boxShadow: '0 2px 8px rgba(0, 0, 0, 0.12)',
        fontSize: '13px',
        lineHeight: '1.5',
        maxWidth: '280px',
        fontFamily:
          '-apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif',
      };

      const labelStyle = {
        color: '#666666',
        fontSize: '11px',
        marginBottom: '2px',
      };

      const valueStyle = {
        color: '#1a1a1a',
        fontWeight: '600',
        fontSize: '14px',
        wordBreak: 'break-all',
      };

      const badgeStyle = {
        display: 'inline-block',
        background: cellType === 'formula' ? '#e8f5e9' : '#e3f2fd',
        color: cellType === 'formula' ? '#2e7d32' : '#1565c0',
        padding: '1px 6px',
        borderRadius: '3px',
        fontSize: '10px',
        marginLeft: '6px',
      };

      // React.createElement 기반 (JSX 트랜스파일 불필요)
      const React = props.__react || window.React;
      if (!React) {
        const el = document.createElement('div');
        el.style.cssText = Object.entries(containerStyle)
          .map(([k, v]) => `${k.replace(/[A-Z]/g, '-$&').toLowerCase()}:${v}`)
          .join(';');
        el.innerHTML = `
          <div style="color:#666;font-size:11px;margin-bottom:2px">
            ${cellAddress}
            <span style="display:inline-block;background:${
              cellType === 'formula' ? '#e8f5e9' : '#e3f2fd'
            };color:${
              cellType === 'formula' ? '#2e7d32' : '#1565c0'
            };padding:1px 6px;border-radius:3px;font-size:10px;margin-left:6px">
              ${cellType}
            </span>
          </div>
          <div style="color:#1a1a1a;font-weight:600;font-size:14px;word-break:break-all">
            ${cellValue}
          </div>
        `;
        return el;
      }

      return React.createElement('div', { style: containerStyle }, [
        React.createElement('div', { style: labelStyle, key: 'label' }, [
          cellAddress,
          React.createElement(
            'span',
            { style: badgeStyle, key: 'badge' },
            cellType
          ),
        ]),
        React.createElement(
          'div',
          { style: valueStyle, key: 'value' },
          cellValue
        ),
      ]);
    }
  );

  // 2) 클릭 이벤트에서 동적으로 팝업 부착
  let currentDisposable = null;

  univerAPI.addEvent(
    univerAPI.Event.CellClicked,
    (params) => {
      const { worksheet, row, column } = params;
      const fRange = worksheet.getRange(row, column);
      const cellValue = fRange.getValue();
      const formula = fRange.getFormula();
      const displayValue = fRange.getDisplayValue();
      const cellNotation = fRange.getA1Notation();

      // 이전 팝업 제거
      if (currentDisposable) {
        currentDisposable.dispose();
        currentDisposable = null;
      }

      if (cellValue == null && !formula) return;

      const cellType = formula
        ? 'formula'
        : typeof cellValue === 'number'
          ? 'number'
          : 'text';

      const shownValue = formula
        ? `${displayValue} (${formula})`
        : String(cellValue);

      currentDisposable = fRange.attachPopup({
        componentKey: 'CellTooltip',
        direction: 'bottom',
        cellAddress: cellNotation,
        cellValue: shownValue,
        cellType: cellType,
      });
    }
  );
}

// =============================================================
// 방법 3: attachPopup + 조건부 스타일 - 값 유형별 다른 툴팁
// =============================================================

/**
 * 셀 값의 유형(숫자/텍스트/수식/에러)에 따라
 * 서로 다른 스타일의 툴팁을 동적으로 생성합니다.
 *
 * @param {object} univerAPI - createUniver()에서 반환된 Facade API 인스턴스
 */
function setupConditionalTooltip(univerAPI) {
  const TOOLTIP_CONFIGS = {
    number: {
      bg: '#e3f2fd',
      border: '#90caf9',
      icon: '#',
      label: '숫자',
    },
    text: {
      bg: '#fff3e0',
      border: '#ffcc80',
      icon: 'T',
      label: '텍스트',
    },
    formula: {
      bg: '#e8f5e9',
      border: '#a5d6a7',
      icon: 'fx',
      label: '수식',
    },
    boolean: {
      bg: '#fce4ec',
      border: '#f48fb1',
      icon: '?',
      label: '논리값',
    },
    error: {
      bg: '#ffebee',
      border: '#ef9a9a',
      icon: '!',
      label: '오류',
    },
  };

  // 유형별 컴포넌트 등록
  Object.entries(TOOLTIP_CONFIGS).forEach(([type, config]) => {
    univerAPI.registerComponent(
      `Tooltip_${type}`,
      (props) => {
        const el = document.createElement('div');
        el.style.cssText = [
          `background:${config.bg}`,
          `border:1px solid ${config.border}`,
          'border-radius:6px',
          'padding:8px 12px',
          'font-size:13px',
          'box-shadow:0 2px 6px rgba(0,0,0,0.08)',
          'max-width:300px',
        ].join(';');

        el.innerHTML = `
          <div style="display:flex;align-items:center;gap:6px;margin-bottom:4px">
            <span style="
              width:20px;height:20px;border-radius:50%;
              background:${config.border};color:#fff;
              display:flex;align-items:center;justify-content:center;
              font-size:11px;font-weight:bold;
            ">${config.icon}</span>
            <span style="font-size:11px;color:#666">
              ${props.cellAddress} · ${config.label}
            </span>
          </div>
          <div style="font-weight:600;color:#1a1a1a;word-break:break-all">
            ${props.displayText}
          </div>
          ${
            props.formulaText
              ? `<div style="font-size:11px;color:#888;margin-top:4px;font-family:monospace">
                   ${props.formulaText}
                 </div>`
              : ''
          }
        `;
        return el;
      }
    );
  });

  let currentDisposable = null;

  univerAPI.addEvent(
    univerAPI.Event.CellClicked,
    (params) => {
      const { worksheet, row, column } = params;
      const fRange = worksheet.getRange(row, column);
      const rawValue = fRange.getValue();
      const formula = fRange.getFormula();
      const displayValue = fRange.getDisplayValue();
      const notation = fRange.getA1Notation();

      if (currentDisposable) {
        currentDisposable.dispose();
        currentDisposable = null;
      }

      if (rawValue == null && !formula) return;

      let valueType = 'text';
      if (formula) {
        valueType = 'formula';
      } else if (typeof rawValue === 'number') {
        valueType = 'number';
      } else if (typeof rawValue === 'boolean') {
        valueType = 'boolean';
      } else if (
        typeof rawValue === 'string' &&
        rawValue.startsWith('#')
      ) {
        valueType = 'error';
      }

      currentDisposable = fRange.attachPopup({
        componentKey: `Tooltip_${valueType}`,
        direction: 'bottom',
        cellAddress: notation,
        displayText: displayValue ?? String(rawValue),
        formulaText: formula || null,
      });
    }
  );
}

// =============================================================
// 방법 4: CellHover 이벤트 기반 - 호버 시 툴팁 (마우스 올릴 때)
// =============================================================

/**
 * 셀 위에 마우스를 올릴 때 툴팁을 표시하고,
 * 다른 셀로 이동하면 자동으로 이전 툴팁을 제거합니다.
 *
 * @param {object} univerAPI - createUniver()에서 반환된 Facade API 인스턴스
 */
function setupHoverTooltip(univerAPI) {
  univerAPI.registerComponent(
    'HoverTooltip',
    (props) => {
      const el = document.createElement('div');
      el.style.cssText = [
        'background:rgba(0,0,0,0.8)',
        'color:#fff',
        'border-radius:4px',
        'padding:6px 10px',
        'font-size:12px',
        'max-width:240px',
        'word-break:break-all',
        'pointer-events:none',
      ].join(';');
      el.textContent = props.text;
      return el;
    }
  );

  let currentDisposable = null;
  let lastCellKey = null;

  univerAPI.addEvent(
    univerAPI.Event.CellHover,
    (params) => {
      const { worksheet, row, column } = params;
      const cellKey = `${row}:${column}`;

      // 같은 셀이면 중복 처리하지 않음
      if (cellKey === lastCellKey) return;
      lastCellKey = cellKey;

      // 이전 툴팁 제거
      if (currentDisposable) {
        currentDisposable.dispose();
        currentDisposable = null;
      }

      const fRange = worksheet.getRange(row, column);
      const cellValue = fRange.getValue();

      if (cellValue == null || cellValue === '') return;

      currentDisposable = fRange.attachPopup({
        componentKey: 'HoverTooltip',
        direction: 'top',
        text: `${fRange.getA1Notation()}: ${cellValue}`,
      });
    }
  );
}

// =============================================================
// 전체 통합 예제 (Preset 모드 기본 설정 포함)
// =============================================================

/*
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
      },
    },
  },
});

// 원하는 방법 선택:
// setupAlertTooltip(univerAPI);    // 방법 1: 간단한 알림 팝업
// setupCustomTooltip(univerAPI);   // 방법 2: 커스텀 UI 툴팁
// setupConditionalTooltip(univerAPI); // 방법 3: 값 유형별 툴팁
// setupHoverTooltip(univerAPI);    // 방법 4: 호버 시 툴팁
*/
