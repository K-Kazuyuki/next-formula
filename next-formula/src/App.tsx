import { useState, useMemo, useEffect, useRef } from 'react';
import { AgGridReact } from 'ag-grid-react';
import { ClientSideRowModelModule, ModuleRegistry, type ColDef } from 'ag-grid-community';
import Editor, { useMonaco } from '@monaco-editor/react';
import { HyperFormula } from 'hyperformula';
import 'ag-grid-community/styles/ag-grid.css';
import 'ag-grid-community/styles/ag-theme-quartz.css';

// Register AG Grid modules
ModuleRegistry.registerModules([ClientSideRowModelModule]);

function a1ToCoords(a1: string) {
  const match = a1.match(/^([A-Z]+)([0-9]+)$/i);
  if (!match) return null;
  const colStr = match[1].toUpperCase();
  const rowStr = match[2];
  let col = 0;
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64);
  }
  return { col: col - 1, row: parseInt(rowStr, 10) - 1 };
}

function escapeRegExp(string: string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
}

export type CellData = {
  value: string | number | null;
  isFormula: boolean;
};

const INITIAL_DATA: CellData[][] = [
  [{value: 'HP', isFormula: false}, {value: 'Attack', isFormula: false}, {value: 'Result', isFormula: false}],
  [{value: 100, isFormula: false}, {value: 20, isFormula: false}, {value: null, isFormula: true}],
  [{value: 200, isFormula: false}, {value: 30, isFormula: false}, {value: null, isFormula: true}],
  [{value: 50, isFormula: false}, {value: 10, isFormula: false}, {value: null, isFormula: true}]
];

const INITIAL_FORMULA = `ALIAS [HP] = A
ALIAS [攻撃力] = B
ALIAS [結果] = C

[結果]2 = [HP]2 + [攻撃力]2
[結果]3 = [HP]3 + [攻撃力]3
[結果]4 = [HP]4 + [攻撃力]4`;

export default function App() {
  const [dataState, setDataState] = useState<CellData[][]>(INITIAL_DATA);
  const [formulaState, setFormulaState] = useState(INITIAL_FORMULA);
  const [rowData, setRowData] = useState<any[]>([]);
  
  const [focusedCell, setFocusedCell] = useState<string | null>(null);
  const [formulaBarText, setFormulaBarText] = useState<string>('');

  const monaco = useMonaco();
  const gridApiRef = useRef<any>(null);

  const hf = useMemo(() => {
    const instance = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });
    instance.addSheet('Sheet1');
    return instance;
  }, []);

  useEffect(() => {
    const sheetId = hf.getSheetId('Sheet1')!;
    hf.clearSheet(sheetId);
    
    const rawData = dataState.map(row => row.map(cell => cell ? cell.value : null));
    hf.setCellContents({ sheet: sheetId, col: 0, row: 0 }, rawData);
    
    const formulaCells = new Set<string>();
    const lines = formulaState.split('\n');
    const newMarkers: any[] = [];
    const parsedFormulas = new Map<string, string>();

    // Phase 4: 1. Extract Aliases
    const aliases = new Map<string, string>();
    lines.forEach((line) => {
      const matchAlias = line.match(/^ALIAS\s+\[(.*?)\]\s*=\s*([A-Z0-9]+)$/i);
      if (matchAlias) {
         aliases.set(matchAlias[1], matchAlias[2].toUpperCase());
      }
    });

    const resolveAliases = (text: string) => {
      let resolved = text;
      aliases.forEach((ref, name) => {
        const regex = new RegExp(`\\[${escapeRegExp(name)}\\]`, 'gi');
        resolved = resolved.replace(regex, ref);
      });
      return resolved;
    };

    // 2. Parse formulas
    lines.forEach((line, index) => {
      if (!line.trim() || line.startsWith('//') || line.toUpperCase().startsWith('ALIAS ')) return;
      
      const resolvedLine = resolveAliases(line);
      const match = resolvedLine.match(/^([A-Z]+[0-9]+)\s*=\s*(.*)$/i);
      
      if (match) {
        const a1 = match[1].trim().toUpperCase();
        
        // For formula bar, we show the original unresolved right side (or resolved? Original is better)
        // Let's deduce the original right side from the original line if possible
        const originalMatch = line.match(/^(?:.*?)\s*=\s*(.*)$/i);
        parsedFormulas.set(a1, '=' + (originalMatch ? originalMatch[1].trim() : match[2].trim()));

        const coords = a1ToCoords(a1);
        if (coords) {
          try {
            hf.setCellContents(
              { sheet: sheetId, col: coords.col, row: coords.row },
              [['=' + match[2].trim()]]
            );
            formulaCells.add(a1);
          } catch (e) {
            newMarkers.push({
              startLineNumber: index + 1, startColumn: 1,
              endLineNumber: index + 1, endColumn: line.length + 1,
              message: String(e), severity: 8
            });
          }
        }
      } else {
        // Line exists but didn't match after translation
        newMarkers.push({
          startLineNumber: index + 1, startColumn: 1,
          endLineNumber: index + 1, endColumn: line.length + 1,
          message: 'Syntax error: Invalid assignment', severity: 8
        });
      }
    });

    lines.forEach((line, index) => {
      if (!line.trim() || line.startsWith('//') || line.toUpperCase().startsWith('ALIAS ')) return;
      const resolvedLine = resolveAliases(line);
      const match = resolvedLine.match(/^([A-Z]+[0-9]+)\s*=\s*(.*)$/i);
      if (match) {
        const a1 = match[1].trim().toUpperCase();
        const coords = a1ToCoords(a1);
        if (coords) {
          const val = hf.getCellValue({ sheet: sheetId, col: coords.col, row: coords.row });
          if (val && typeof val === 'object') {
            newMarkers.push({
              startLineNumber: index + 1, startColumn: 1,
              endLineNumber: index + 1, endColumn: line.length + 1,
              message: `Error: ${(val as any).message || (val as any).value || 'Unknown error'}`, severity: 8
            });
          }
        }
      }
    });

    if (monaco) {
      const model = monaco.editor.getModels()[0];
      if (model) {
        monaco.editor.setModelMarkers(model, 'hyperformula', newMarkers);
      }
    }

    const newRowData = [];
    for (let r = 0; r < 20; r++) {
      const rowObj: any = { id: String(r) }; // idを追加して再描画の点滅を防ぐ
      for (let c = 0; c < 10; c++) {
        const colLetter = String.fromCharCode(65 + c);
        const val = hf.getCellValue({ sheet: sheetId, col: c, row: r });
        let displayVal = val;
        let isError = false;
        
        if (val && typeof val === 'object') {
          displayVal = (val as any).value !== undefined ? String((val as any).value) : '#ERROR';
          isError = true;
        }
        
        rowObj[colLetter] = displayVal ?? '';
        rowObj[`_${colLetter}_isError`] = isError;
        rowObj[`_${colLetter}_isFormula`] = formulaCells.has(`${colLetter}${r + 1}`);
      }
      newRowData.push(rowObj);
    }
    setRowData(newRowData);

    if (focusedCell) {
      if (parsedFormulas.has(focusedCell)) {
        setFormulaBarText(parsedFormulas.get(focusedCell)!);
      } else {
        const coords = a1ToCoords(focusedCell);
        if (coords && dataState[coords.row] && dataState[coords.row][coords.col]) {
           setFormulaBarText(String(dataState[coords.row][coords.col]?.value ?? ''));
        } else {
           setFormulaBarText('');
        }
      }
    }

  }, [dataState, formulaState, hf, monaco, focusedCell]);

  const customCellRenderer = (params: any) => {
    if (!params.colDef) return null;
    const field = params.colDef.field;
    const isError = params.data[`_${field}_isError`];
    const isFormula = params.data[`_${field}_isFormula`];
    
    let backgroundColor = 'transparent';
    if (isError) {
      backgroundColor = '#fee2e2';
    } else if (isFormula) {
      backgroundColor = '#f0fdf4';
    }

    return (
      <div style={{
        width: '100%', height: '100%', backgroundColor,
        color: isError ? '#ef4444' : 'inherit',
        fontWeight: isError ? 'bold' : 'normal',
        padding: '0 8px', boxSizing: 'border-box',
        overflow: 'hidden', whiteSpace: 'nowrap'
      }}>
        {params.value}
      </div>
    );
  };

  const columnDefs = useMemo<ColDef[]>(() => {
    return Array.from({ length: 10 }).map((_, i) => {
      const colLetter = String.fromCharCode(65 + i);
      return { 
        headerName: colLetter, field: colLetter, 
        editable: true, width: 100, cellRenderer: customCellRenderer
      };
    });
  }, []);

  const commitCellChange = (a1: string, newValue: string) => {
    const coords = a1ToCoords(a1);
    if (!coords) return;
    const { col: colIndex, row: rowIndex } = coords;
    
    const lines = formulaState.split('\n');
    let formulaLineIndex = -1;
    let isFormulaCurrently = false;

    // For editing from formula bar, we must resolve aliases to see what line matches A1
    // Actually, formula state might say `[結果]2 = ...`
    // So looking for `C2 = ...` requires resolving left sides first!
    const aliases = new Map<string, string>();
    lines.forEach(line => {
      const matchAlias = line.match(/^ALIAS\s+\[(.*?)\]\s*=\s*([A-Z0-9]+)$/i);
      if (matchAlias) { aliases.set(matchAlias[1], matchAlias[2].toUpperCase()); }
    });
    const resolveAliases = (text: string) => {
      let resolved = text;
      aliases.forEach((ref, name) => {
        const regex = new RegExp(`\\[${escapeRegExp(name)}\\]`, 'gi');
        resolved = resolved.replace(regex, ref);
      });
      return resolved;
    };

    lines.forEach((line, i) => {
       if (line.toUpperCase().startsWith('ALIAS ')) return;
       const resolvedLine = resolveAliases(line);
       const match = resolvedLine.match(/^([A-Z]+[0-9]+)\s*=\s*/i);
       if (match && match[1].trim().toUpperCase() === a1) {
         isFormulaCurrently = true;
         formulaLineIndex = i;
       }
    });

    let isFormulaNext = false;
    let nextFormulaState = formulaState;

    if (newValue.startsWith('=')) {
      isFormulaNext = true;
      const formulaRight = newValue.substring(1).trim();
      // Wait, if it was an alias [結果]2, we should probably keep it.
      // But if the user typed =100 from the UI, we just write `A1 = 100` because we don't know the preferred alias.
      // If we are replacing, maybe we just replace the whole line with `A1 = ...` for simplicity.
      // Actually preserving left side is better:
      if (isFormulaCurrently) {
         const oldLine = lines[formulaLineIndex];
         const matchOriginalLeft = oldLine.match(/^(.*?)\s*=/);
         if (matchOriginalLeft) {
            lines[formulaLineIndex] = `${matchOriginalLeft[1].trim()} = ${formulaRight}`;
         } else {
            lines[formulaLineIndex] = `${a1} = ${formulaRight}`;
         }
      } else {
         lines.push(`${a1} = ${formulaRight}`);
      }
      nextFormulaState = lines.join('\n');
    } else {
      if (isFormulaCurrently) {
         lines.splice(formulaLineIndex, 1);
         nextFormulaState = lines.join('\n');
      }
    }

    setFormulaState(nextFormulaState);

    const newData = [...dataState];
    while(newData.length <= rowIndex) newData.push(Array(10).fill({ value: null, isFormula: false }));
    const rowArray = [...newData[rowIndex]];
    while(rowArray.length <= colIndex) rowArray.push({ value: null, isFormula: false });
    
    if (isFormulaNext) {
       rowArray[colIndex] = { value: null, isFormula: true };
    } else {
       const numericVal = Number(newValue);
       rowArray[colIndex] = {
         value: newValue === '' ? null : (isNaN(numericVal) ? newValue : numericVal),
         isFormula: false
       };
    }
    
    newData[rowIndex] = rowArray;
    setDataState(newData);
  };

  const onCellValueChanged = (params: any) => {
    const { colDef, newValue, node, oldValue } = params;
    if (newValue === oldValue) return;
    const a1 = `${colDef.field}${node.rowIndex + 1}`;
    commitCellChange(a1, newValue);
  };

  const onCellFocused = (params: any) => {
    if (params.column == null || params.rowIndex == null) return;
    const a1 = `${params.column.getColId()}${params.rowIndex + 1}`;
    setFocusedCell(a1);
  };

  const handleFormulaBarKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter' && focusedCell) {
      commitCellChange(focusedCell, formulaBarText);
      gridApiRef.current?.setFocusedCell(
         parseInt(focusedCell.match(/[0-9]+/)![0]) - 1, 
         focusedCell.match(/[A-Z]+/)![0]
      );
    }
  };

  return (
    <div style={{ display: 'flex', height: '100vh', width: '100vw', margin: 0 }}>
      {/* Left Pane: Spreadsheet View */}
      <div style={{ flex: 1, borderRight: '1px solid #ccc', display: 'flex', flexDirection: 'column' }}>
        
        <div style={{ padding: '8px 16px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', display: 'flex', alignItems: 'center', gap: '8px' }}>
          <div style={{ fontWeight: 'bold', width: '40px', textAlign: 'center', background: '#e2e8f0', padding: '4px', borderRadius: '4px' }}>
            {focusedCell || '-'}
          </div>
          <input 
            type="text" 
            style={{ flex: 1, padding: '6px', fontSize: '14px', border: '1px solid #cbd5e1', borderRadius: '4px' }}
            placeholder="Select a cell to view / edit (start with '=' for formula)"
            value={formulaBarText}
            onChange={(e) => setFormulaBarText(e.target.value)}
            onKeyDown={handleFormulaBarKeyDown}
            disabled={!focusedCell}
          />
        </div>

        <div className="ag-theme-quartz" style={{ flex: 1, height: '100%' }}>
          <AgGridReact
            rowData={rowData}
            columnDefs={columnDefs}
            getRowId={(params) => params.data.id}
            onCellValueChanged={onCellValueChanged}
            onCellFocused={onCellFocused}
            onGridReady={(params) => { gridApiRef.current = params.api; }}
          />
        </div>
      </div>

      {/* Right Pane: Code View */}
      <div style={{ flex: 1, display: 'flex', flexDirection: 'column' }}>
        <div style={{ padding: '12px 16px', background: '#f8fafc', borderBottom: '1px solid #e2e8f0', fontWeight: 'bold' }}>
          Code View (formulaState)
        </div>
        <div style={{ flex: 1, paddingTop: '8px' }}>
          <Editor
            defaultLanguage="plaintext"
            value={formulaState}
            onChange={(val) => setFormulaState(val || '')}
            options={{ 
              minimap: { enabled: false }, 
              fontSize: 14,
              lineNumbers: 'on',
              padding: { top: 16 }
            }}
          />
        </div>
      </div>
    </div>
  );
}
