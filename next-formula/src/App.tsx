import { useState, useMemo, useEffect, useRef } from 'react';
import { AgGridReact } from 'ag-grid-react';
import { AllCommunityModule, ModuleRegistry, type ColDef } from 'ag-grid-community';
import Editor, { useMonaco } from '@monaco-editor/react';
import { HyperFormula } from 'hyperformula';
import 'ag-grid-community/styles/ag-grid.css';
import 'ag-grid-community/styles/ag-theme-quartz.css';

// Register AG Grid modules
ModuleRegistry.registerModules([AllCommunityModule]);

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

function shiftFormula(formula: string, deltaCol: number, deltaRow: number) {
  let shifted = formula.replace(/(^|[^A-Za-z])(\$?[A-Za-z]+)(\$?[0-9]+)\b/g, (_, prefix, colPart, rowPart) => {
    let isAbsCol = colPart.startsWith('$');
    let isAbsRow = rowPart.startsWith('$');
    
    let colStr = isAbsCol ? colPart.substring(1).toUpperCase() : colPart.toUpperCase();
    let rowNum = isAbsRow ? parseInt(rowPart.substring(1), 10) : parseInt(rowPart, 10);
    
    let newColStr = colStr;
    if (!isAbsCol && deltaCol !== 0) {
       let cNum = 0;
       for (let i = 0; i < colStr.length; i++) {
           cNum = cNum * 26 + (colStr.charCodeAt(i) - 64);
       }
       cNum += deltaCol;
       if (cNum < 1) cNum = 1; // prevent negative columns
       newColStr = '';
       let temp = cNum;
       while (temp > 0) {
           let rem = (temp - 1) % 26;
           newColStr = String.fromCharCode(65 + rem) + newColStr;
           temp = Math.floor((temp - 1) / 26);
       }
    }
    
    let newRowNum = rowNum;
    if (!isAbsRow && deltaRow !== 0) {
       newRowNum += deltaRow;
       if (newRowNum < 1) newRowNum = 1;
    }
    
    return prefix + (isAbsCol ? '$' : '') + newColStr + (isAbsRow ? '$' : '') + newRowNum;
  });

  // Shift alias references like [HP]2 or [HP]$2
  shifted = shifted.replace(/(\[[^\]]+\])(\$?[0-9]+)\b/g, (_, aliasPart, rowPart) => {
    let isAbsRow = rowPart.startsWith('$');
    let rowNum = isAbsRow ? parseInt(rowPart.substring(1), 10) : parseInt(rowPart, 10);
    
    let newRowNum = rowNum;
    if (!isAbsRow && deltaRow !== 0) {
       newRowNum += deltaRow;
       if (newRowNum < 1) newRowNum = 1;
    }
    return aliasPart + (isAbsRow ? '$' : '') + newRowNum;
  });

  return shifted;
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

  // Phase 5: Custom Range Selection State
  const [selectionStart, setSelectionStart] = useState<{ col: number, row: number } | null>(null);
  const [selectionEnd, setSelectionEnd] = useState<{ col: number, row: number } | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [internalClipboard, setInternalClipboard] = useState<{ minRow: number, minCol: number, tsv: string } | null>(null);

  // Phase 7: Context Menu
  const [contextMenu, setContextMenu] = useState<{ x: number, y: number, a1: string } | null>(null);

  // Phase 8: Point Mode
  const [pointMode, setPointMode] = useState<{ active: boolean, cell: string, startIndex: number } | null>(null);

  // Parse global aliases map to show inside the grid
  const cellAliases = useMemo(() => {
    const map = new Map<string, string>(); // A1 -> AliasName
    formulaState.split('\n').forEach(line => {
      const matchAlias = line.match(/^ALIAS\s+\[(.*?)\]\s*=\s*([A-Z0-9]+)$/i);
      if (matchAlias) {
         map.set(matchAlias[2].toUpperCase(), matchAlias[1]);
      }
    });
    return map;
  }, [formulaState]);

  // Sync cell selection and aliases to grid context
  useEffect(() => {
    if (gridApiRef.current) {
      gridApiRef.current.setGridOption('context', { selectionStart, selectionEnd, cellAliases });
      gridApiRef.current.redrawRows(); // Force complete re-render to update badges reliably
    }
  }, [selectionStart, selectionEnd, cellAliases]);

  // Global listeners (mouseup, click off context menu)
  useEffect(() => {
    const handleMouseUp = () => setIsDragging(false);
    const handleClickOffMenu = () => setContextMenu(null);
    window.addEventListener('mouseup', handleMouseUp);
    window.addEventListener('click', handleClickOffMenu);
    return () => {
       window.removeEventListener('mouseup', handleMouseUp);
       window.removeEventListener('click', handleClickOffMenu);
    }
  }, []);

  // Copy / Paste handling
  useEffect(() => {
    const handleKeyDown = async (e: KeyboardEvent) => {
      const tag = document.activeElement?.tagName.toLowerCase();
      // Ignore if user is typing into an input, textarea, or Monaco editor
      if (tag === 'input' || tag === 'textarea' || document.activeElement?.className.includes('monaco-editor')) {
         return; 
      }

      if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
         if (!selectionStart || !selectionEnd) return;
         const minCol = Math.min(selectionStart.col, selectionEnd.col);
         const maxCol = Math.max(selectionStart.col, selectionEnd.col);
         const minRow = Math.min(selectionStart.row, selectionEnd.row);
         const maxRow = Math.max(selectionStart.row, selectionEnd.row);
         
         const lines = formulaState.split('\n');
         
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

         let tsv = '';
         for (let r = minRow; r <= maxRow; r++) {
            const rowVals = [];
            for (let c = minCol; c <= maxCol; c++) {
               const colLetter = String.fromCharCode(65 + c);
               const a1 = `${colLetter}${r + 1}`;
               
               let cellFormula = '';
               lines.forEach(line => {
                  if (line.toUpperCase().startsWith('ALIAS ')) return;
                  const resolvedLine = resolveAliases(line);
                  const match = resolvedLine.match(/^([A-Z]+[0-9]+)\s*=\s*(.*)$/i);
                  if (match && match[1].trim().toUpperCase() === a1) {
                     const originalMatch = line.match(/^(?:.*?)\s*=\s*(.*)$/i);
                     cellFormula = '=' + (originalMatch ? originalMatch[1].trim() : match[2].trim());
                  }
               });
               
               if (cellFormula) {
                  rowVals.push(cellFormula);
               } else if (dataState[r] && dataState[r][c] && dataState[r][c].value != null) {
                  rowVals.push(String(dataState[r][c].value));
               } else {
                  rowVals.push('');
               }
            }
            tsv += rowVals.join('\t') + '\n';
         }
         navigator.clipboard.writeText(tsv);
         setInternalClipboard({ minRow, minCol, tsv });
         
         // Optional: Visual feedback or just prevent default
         e.preventDefault();
         
      } else if ((e.ctrlKey || e.metaKey) && e.key === 'v') {
         if (!selectionStart || !selectionEnd) return;
         const text = await navigator.clipboard.readText();
         if (!text) return;
         
         const minCol = Math.min(selectionStart.col, selectionEnd.col);
         const minRow = Math.min(selectionStart.row, selectionEnd.row);
         
         const pasteRows = text.split(/\r?\n/).map(line => line.split('\t'));
         if (pasteRows.length > 0 && pasteRows[pasteRows.length - 1].length === 1 && pasteRows[pasteRows.length - 1][0] === '') {
             pasteRows.pop(); // Remove trailing empty newline
         }

         let deltaRow = 0;
         let deltaCol = 0;
         if (internalClipboard && internalClipboard.tsv === text) {
            deltaRow = minRow - internalClipboard.minRow;
            deltaCol = minCol - internalClipboard.minCol;
         }

         const newFormulaLines = [...formulaState.split('\n')];
         const newData = [...dataState];

         for (let pr = 0; pr < pasteRows.length; pr++) {
            const rowVals = pasteRows[pr];
            const targetRow = minRow + pr;
            while(newData.length <= targetRow) newData.push(Array(10).fill({ value: null, isFormula: false }));
            const rArray = [...newData[targetRow]];

            for (let pc = 0; pc < rowVals.length; pc++) {
               let val = rowVals[pc];
               // Apply relative/absolute shifting
               if (val.startsWith('=') && (deltaRow !== 0 || deltaCol !== 0)) {
                  val = shiftFormula(val, deltaCol, deltaRow);
               }

               const targetCol = minCol + pc;
               const colLetter = String.fromCharCode(65 + targetCol);
               const a1 = `${colLetter}${targetRow + 1}`;
               
               // Remove old formula if any
               const existingIdx = newFormulaLines.findIndex(l => {
                  const match = l.match(/^([A-Z]+[0-9]+)\s*=\s*/i);
                  return match && match[1].trim().toUpperCase() === a1;
               });
               if (existingIdx >= 0) newFormulaLines.splice(existingIdx, 1);

               if (val.startsWith('=')) {
                  newFormulaLines.push(`${a1} = ${val.substring(1).trim()}`);
                  while(rArray.length <= targetCol) rArray.push({ value: null, isFormula: false });
                  rArray[targetCol] = { value: null, isFormula: true };
               } else {
                  while(rArray.length <= targetCol) rArray.push({ value: null, isFormula: false });
                  const numVal = Number(val);
                  rArray[targetCol] = { value: val === '' ? null : (isNaN(numVal) ? val : numVal), isFormula: false };
               }
            }
            newData[targetRow] = rArray;
         }

         setFormulaState(newFormulaLines.join('\n'));
         setDataState(newData);
         e.preventDefault();
      }
    };
    
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [selectionStart, selectionEnd, formulaState, dataState, internalClipboard]);

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
  const { value, data, context, colDef, rowIndex } = params;
  if (!data || !colDef) return null;
  const colLetter = colDef.field;
  const a1 = `${colLetter}${rowIndex + 1}`;
  
  const aliasName = context?.cellAliases?.get(a1);
  const isFormula = data[`_${colLetter}_isFormula`]; // Use the stored flag
  const isError = data[`_${colLetter}_isError`]; // Use the stored flag
  let displayVal = value;
  if (value && typeof value === 'object') {
    displayVal = '#ERROR';
  }

  let backgroundColor = 'transparent';
  if (isError) {
    backgroundColor = '#fee2e2';
  } else if (isFormula) {
    backgroundColor = '#f0fdf4';
  }

  return (
    <div style={{
      position: 'relative', width: '100%', height: '100%', backgroundColor,
      color: isError ? '#ef4444' : 'inherit',
      fontWeight: isError ? 'bold' : 'normal',
      padding: '0 8px', boxSizing: 'border-box',
      overflow: 'hidden', whiteSpace: 'nowrap',
      display: 'flex', alignItems: 'center' // Align content vertically
    }}>
      <span style={{ flexGrow: 1 }}>{displayVal}</span>
      {aliasName && (
         <span style={{ 
            position: 'absolute', right: '4px', top: '50%', transform: 'translateY(-50%)', 
            fontSize: '10px', color: '#888',
            backgroundColor: '#eee', padding: '2px 4px', borderRadius: '4px',
            pointerEvents: 'none', userSelect: 'none'
         }}>
           {aliasName}
         </span>
      )}
    </div>
  );
};

  const columnDefs = useMemo<ColDef[]>(() => {
    return Array.from({ length: 10 }).map((_, i) => {
      const colLetter = String.fromCharCode(65 + i);
      return { 
        headerName: colLetter, field: colLetter, 
        editable: true, width: 100, cellRenderer: customCellRenderer,
        cellClassRules: {
          'custom-selected-cell': (params: any) => {
            const ctx = params.context;
            if (!ctx || !ctx.selectionStart || !ctx.selectionEnd) return false;
            const minCol = Math.min(ctx.selectionStart.col, ctx.selectionEnd.col);
            const maxCol = Math.max(ctx.selectionStart.col, ctx.selectionEnd.col);
            const minRow = Math.min(ctx.selectionStart.row, ctx.selectionEnd.row);
            const maxRow = Math.max(ctx.selectionStart.row, ctx.selectionEnd.row);
            
            const colIndex = params.colDef.field.charCodeAt(0) - 65;
            const rowIndex = params.node.rowIndex;
            return colIndex >= minCol && colIndex <= maxCol && rowIndex >= minRow && rowIndex <= maxRow;
          }
        }
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
    
    // Also use focused cell as selection start if we are not dragging
    if (!isDragging) {
       const colIndex = params.column.getColId().charCodeAt(0) - 65;
       setSelectionStart({ col: colIndex, row: params.rowIndex });
       setSelectionEnd({ col: colIndex, row: params.rowIndex });
    }
  };

  const onCellMouseDown = (params: any) => {
    if (params.event.button !== 0) return; // Only left click
    const colIndex = params.column.getColId().charCodeAt(0) - 65;
    const rowIndex = params.rowIndex;
    setSelectionStart({ col: colIndex, row: rowIndex });
    setSelectionEnd({ col: colIndex, row: rowIndex });
    setIsDragging(true);
  };

  const onCellMouseOver = (params: any) => {
    if (!isDragging) return;
    const colIndex = params.column.getColId().charCodeAt(0) - 65;
    const rowIndex = params.rowIndex;
    setSelectionEnd({ col: colIndex, row: rowIndex });
  };

  const onCellContextMenu = (params: any) => {
    params.event.preventDefault(); // Stop native browser menu
    const colIndex = params.column.getColId();
    const rowIndex = params.rowIndex;
    const a1 = `${colIndex}${rowIndex + 1}`;
    // Show menu slightly offset from the exact click location
    setContextMenu({ x: params.event.clientX, y: params.event.clientY, a1 });
  };

  const handleSetAlias = () => {
    if (!contextMenu) return;
    const name = window.prompt(`Enter an alias name for cell ${contextMenu.a1} (e.g. HP):`);
    if (name && name.trim()) {
       // Append to formulaState
       const newFormula = formulaState.trim() + (formulaState.trim() ? '\n' : '') + `ALIAS [${name.trim()}] = ${contextMenu.a1}`;
       setFormulaState(newFormula);
    }
    setContextMenu(null);
  };

  const handleFormulaBarKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === 'Enter') {
      if (focusedCell) commitCellChange(focusedCell, formulaBarText);
      setPointMode(null);
      return;
    }

    if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
      const text = formulaBarText;
      const lastChar = text.slice(-1);
      const isOperator = /=|\+|-|\*|\/|\(|,/.test(lastChar);
      
      if (isOperator || pointMode?.active) {
        e.preventDefault();
        
        // Determine starting cell for the arrow movement
        let startCell = focusedCell || 'A1';
        if (pointMode?.active) {
          startCell = pointMode.cell;
        }

        const colLetter = startCell.match(/[A-Z]+/)?.[0] || 'A';
        const rowNum = parseInt(startCell.match(/[0-9]+/)?.[0] || '1', 10);
        let colIdx = colLetter.charCodeAt(0) - 65;
        let rowIdx = rowNum - 1;

        if (e.key === 'ArrowUp') rowIdx = Math.max(0, rowIdx - 1);
        if (e.key === 'ArrowDown') rowIdx++; // No strict max bound since grid dynamically scales
        if (e.key === 'ArrowLeft') colIdx = Math.max(0, colIdx - 1);
        if (e.key === 'ArrowRight') colIdx = Math.min(9, colIdx + 1); // limit to J column (10 cols)

        const newCell = `${String.fromCharCode(65 + colIdx)}${rowIdx + 1}`;
        
        // Move selection highlight (Point Mode)
        setSelectionStart({ col: colIdx, row: rowIdx });
        setSelectionEnd({ col: colIdx, row: rowIdx });
        
        if (gridApiRef.current) {
           gridApiRef.current.ensureIndexVisible(rowIdx);
           gridApiRef.current.ensureColumnVisible(String.fromCharCode(65 + colIdx));
        }
        
        if (pointMode?.active) {
           // Replace the point mode cell reference with the new one
           const newText = text.substring(0, pointMode.startIndex) + newCell;
           setFormulaBarText(newText);
           setPointMode({ active: true, cell: newCell, startIndex: pointMode.startIndex });
        } else {
           // Append new cell reference
           const newText = text + newCell;
           setFormulaBarText(newText);
           setPointMode({ active: true, cell: newCell, startIndex: text.length });
        }
        return;
      }
    } else {
       // Stop point mode if they type letters or numbers
       if (pointMode) setPointMode(null);
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
            context={{ selectionStart, selectionEnd, cellAliases }}
            rowData={rowData}
            columnDefs={columnDefs}
            getRowId={(params) => params.data.id}
            onCellValueChanged={onCellValueChanged}
            onCellFocused={onCellFocused}
            onCellMouseDown={onCellMouseDown}
            onCellMouseOver={onCellMouseOver}
            onCellContextMenu={onCellContextMenu}
            onGridReady={(params) => { gridApiRef.current = params.api; }}
            suppressCellFocus={false}
            preventDefaultOnContextMenu={true}
          />
        </div>
      </div>

      {contextMenu && (
        <div style={{
          position: 'fixed',
          top: contextMenu.y,
          left: contextMenu.x,
          backgroundColor: 'white',
          border: '1px solid #ccc',
          boxShadow: '0 4px 6px rgba(0,0,0,0.1)',
          zIndex: 1000,
          padding: '4px 0',
          borderRadius: '4px',
          minWidth: '150px'
        }}>
          <div 
             style={{ padding: '8px 16px', cursor: 'pointer', fontSize: '13px', color: '#333' }}
             onMouseEnter={(e) => e.currentTarget.style.backgroundColor = '#f1f5f9'}
             onMouseLeave={(e) => e.currentTarget.style.backgroundColor = 'white'}
             onClick={handleSetAlias}
          >
            📋 Set Alias for {contextMenu.a1}...
          </div>
        </div>
      )}

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
