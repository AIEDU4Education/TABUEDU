import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import { 
  FileSpreadsheet, 
  Upload, 
  Settings, 
  Check, 
  FileText,
  HelpCircle,
  Hash,
  RefreshCw,
  Trash2,
  AlertCircle
} from 'lucide-react';
import { CardData, CardStatus, ExcelColumnMapping } from '../types';
import { Button } from '@/components/ui/button';
import { Label } from '@/components/ui/label';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { toast } from 'sonner';

interface ExcelImporterProps {
  onImportComplete: (cards: CardData[]) => void;
  loadedCardsLength: number;
  onResetSet: () => void;
  onClearSet: () => void;
}

export const ExcelImporter: React.FC<ExcelImporterProps> = ({ 
  onImportComplete, 
  loadedCardsLength,
  onResetSet,
  onClearSet
}) => {
  const [dragActive, setDragActive] = useState(false);
  const [file, setFile] = useState<File | null>(null);
  const [columns, setColumns] = useState<string[]>([]);
  const [rawData, setRawData] = useState<any[][]>([]);
  
  // Mapping State
  const [mapping, setMapping] = useState<ExcelColumnMapping>({
    wordColumn: '',
    tabooColumns: [],
    tabooDelimiter: ',',
    isSingleColumnForTaboo: false, // Default is multiple columns as shown in screenshot
    statusColumn: undefined,
    pistaColumn: undefined,
    categoriaColumn: undefined,
  });

  const [previewCards, setPreviewCards] = useState<CardData[]>([]);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Heuristic Dictionaries
  const HEURISTICS = {
    word: ['palabra', 'concepto', 'principal', 'target', 'word', 'título', 'secreto', 'titulo'],
    pista: ['pista', 'clue', 'definicion', 'definición', 'ayuda', 'hint', 'explicacion', 'explicación'],
    status: [
      'status', 'logro', 'nivel', 'nivel log', 'nivelLogro', 'estado', 'complejidad', 
      'dificultad', 'fase', 'etapa', 'rango', 'nivel de logro', 'categoria de logro',
      'categoría de logro', 'calificacion', 'calificación', 'avance', 'progreso', 'dificultades'
    ],
    category: ['categoría', 'categoria', 'tema', 'grupo', 'unidad', 'curso', 'materia'],
    taboo: ['tabú', 'tabu', 'prohibidas', 'prohibida', 'taboo', 'no decir']
  };

  // Status mapping criteria
  const STATUS_DICT = {
    Inicio: [
      '1', '1.0', 'inicio', 'inicios', 'principiante', 'red', 'rojo', 'roja', 'básico', 'basico',
      'nivel 1', 'fácil', 'facil', 'comienzo', 'comenzar'
    ],
    Proceso: [
      '2', '2.0', 'proceso', 'procesos', 'proeso', 'proesos', 'desarrollo', 'orange', 'naranja',
      'medio', 'intermedio', 'nivel 2', 'en proceso', 'en proeso'
    ],
    Logrado: [
      '3', '3.0', 'logrado', 'lograda', 'logrados', 'logradas', 'completado', 'completada',
      'aprobado', 'aprobada', 'green', 'verde', 'alto', 'nivel 3', 'difícil', 'dificil'
    ],
    Destacado: [
      '4', '4.0', 'destacado', 'destacada', 'destacados', 'destacadas', 'excelente', 'expert',
      'experto', 'experta', 'nivel experto', 'superior', 'celeste', 'azul', 'nivel 4', 'muy difícil', 'muy dificil'
    ]
  };

  const mapStatusValue = (val: any): CardStatus => {
    if (!val) return 'Inicio';
    const s = String(val).toLowerCase().trim();
    
    // 1. First seek exact matchers
    for (const [status, matchers] of Object.entries(STATUS_DICT)) {
      if (matchers.some(m => s === m)) {
        return status as CardStatus;
      }
    }
    
    // 2. Seek partial substring matches
    for (const [status, matchers] of Object.entries(STATUS_DICT)) {
      if (matchers.some(m => {
        if (m === '1' || m === '2' || m === '3' || m === '4' || m === '1.0' || m === '2.0' || m === '3.0' || m === '4.0') {
          return false;
        }
        return s.includes(m);
      })) {
        return status as CardStatus;
      }
    }

    // 3. Fallback check for numbers in case they are isolated
    if (s.includes('1')) return 'Inicio';
    if (s.includes('2')) return 'Proceso';
    if (s.includes('3')) return 'Logrado';
    if (s.includes('4')) return 'Destacado';
    
    return 'Inicio';
  };

  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  const processExcelFile = (uploadedFile: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        if (!data) return;
        
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];
        
        // Read raw data representing complete grid
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[][];
        if (rows.length < 2) {
          toast.error("El archivo Excel no parece contener suficientes filas.");
          return;
        }

        // Get headers and prevent duplicate keys
        const rawHeaders = (rows[0] || []).map(h => String(h || "").trim());
        const uniqueHeaders: string[] = [];
        const seen: { [key: string]: number } = {};

        rawHeaders.forEach((h) => {
          if (!h) {
            uniqueHeaders.push(`Columna_${uniqueHeaders.length + 1}`);
            return;
          }
          if (seen[h] !== undefined) {
            seen[h]++;
            uniqueHeaders.push(`${h} (${seen[h] + 1})`);
          } else {
            seen[h] = 0;
            uniqueHeaders.push(h);
          }
        });

        setColumns(uniqueHeaders);
        setRawData(rows.slice(1));
        setFile(uploadedFile);
        toast.success(`Archivo "${uploadedFile.name}" cargado con éxito.`);

        // Apply dynamic smart heuristics over columns
        applySmartMapping(uniqueHeaders, rows.slice(1));
      } catch (err) {
        console.error("Error reading Excel:", err);
        toast.error("Error al leer el archivo Excel. Asegúrate de que sea un archivo XLSX o XLS válido.");
      }
    };
    reader.readAsArrayBuffer(uploadedFile);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      processExcelFile(e.dataTransfer.files[0]);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      processExcelFile(e.target.files[0]);
    }
  };

  // Applied intelligent mapping algorithm
  const applySmartMapping = (cols: string[], rows: any[][]) => {
    let wordCol = cols[0] || '';
    let pistaCol = '';
    let statusCol = '';
    let categoriaCol = '';
    let tabooCols: string[] = [];
    let isSingleColumnForTaboo = false;

    // Apply strict semantic matching heuristics
    cols.forEach((col) => {
      const cLower = col.toLowerCase();
      
      // Word heuristic
      if (HEURISTICS.word.some(h => cLower.includes(h))) {
        wordCol = col;
      }
      // Clue heuristic
      if (HEURISTICS.pista.some(h => cLower.includes(h))) {
        pistaCol = col;
      }
      // Category Heuristic
      if (HEURISTICS.category.some(h => cLower.includes(h))) {
        categoriaCol = col;
      }
    });

    // Search for a status column by explicit column name keywords
    const statusColByName = cols.find(col => 
      col !== wordCol && 
      col !== pistaCol && 
      col !== categoriaCol &&
      HEURISTICS.status.some(h => col.toLowerCase().includes(h))
    );

    if (statusColByName) {
      statusCol = statusColByName;
    } else {
      // Fallback to advanced density scanner over values
      let maxStatusDensity = -1;
      let bestStatusCol = '';

      cols.forEach((col, colIdx) => {
        if (col === wordCol || col === pistaCol || col === categoriaCol) return;

        const scanRows = rows.slice(0, 30);
        let matchCount = 0;
        
        scanRows.forEach((row) => {
          const val = String(row[colIdx] || '').toLowerCase().trim();
          if (!val) return;

          // Check if val fits status definitions
          const matchesAny = Object.values(STATUS_DICT).some(matchers => 
            matchers.some(m => {
              if (m === '1' || m === '2' || m === '3' || m === '4') {
                return val === m;
              }
              return val === m || val.includes(m);
            })
          );
          if (matchesAny) matchCount++;
        });

        if (matchCount > maxStatusDensity && matchCount > 0) {
          maxStatusDensity = matchCount;
          bestStatusCol = col;
        }
      });

      if (bestStatusCol) {
        statusCol = bestStatusCol;
      }
    }

    // Taboo word mapping heuristic
    const matchingTaboos = cols.filter(col => col !== wordCol && col !== statusCol && HEURISTICS.taboo.some(h => col.toLowerCase().includes(h)));
    if (matchingTaboos.length > 1) {
      tabooCols = matchingTaboos.slice(0, 4);
      isSingleColumnForTaboo = false;
    } else if (matchingTaboos.length === 1) {
      tabooCols = [matchingTaboos[0]];
      isSingleColumnForTaboo = true;
    } else {
      // Choose columns that started with 'tabu' or remaining columns as fallback
      const remainingCols = cols.filter(c => c !== wordCol && c !== pistaCol && c !== statusCol && c !== categoriaCol);
      const possibleTaboos = remainingCols.filter(c => c.toLowerCase().includes('tabu') || c.toLowerCase().includes('prohibida'));
      if (possibleTaboos.length > 0) {
        tabooCols = possibleTaboos.slice(0, 4);
        isSingleColumnForTaboo = possibleTaboos.length === 1;
      } else if (remainingCols.length >= 4) {
        tabooCols = remainingCols.slice(0, 4);
        isSingleColumnForTaboo = false;
      } else if (remainingCols.length > 0) {
        tabooCols = [remainingCols[0]];
        isSingleColumnForTaboo = true;
      }
    }

    const nextMapping = {
      wordColumn: wordCol,
      tabooColumns: tabooCols,
      tabooDelimiter: ',',
      isSingleColumnForTaboo,
      statusColumn: statusCol || undefined,
      pistaColumn: pistaCol || undefined,
      categoriaColumn: categoriaCol || undefined,
    };

    setMapping(nextMapping);
    generatePreview(nextMapping, cols, rows);
  };

  const generatePreview = (
    currentMap: ExcelColumnMapping, 
    cols: string[] = columns, 
    data: any[][] = rawData
  ) => {
    if (!currentMap.wordColumn) return;

    const wordIdx = cols.indexOf(currentMap.wordColumn);
    const pistaIdx = currentMap.pistaColumn ? cols.indexOf(currentMap.pistaColumn) : -1;
    const statusIdx = currentMap.statusColumn ? cols.indexOf(currentMap.statusColumn) : -1;
    const catIdx = currentMap.categoriaColumn ? cols.indexOf(currentMap.categoriaColumn) : -1;
    
    const tabooIndices = currentMap.tabooColumns.map(c => cols.indexOf(c));

    const cards: CardData[] = [];

    data.forEach((row, rowIndex) => {
      const secreto = String(row[wordIdx] || "").trim();
      if (!secreto) return; // Skip empty concepts

      // Taboo parser
      let tabooList: string[] = [];
      if (currentMap.isSingleColumnForTaboo) {
        const tabColIdx = tabooIndices[0];
        if (tabColIdx !== -1 && row[tabColIdx] !== undefined) {
          const rawTabooStr = String(row[tabColIdx] || "");
          const del = currentMap.tabooDelimiter;
          
          if (del === '\\n' || del === '\n') {
            tabooList = rawTabooStr.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
          } else {
            tabooList = rawTabooStr.split(del).map(s => s.trim()).filter(Boolean);
          }
        }
      } else {
        // Multi column map
        tabooIndices.forEach((tIdx) => {
          if (tIdx !== -1 && row[tIdx] !== undefined) {
            const val = String(row[tIdx] || "").trim();
            if (val) tabooList.push(val);
          }
        });
      }

      // If less than 4, pad or fill with blanks if empty
      while (tabooList.length < 4) {
        tabooList.push('');
      }
      tabooList = tabooList.slice(0, 4);

      const statusVal = statusIdx !== -1 ? mapStatusValue(row[statusIdx]) : 'Inicio';
      const pistaVal = pistaIdx !== -1 ? String(row[pistaIdx] || "").trim() : '';
      const catVal = catIdx !== -1 ? String(row[catIdx] || "").trim() : 'General';

      cards.push({
        id: `excel-${rowIndex}-${Date.now()}`,
        secreto,
        tabu: tabooList,
        nivelLogro: statusVal,
        pista: pistaVal,
        categoria: catVal || 'General',
        word: secreto,
        status: statusVal,
        tabooWords: tabooList
      });
    });

    setPreviewCards(cards);
  };

  const handleMappingChange = (field: keyof ExcelColumnMapping, value: any) => {
    const updated = { ...mapping, [field]: value };
    setMapping(updated);
    generatePreview(updated);
  };

  const handleApplyImport = () => {
    if (previewCards.length === 0) {
      toast.error("No hay tarjetas válidas para importar.");
      return;
    }
    
    onImportComplete(previewCards);
    toast.success(`¡Se han cargado con éxito ${previewCards.length} tarjetas del Excel!`);
    setFile(null);
    setPreviewCards([]);
  };

  return (
    <div className="flex flex-col gap-4">
      {/* Paso 1 Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-100 pb-3">
        <div className="flex items-center gap-2.5">
          <div className="w-1.5 h-6 bg-blue-600 rounded-full shrink-0" />
          <h2 className="text-slate-900 font-extrabold text-sm uppercase tracking-wide">
            PASO 1: IMPORTAR BARAJA DE PALABRAS
          </h2>
        </div>
        <div className="bg-slate-50 text-slate-500 font-bold text-[11px] px-3 py-1 rounded-full border border-slate-200 shrink-0">
          Total en baraja: <span className="font-extrabold text-blue-600">{loadedCardsLength}</span> tarjetas
        </div>
      </div>

      <p className="text-xs text-slate-550 leading-relaxed font-normal">
        Sube un archivo de hojas de cálculo Excel para cargar múltiples términos a la vez. El mapeo inteligente extraerá automáticamente las palabras secretas, las pistas y todas las palabras tabú para crear las tarjetas directamente.
      </p>

      {/* Main Container / Content Body */}
      {!file ? (
        // Standard File Uploader Stage
        <div 
          onClick={() => fileInputRef.current?.click()}
          onDragEnter={handleDrag}
          onDragOver={handleDrag}
          onDragLeave={handleDrag}
          onDrop={handleDrop}
          className={`border-2 border-dashed rounded-2xl p-8 flex flex-col items-center justify-center text-center cursor-pointer transition-all group shrink-0 ${
            dragActive ? 'border-emerald-500 bg-emerald-50/20' : 'border-slate-200 hover:border-blue-400 hover:bg-blue-50/20 bg-white'
          }`}
        >
          <div className="bg-blue-50 text-blue-500 p-2.5 rounded-full mb-2 group-hover:scale-110 transition-transform">
            <Upload size={20} />
          </div>
          <p className="text-slate-900 font-black text-xs uppercase tracking-wider">Arrastra tu archivo Excel aquí</p>
          <p className="text-slate-400 text-[10px] font-bold mt-1 uppercase tracking-widest">O haz clic para seleccionarlo de tu dispositivo (.xlsx, .xls)</p>
          
          <span className="mt-4 bg-slate-50 text-slate-500 font-bold text-[9px] uppercase tracking-wider py-1.5 px-3 rounded-lg border border-slate-100 flex items-center gap-1.5">
            ℹ️ Detección inteligente de palabra, estado (Inicio/Proceso...) y pistas a partir de encabezados
          </span>

          <input 
            type="file" 
            ref={fileInputRef} 
            onChange={handleFileChange} 
            accept=".xlsx, .xls" 
            className="hidden" 
          />
        </div>
      ) : (
        // Advanced Inline Column Matcher and Live Card Preview Block
        <div className="border border-slate-200/95 rounded-2xl p-4 bg-slate-50/20 flex flex-col gap-4">
          
          {/* File summary loaded row with change action */}
          <div className="flex items-center justify-between bg-white border border-slate-200/60 px-4 py-2.5 rounded-xl gap-3">
            <div className="flex items-center gap-2.5">
              <div className="bg-emerald-50 p-1.5 rounded-lg border border-emerald-100 flex items-center justify-center">
                <FileSpreadsheet size={16} className="text-emerald-600" />
              </div>
              <div className="flex flex-col">
                <span className="text-slate-900 font-extrabold text-xs leading-normal">{file.name}</span>
                <span className="text-[10px] text-slate-450 font-bold uppercase tracking-wider">
                  {(file.size / 1024).toFixed(1)} KB • {rawData.length} filas detectadas
                </span>
              </div>
            </div>
            <button 
              type="button"
              onClick={() => {
                setFile(null);
                setPreviewCards([]);
              }}
              className="text-red-500 hover:text-red-650 font-extrabold text-[11px] uppercase flex items-center gap-1.5 bg-red-50/50 hover:bg-red-50 px-3 py-1.5 rounded-xl border border-red-100/50 transition-colors"
            >
              <RefreshCw size={12} /> Cambiar archivo
            </button>
          </div>

          {/* Grid splitting Mapper Inputs (Left) and Card Visual Previews (Right) */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-5 items-stretch min-h-0">
            
            {/* Left Col: Column Mapping Configuration */}
            <div className="flex flex-col gap-4">
              <span className="text-[10px] font-black uppercase text-slate-400 tracking-wider">MAPEO DE COLUMNAS</span>
              
              {/* Palabra Secreta target */}
              <div className="space-y-1.5">
                <Label className="text-[11px] font-bold text-slate-700">Palabra Principal (Palabra Secreta):</Label>
                <Select value={mapping.wordColumn} onValueChange={(val) => handleMappingChange('wordColumn', val)}>
                  <SelectTrigger className="bg-white border-slate-200 text-slate-950 font-extrabold h-10 text-xs rounded-xl text-left">
                    <SelectValue placeholder="Seleccionar" />
                  </SelectTrigger>
                  <SelectContent className="rounded-xl font-bold text-slate-900">
                    {columns.map(col => <SelectItem key={col} value={col}>{col}</SelectItem>)}
                  </SelectContent>
                </Select>
              </div>

              {/* Pista / Clave target */}
              <div className="space-y-1.5">
                <Label className="text-[11px] font-bold text-slate-700">Columna de Pista / Clave (Aparece abajo redondeada):</Label>
                <Select value={mapping.pistaColumn || '_none'} onValueChange={(val) => handleMappingChange('pistaColumn', val === '_none' ? undefined : val)}>
                  <SelectTrigger className="bg-white border-slate-200 text-slate-950 font-extrabold h-10 text-xs rounded-xl text-left">
                    <SelectValue placeholder="Autodetectar" />
                  </SelectTrigger>
                  <SelectContent className="rounded-xl font-bold text-slate-900">
                    <SelectItem value="_none">Ninguna (Crear en blanco)</SelectItem>
                    {columns.map(col => <SelectItem key={col} value={col}>{col}</SelectItem>)}
                  </SelectContent>
                </Select>
              </div>

              {/* Level / status column */}
              <div className="space-y-1.5">
                <Label className="text-[11px] font-bold text-slate-700">Columna de Estado / Nivel de Dificultad (Rojo, Naranja, Verde, Azul claro):</Label>
                <Select value={mapping.statusColumn || '_none'} onValueChange={(val) => handleMappingChange('statusColumn', val === '_none' ? undefined : val)}>
                  <SelectTrigger className="bg-white border-slate-200 text-slate-950 font-extrabold h-10 text-xs rounded-xl text-left">
                    <SelectValue placeholder="Autodetectar" />
                  </SelectTrigger>
                  <SelectContent className="rounded-xl font-bold text-slate-900">
                    <SelectItem value="_none">Ninguna (Todos como "Inicio")</SelectItem>
                    {columns.map(col => <SelectItem key={col} value={col}>{col}</SelectItem>)}
                  </SelectContent>
                </Select>
              </div>

              {/* Custom taboo type selector (single versus multi column formats) */}
              <div className="space-y-1.5">
                <Label className="text-[11px] font-bold text-slate-700 font-bold">Formato de Palabras Tabú:</Label>
                <div className="flex gap-5 items-center pl-1 py-0.5">
                  <label className="flex items-center gap-2 cursor-pointer text-xs font-bold text-slate-700">
                    <input 
                      type="radio" 
                      name="tabuFormat" 
                      checked={!mapping.isSingleColumnForTaboo} 
                      onChange={() => handleMappingChange('isSingleColumnForTaboo', false)}
                      className="rounded-full border-slate-300 text-blue-600 focus:ring-blue-505 h-4 w-4"
                    />
                    Múltiples Columnas
                  </label>
                  <label className="flex items-center gap-2 cursor-pointer text-xs font-bold text-slate-700">
                    <input 
                      type="radio" 
                      name="tabuFormat" 
                      checked={mapping.isSingleColumnForTaboo} 
                      onChange={() => handleMappingChange('isSingleColumnForTaboo', true)}
                      className="rounded-full border-slate-300 text-blue-600 focus:ring-blue-505 h-4 w-4"
                    />
                    Columna Única (Por Separador)
                  </label>
                </div>
              </div>

              {/* Dynamic Multiple Selection of columns */}
              <div className="space-y-1.5">
                <Label className="text-[11px] font-bold text-slate-700">
                  {mapping.isSingleColumnForTaboo ? 'Selecciona Columna Única y Separador:' : 'Columnas para palabras prohibidas (Selecciona varias):'}
                </Label>
                
                {mapping.isSingleColumnForTaboo ? (
                  <div className="grid grid-cols-2 gap-2">
                    <Select value={mapping.tabooColumns[0] || ''} onValueChange={(val) => handleMappingChange('tabooColumns', [val])}>
                      <SelectTrigger className="bg-white border-slate-200 text-slate-950 font-extrabold h-10 text-xs rounded-xl text-left">
                        <SelectValue placeholder="Columna" />
                      </SelectTrigger>
                      <SelectContent className="rounded-xl font-bold text-slate-900">
                        {columns.map(col => <SelectItem key={col} value={col}>{col}</SelectItem>)}
                      </SelectContent>
                    </Select>

                    <Select value={mapping.tabooDelimiter} onValueChange={(val) => handleMappingChange('tabooDelimiter', val)}>
                      <SelectTrigger className="bg-white border-slate-200 text-slate-950 font-extrabold h-10 text-xs rounded-xl text-left">
                        <SelectValue placeholder="Separador" />
                      </SelectTrigger>
                      <SelectContent className="rounded-xl font-bold text-slate-900">
                        <SelectItem value=",">Coma ( , )</SelectItem>
                        <SelectItem value=";">Punto y coma ( ; )</SelectItem>
                        <SelectItem value="/">Barra inclinada ( / )</SelectItem>
                        <SelectItem value="\n">Salto de línea</SelectItem>
                      </SelectContent>
                    </Select>
                  </div>
                ) : (
                  <div className="border border-slate-200 rounded-xl h-[160px] overflow-y-auto bg-white p-2.5 space-y-1 custom-scrollbar shadow-inner">
                    {columns.map((col) => {
                      const isChecked = mapping.tabooColumns.includes(col);
                      return (
                        <label key={col} className="flex items-center gap-3 cursor-pointer text-xs font-bold text-slate-700 py-1 px-2 hover:bg-slate-50 rounded-lg select-none transition-all">
                          <input 
                            type="checkbox"
                            checked={isChecked}
                            onChange={(e) => {
                              let updatedTaboos = [...mapping.tabooColumns];
                              if (e.target.checked) {
                                if (!updatedTaboos.includes(col)) updatedTaboos.push(col);
                              } else {
                                updatedTaboos = updatedTaboos.filter(c => c !== col);
                              }
                              handleMappingChange('tabooColumns', updatedTaboos);
                            }}
                            className="rounded border-slate-300 text-blue-600 focus:ring-blue-500 h-4 w-4"
                          />
                          <span className="truncate">{col}</span>
                        </label>
                      );
                    })}
                  </div>
                )}
              </div>
            </div>

            {/* Right Col: PREVISUALIZACIÓN DEL MAPEO */}
            <div className="flex flex-col bg-[#f8fafc] border border-slate-200/50 rounded-2xl p-4 min-h-0 gap-3 justify-between shadow-inner">
              <span className="text-[10px] font-black uppercase text-slate-400 tracking-wider">PREVISUALIZACIÓN DEL MAPEO</span>
              
              <div className="flex-1 overflow-y-auto space-y-3.5 pr-1 max-h-[300px] min-h-[220px] custom-scrollbar">
                {previewCards.length === 0 ? (
                  <div className="flex flex-col items-center justify-center text-center h-full p-6 text-slate-400 font-extrabold uppercase text-[10px] tracking-widest gap-2">
                    <AlertCircle size={20} className="text-slate-300" />
                    <span>Configura mapeo de Palabra Secreta para ver los naipes</span>
                  </div>
                ) : (
                  previewCards.map((card, idx) => {
                    const getStatusColor = (status: CardStatus) => {
                      switch (status) {
                        case 'Inicio': return 'bg-[#dc2626]';
                        case 'Proceso': return 'bg-[#f97316]';
                        case 'Logrado': return 'bg-[#10b981]';
                        case 'Destacado': return 'bg-[#0ea5e9]';
                        default: return 'bg-[#dc2626]';
                      }
                    };
                    const colorBg = getStatusColor(card.nivelLogro);

                    return (
                      <div key={idx} className="bg-white border border-slate-200/40 rounded-xl p-3 flex flex-col items-start gap-1.5 shadow-sm select-none">
                        <div className="flex">
                          <span className={`${colorBg} text-white font-extrabold text-[10px] px-2.5 py-0.5 rounded uppercase tracking-wider`}>
                            {card.secreto || 'SIN PALABRA'}
                          </span>
                        </div>
                        <p className="text-[11px] text-slate-600 font-semibold leading-none mt-1">
                          <strong className="text-slate-400 font-extrabold uppercase text-[10px]">Tabú:</strong>{' '}
                          {card.tabu.filter(Boolean).map(t => t.toUpperCase()).join(', ') || 'NINGUNA'}
                        </p>
                        {card.pista && (
                          <p className="text-[10px] text-slate-500 font-medium italic mt-0.5 leading-relaxed">
                            pista: <span className="text-slate-700">{card.pista}</span>
                          </p>
                        )}
                      </div>
                    );
                  })
                )}
              </div>

              {/* Confirmation row at the bottom of preview panel */}
              <div className="flex items-center justify-between border-t border-slate-200/60 pt-3.5 mt-1 shrink-0">
                <span className="text-xs font-bold text-slate-600">
                  Total: <strong className="font-mono text-sm text-[#4f46e5]/90 font-black">{previewCards.length}</strong> tarjetas
                </span>
                <Button 
                  onClick={handleApplyImport}
                  disabled={previewCards.length === 0}
                  className="bg-[#4f46e5] hover:bg-[#4338ca] disabled:opacity-50 text-white font-extrabold text-xs uppercase tracking-wider h-10 px-5 rounded-xl shadow-md transition-all flex items-center gap-2"
                >
                  <Check size={14} /> Cargar {previewCards.length} Tarjetas
                </Button>
              </div>
            </div>

          </div>
        </div>
      )}

      {/* Under deck management footer action buttons (Cargar Ejemplo / Vaciar Baraja) */}
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 justify-stretch mt-1 shrink-0">
        <Button 
          variant="outline"
          onClick={onResetSet}
          className="rounded-xl border-slate-200 text-slate-700 font-extrabold text-xs uppercase tracking-wider h-11 transition-all flex items-center justify-center gap-2 bg-slate-50/50 hover:bg-slate-100"
        >
          <RefreshCw size={14} className="text-slate-400 shrink-0" />
          Cargar set de ejemplo (8 cartas)
        </Button>
        <Button 
          variant="outline"
          onClick={onClearSet}
          className="rounded-xl border-red-200 text-red-650 hover:text-red-700 font-extrabold text-xs uppercase tracking-wider h-11 transition-all flex items-center justify-center gap-2 bg-red-50/10 hover:bg-red-50/40 hover:border-red-300"
        >
          <Trash2 size={14} className="text-red-400 shrink-0" />
          Vaciar baraja actual
        </Button>
      </div>
    </div>
  );
};
