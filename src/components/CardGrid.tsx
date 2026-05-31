import React, { useState } from 'react';
import { 
  Search, 
  Grid, 
  List, 
  Plus, 
  Edit3, 
  Trash2, 
  Check, 
  Sparkles,
  RefreshCw,
  X,
  PlusCircle,
  AlertCircle,
  HelpCircle,
  BookOpen
} from 'lucide-react';
import { CardData, CardStatus } from '../types';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Badge } from '@/components/ui/badge';
import { toast } from 'sonner';

interface CardGridProps {
  cards: CardData[];
  onCardsChange: (updatedCards: CardData[]) => void;
  onGenerateAIAll: () => void;
  isGeneratingAI: boolean;
}

export const CardGrid: React.FC<CardGridProps> = ({ 
  cards, 
  onCardsChange, 
  onGenerateAIAll,
  isGeneratingAI 
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [statusFilter, setStatusFilter] = useState<string>('all');
  const [viewMode, setViewMode] = useState<'grid' | 'list'>('grid');
  const [gridUiColumns, setGridUiColumns] = useState<number>(4);
  
  // CRUD editing states
  const [editingCard, setEditingCard] = useState<CardData | null>(null);
  const [isAddingCard, setIsAddingCard] = useState(false);
  const [newCardData, setNewCardData] = useState<Partial<CardData>>({
    secreto: '',
    tabu: ['', '', '', ''],
    nivelLogro: 'Inicio',
    categoria: 'General',
    pista: ''
  });

  // Level config values
  const STATUS_COLORS: { [key in CardStatus]: { bg: string, text: string, border: string, headerBg: string, cardBorder: string } } = {
    Inicio: { 
      bg: 'bg-red-50 border-red-200', 
      text: 'text-red-600', 
      border: 'border-red-500/30', 
      headerBg: 'bg-red-500',
      cardBorder: 'border-red-200 hover:border-red-500'
    },
    Proceso: { 
      bg: 'bg-orange-50 border-orange-200', 
      text: 'text-orange-600', 
      border: 'border-orange-500/30', 
      headerBg: 'bg-orange-500',
      cardBorder: 'border-orange-200 hover:border-orange-500'
    },
    Logrado: { 
      bg: 'bg-green-50 border-green-200', 
      text: 'text-green-600', 
      border: 'border-green-500/30', 
      headerBg: 'bg-green-500',
      cardBorder: 'border-green-200 hover:border-green-500'
    },
    Destacado: { 
      bg: 'bg-sky-50 border-sky-200', 
      text: 'text-sky-600', 
      border: 'border-sky-500/30', 
      headerBg: 'bg-sky-500',
      cardBorder: 'border-sky-200 hover:border-sky-500'
    }
  };

  // Get dynamic numeric counters
  const getCounts = () => {
    const counts = { Inicio: 0, Proceso: 0, Logrado: 0, Destacado: 0 };
    cards.forEach(c => {
      const lvl = (c.nivelLogro || 'Inicio') as CardStatus;
      if (counts[lvl] !== undefined) counts[lvl]++;
    });
    return counts;
  };

  const counts = getCounts();

  // Filter logic
  const filteredCards = cards.filter(card => {
    const sTerm = searchTerm.toLowerCase().trim();
    const wordMatches = card.secreto.toLowerCase().includes(sTerm);
    const hintMatches = card.pista ? card.pista.toLowerCase().includes(sTerm) : false;
    const catMatches = card.categoria.toLowerCase().includes(sTerm);
    const textMatches = sTerm === '' || wordMatches || hintMatches || catMatches;
    
    const levelMatches = statusFilter === 'all' || card.nivelLogro === statusFilter;
    
    return textMatches && levelMatches;
  });

  // Actions
  const handleEditClick = (card: CardData, e?: React.MouseEvent) => {
    if (e) e.stopPropagation();
    // Fill up safe values with clones
    setEditingCard({
      ...card,
      tabu: [...card.tabu]
    });
  };

  const handleUpdateCard = () => {
    if (!editingCard) return;
    if (!editingCard.secreto.trim()) {
      toast.error("La palabra principal no puede estar vacía.");
      return;
    }
    if (editingCard.tabu.some(t => !t.trim())) {
      toast.error("Por favor completa las 4 palabras prohibidas.");
      return;
    }

    const next = cards.map(c => c.id === editingCard.id ? { 
      ...editingCard, 
      word: editingCard.secreto, 
      status: editingCard.nivelLogro as CardStatus,
      tabooWords: editingCard.tabu 
    } : c);

    onCardsChange(next);
    setEditingCard(null);
    toast.success("¡Tarjeta actualizada en caliente!");
  };

  const handleDeleteCard = (cardId: string, e?: React.MouseEvent) => {
    if (e) e.stopPropagation();
    const next = cards.filter(c => c.id !== cardId);
    onCardsChange(next);
    toast.success("Tarjeta eliminada.");
  };

  const handleCreateCard = () => {
    const { secreto, tabu, nivelLogro, categoria, pista } = newCardData;
    if (!secreto?.trim()) {
      toast.error("Por favor ingresa la palabra principal.");
      return;
    }
    if (!tabu || tabu.some(t => !t.trim())) {
      toast.error("Por favor completa las 4 palabras prohibidas.");
      return;
    }

    const nextCard: CardData = {
      id: `manual-${Date.now()}`,
      secreto: secreto.trim(),
      tabu: tabu.map(t => t.trim()),
      nivelLogro: (nivelLogro || 'Inicio') as CardStatus,
      categoria: (categoria || 'General').trim(),
      pista: (pista || '').trim(),
      // Backward compatibility aliases
      word: secreto.trim(),
      status: (nivelLogro || 'Inicio') as CardStatus,
      tabooWords: tabu.map(t => t.trim())
    };

    onCardsChange([nextCard, ...cards]);
    setIsAddingCard(false);
    setNewCardData({
      secreto: '',
      tabu: ['', '', '', ''],
      nivelLogro: 'Inicio',
      categoria: 'General',
      pista: ''
    });
    toast.success("¡Tarjeta creada y añadida con éxito!");
  };

  return (
    <div className="flex flex-col h-full min-h-0 bg-white border border-slate-100 rounded-[2rem] p-6 shadow-xl gap-4">
      
      {/* Paso 2 Header Section */}
      <div className="flex flex-col md:flex-row items-start md:items-center justify-between border-b border-slate-100 pb-3 shrink-0 gap-2">
        <div>
          <h3 className="text-sm font-black text-slate-900 uppercase tracking-widest flex items-center gap-2">
            🎨 PASO 2: EDITA TUS TARJETAS EN TIEMPO REAL
          </h3>
          <p className="text-[11px] text-slate-400 font-bold uppercase tracking-wider mt-0.5">
            Haz clic en el lápiz o la propia tarjeta para editar conceptos, palabras tabú y pistas en caliente.
          </p>
        </div>
      </div>

      {/* Search and view toggle toolbar */}
      <div className="flex flex-col lg:flex-row items-stretch lg:items-center justify-between gap-3 shrink-0">
        
        {/* Search Input */}
        <div className="relative flex-1">
          <Search className="absolute left-3.5 top-3 text-slate-400" size={16} />
          <Input 
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            placeholder="Buscar por palabra, pista o tabú..."
            className="pl-10 h-11 w-full bg-slate-50 border-slate-200 text-slate-900 rounded-xl font-bold font-sans shadow-inner focus-visible:ring-blue-500 ease-in-out duration-155 transition-all"
            style={{ fontSize: '13px' }}
          />
        </div>

        {/* Toolbar Controls Right Side */}
        <div className="flex flex-wrap items-center gap-2">
          {/* Level Filter Dropdown */}
          <Select value={statusFilter} onValueChange={setStatusFilter}>
            <SelectTrigger className="bg-slate-50 border-slate-200 text-slate-950 font-black rounded-xl h-11 px-4 w-44 shadow-inner text-left text-xs shrink-0">
              <SelectValue placeholder="Filtrar por nivel" />
            </SelectTrigger>
            <SelectContent className="rounded-xl font-black text-slate-950">
              <SelectItem value="all">Filtro: Todos ({cards.length})</SelectItem>
              <SelectItem value="Inicio">🔴 Inicio ({counts.Inicio})</SelectItem>
              <SelectItem value="Proceso">🟠 Proceso ({counts.Proceso})</SelectItem>
              <SelectItem value="Logrado">🟢 Logrado ({counts.Logrado})</SelectItem>
              <SelectItem value="Destacado">🔵 Destacado ({counts.Destacado})</SelectItem>
            </SelectContent>
          </Select>

          {/* Grid Columns Width Toggle: Grid UI selector */}
          <div className="flex bg-slate-100 rounded-xl p-0.5 border border-slate-200 h-11 items-center shrink-0">
            <span className="text-[10px] font-black uppercase text-slate-400 px-3 select-none">Grid UI</span>
            <div className="flex gap-0.5 bg-slate-200/50 p-0.5 rounded-lg border border-slate-100">
              {[2, 3, 4].map((cols) => (
                <button
                  key={cols}
                  onClick={() => {
                    setViewMode('grid');
                    setGridUiColumns(cols);
                  }}
                  className={`px-3 py-1 text-[11px] font-black rounded-md transition-all ${
                    viewMode === 'grid' && gridUiColumns === cols
                      ? 'bg-white text-blue-600 shadow-sm'
                      : 'text-slate-500 hover:text-slate-900'
                  }`}
                >
                  {cols} Col
                </button>
              ))}
            </div>
          </div>

          {/* View Mode Toggle: Grid/List Toggle icon */}
          <div className="flex bg-slate-100 rounded-xl p-0.5 border border-slate-200 h-11 items-center justify-center shrink-0">
            <Button 
              variant="ghost" 
              size="icon" 
              onClick={() => setViewMode('grid')}
              className={`rounded-lg h-10 w-10 p-0 ${viewMode === 'grid' ? 'bg-white shadow-md text-slate-950' : 'text-slate-400 hover:text-slate-600'}`}
            >
              <Grid size={16} />
            </Button>
            <Button 
              variant="ghost" 
              size="icon" 
              onClick={() => setViewMode('list')}
              className={`rounded-lg h-10 w-10 p-0 ${viewMode === 'list' ? 'bg-white shadow-md text-slate-950' : 'text-slate-400 hover:text-slate-600'}`}
            >
              <List size={16} />
            </Button>
          </div>

          {/* Add Manual Card Button */}
          <Button 
            onClick={() => setIsAddingCard(true)}
            className="bg-emerald-500 hover:bg-emerald-600 font-extrabold text-white text-xs px-5 h-11 rounded-xl shadow-md flex items-center gap-1.5 uppercase tracking-wider transition-all"
          >
            <Plus size={16} />
            <span>+ Crear Tarjeta</span>
          </Button>
        </div>
      </div>

      {/* Dynamic Filter Bento badges aligned in a clean row */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-2 shrink-0">
        {(['Inicio', 'Proceso', 'Logrado', 'Destacado'] as CardStatus[]).map((lvl) => {
          const cfg = STATUS_COLORS[lvl];
          const isActive = statusFilter === lvl;
          return (
            <button 
              key={lvl}
              onClick={() => setStatusFilter(isActive ? 'all' : lvl)}
              className={`flex items-center justify-between p-2.5 px-4 rounded-xl border text-left transition-all ${
                isActive 
                  ? `${cfg.bg} border-slate-900 ring-2 ring-slate-900 shadow-md` 
                  : 'bg-slate-50 hover:bg-slate-100 border-slate-200/60 text-slate-700'
              }`}
            >
              <div className="flex flex-col">
                <span className="text-[8px] font-black uppercase tracking-widest text-slate-400">Nivel</span>
                <span className={`text-xs font-black uppercase tracking-tighter ${cfg.text}`}>{lvl}</span>
              </div>
              <Badge className="font-extrabold font-mono text-[11px] h-6 min-w-6 flex items-center justify-center rounded-lg pb-0.5" variant="secondary">
                {counts[lvl]}
              </Badge>
            </button>
          );
        })}
      </div>

      {/* Main Container Content list or dynamic card grid (Scrollable section ONLY) */}
      <div className="flex-1 min-h-[350px] max-h-[580px] overflow-y-auto pr-2 custom-scrollbar border border-slate-100/60 rounded-2xl bg-slate-50/20 p-2">
        {filteredCards.length === 0 ? (
          <div className="h-full min-h-[300px] flex flex-col justify-center items-center text-center p-8 border-2 border-dashed border-slate-100 rounded-3xl bg-slate-50/50">
            <AlertCircle size={36} className="text-slate-300 mb-2" />
            <p className="text-slate-800 font-black uppercase tracking-tight">No se encontraron tarjetas</p>
            <p className="text-slate-400 text-[10px] font-bold uppercase tracking-widest mt-1">Prueba quitando filtros o realiza una búsqueda distinta</p>
          </div>
        ) : viewMode === 'grid' ? (
          // Grid View: simulating exactly classic game boards with customizable column layout
          <div className={`grid gap-4 ${
            gridUiColumns === 2 ? "grid-cols-1 sm:grid-cols-2" :
            gridUiColumns === 3 ? "grid-cols-1 sm:grid-cols-2 md:grid-cols-3" :
            "grid-cols-1 sm:grid-cols-2 md:grid-cols-3 xl:grid-cols-4"
          }`}>
            {filteredCards.map((card) => {
              const info = STATUS_COLORS[card.nivelLogro as CardStatus] || STATUS_COLORS.Inicio;
              return (
                <div 
                  key={card.id}
                  onClick={() => handleEditClick(card)}
                  className={`flex flex-col bg-white rounded-[24px] overflow-hidden shadow-md hover:shadow-2xl hover:-translate-y-1 transition-all duration-300 cursor-pointer border-[3px] ${info.cardBorder} group selection:bg-none relative`}
                  style={{ 
                    aspectRatio: '3 / 4.7'
                  }}
                >
                  {/* Top solid colored header banner */}
                  <div className={`py-4 px-2 text-center relative overflow-hidden shrink-0 ${info.headerBg} text-white flex flex-col justify-center items-center gap-1`}>
                    <div className="flex justify-between items-center w-full px-2 text-[9px] font-black uppercase tracking-wider text-white/85 shrink-0">
                      <span className="truncate max-w-[65%]">
                        {card.categoria || 'General'}
                      </span>
                      <span>
                        {card.nivelLogro}
                      </span>
                    </div>

                    {/* Word Secret */}
                    <h3 className="text-base md:text-lg font-black uppercase tracking-wide text-white line-clamp-1 pt-1 drop-shadow-sm select-all">
                      {card.secreto}
                    </h3>
                  </div>

                  {/* Intersect Contrasting Taboo words list panel */}
                  <div className="flex-1 min-h-0 flex flex-col justify-center px-4 py-4 bg-white">
                    <div className="w-full flex flex-col items-center justify-center gap-1.5 flex-1 py-1">
                      {card.tabu.map((word, i) => (
                        <React.Fragment key={i}>
                          <span className="text-xs md:text-sm font-black uppercase tracking-widest text-slate-850 truncate max-w-full text-center">
                            {word}
                          </span>
                          {i < card.tabu.length - 1 && (
                            <div className="w-8 h-[1px] bg-slate-200 shrink-0 rounded" />
                          )}
                        </React.Fragment>
                      ))}
                    </div>
                  </div>

                  {/* Elegant Clue footer aligned perfectly separated by dashed border */}
                  <div className="border-t border-dashed border-slate-200 px-4 py-3 shrink-0 flex flex-col items-center justify-center bg-slate-50/50">
                    <p className="text-[10px] md:text-[11px] leading-relaxed italic text-slate-500 font-semibold lowercase line-clamp-2 text-center select-none">
                      <span className="font-extrabold pr-1.5 text-indigo-500 uppercase tracking-widest text-[9px] not-italic">pista:</span>
                      {card.pista || 'Sin ayuda registrada.'}
                    </p>
                  </div>

                  {/* hover button edit overlay */}
                  <div className="absolute right-2.5 top-2.5 opacity-0 group-hover:opacity-100 transition-opacity">
                    <Button 
                      size="icon" 
                      variant="secondary"
                      onClick={(e) => handleEditClick(card, e)}
                      className="h-7 w-7 rounded-lg shadow-md bg-white hover:bg-slate-100 text-slate-700"
                    >
                      <Edit3 size={12} />
                    </Button>
                  </div>
                </div>
              );
            })}
          </div>
        ) : (
          // Compact List View Table: Ideal for bulk review or extremely fast manipulation
          <div className="border border-slate-100 rounded-2xl overflow-hidden bg-white/50">
            <table className="w-full text-left text-xs border-collapse">
              <thead className="bg-slate-100/60 text-[10px] font-black uppercase tracking-wider text-slate-400">
                <tr>
                  <th className="p-3">Palabra Secreta</th>
                  <th className="p-3">Palabras Prohibidas (Tabú)</th>
                  <th className="p-3">Pista</th>
                  <th className="p-3">Categoría</th>
                  <th className="p-3">Nivel</th>
                  <th className="p-3 text-right">Acciones</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 font-extrabold uppercase text-slate-700">
                {filteredCards.map((card) => {
                  const info = STATUS_COLORS[card.nivelLogro as CardStatus] || STATUS_COLORS.Inicio;
                  return (
                    <tr key={card.id} className="hover:bg-slate-50/50 transition-colors cursor-pointer group" onClick={() => handleEditClick(card)}>
                      <td className="p-3 font-black text-slate-900 tracking-tight">{card.secreto}</td>
                      <td className="p-3">
                        <div className="flex flex-wrap gap-1">
                          {card.tabu.map((t, i) => (
                            <span key={i} className="bg-slate-100/80 px-2 py-0.5 rounded text-[10px] text-slate-600 font-semibold uppercase">{t}</span>
                          ))}
                        </div>
                      </td>
                      <td className="p-3 max-w-[220px] truncate italic lowercase text-slate-500 font-bold">{card.pista || <span className="text-slate-300">Ninguna</span>}</td>
                      <td className="p-3 font-bold text-[10px] text-slate-400">{card.categoria}</td>
                      <td className="p-3">
                        <span className={`px-2.5 py-0.5 rounded-full text-[9px] font-black border uppercase ${
                          card.nivelLogro === 'Inicio' ? 'bg-red-50 text-red-600 border-red-100' :
                          card.nivelLogro === 'Proceso' ? 'bg-orange-50 text-orange-600 border-orange-100' :
                          card.nivelLogro === 'Logrado' ? 'bg-emerald-50 text-emerald-600 border-emerald-100' :
                          'bg-sky-50 text-sky-600 border-sky-100'
                        }`}>
                          {card.nivelLogro}
                        </span>
                      </td>
                      <td className="p-3 text-right" onClick={(e) => e.stopPropagation()}>
                        <div className="flex items-center justify-end gap-1.5 opacity-60 group-hover:opacity-100 transition-opacity">
                          <Button 
                            variant="ghost" 
                            size="icon" 
                            onClick={() => handleEditClick(card)}
                            className="h-7 w-7 rounded-lg text-slate-500 hover:bg-slate-100 hover:text-slate-850"
                          >
                            <Edit3 size={13} />
                          </Button>
                          <Button 
                            variant="ghost" 
                            size="icon" 
                            onClick={() => handleDeleteCard(card.id)}
                            className="h-7 w-7 rounded-lg text-rose-500 hover:bg-rose-50 hover:text-rose-700"
                          >
                            <Trash2 size={13} />
                          </Button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* MODAL 1: Inline Hot Editing Card detail */}
      {editingCard && (
        <div className="fixed inset-0 z-50 bg-slate-950/40 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl border border-slate-100 shadow-2xl w-full max-w-lg overflow-hidden flex flex-col p-6 animate-in zoom-in-95 duration-200">
            <div className="flex justify-between items-center pb-3 border-b border-slate-50">
              <h3 className="text-sm font-black text-slate-900 uppercase tracking-tight flex items-center gap-1.5">
                <Edit3 className="text-blue-500" size={16} /> Editar Tarjeta en Caliente
              </h3>
              <Button variant="ghost" size="icon" className="h-7 w-7 rounded-full text-slate-400" onClick={() => setEditingCard(null)}>
                <X size={16} />
              </Button>
            </div>

            <div className="space-y-3.5 py-4 flex-1">
              {/* Secreto Word */}
              <div className="space-y-1">
                <Label className="text-[10px] font-black text-slate-500 uppercase tracking-widest pl-1">Concepto Principal</Label>
                <Input 
                  value={editingCard.secreto} 
                  onChange={(e) => setEditingCard({ ...editingCard, secreto: e.target.value })}
                  className="bg-slate-50 font-black text-slate-900 rounded-xl uppercase tracking-tighter"
                />
              </div>

              {/* Categoría & Nivel */}
              <div className="grid grid-cols-2 gap-3.5">
                <div className="space-y-1">
                  <Label className="text-[10px] font-black text-slate-500 uppercase tracking-widest pl-1">Categoría</Label>
                  <Input 
                    value={editingCard.categoria} 
                    onChange={(e) => setEditingCard({ ...editingCard, categoria: e.target.value })}
                    className="bg-slate-50 font-black text-slate-900 rounded-xl"
                  />
                </div>
                <div className="space-y-1">
                  <Label className="text-[10px] font-black text-slate-500 uppercase tracking-widest pl-1">Nivel de Logro</Label>
                  <Select 
                    value={editingCard.nivelLogro} 
                    onValueChange={(val) => setEditingCard({ ...editingCard, nivelLogro: val as CardStatus })}
                  >
                    <SelectTrigger className="bg-slate-50 border-slate-200 text-slate-950 font-black rounded-xl">
                      <SelectValue placeholder="Escoger Nivel" />
                    </SelectTrigger>
                    <SelectContent className="rounded-xl font-black text-slate-900">
                      <SelectItem value="Inicio">🔴 Inicio</SelectItem>
                      <SelectItem value="Proceso">🟠 Proceso</SelectItem>
                      <SelectItem value="Logrado">🟢 Logrado</SelectItem>
                      <SelectItem value="Destacado">🔵 Destacado</SelectItem>
                    </SelectContent>
                  </Select>
                </div>
              </div>

              {/* Taboo words grid */}
              <div className="space-y-2 pt-2 border-t border-slate-50">
                <Label className="text-[10px] font-black text-slate-500 uppercase tracking-widest pl-1">Palabras Prohibidas (Tabú)</Label>
                <div className="grid grid-cols-2 gap-2">
                  {[0, 1, 2, 3].map((index) => (
                    <div key={index} className="relative">
                      <Input 
                        value={editingCard.tabu[index] || ''} 
                        onChange={(e) => {
                          const nextTaboo = [...editingCard.tabu];
                          nextTaboo[index] = e.target.value;
                          setEditingCard({ ...editingCard, tabu: nextTaboo });
                        }}
                        className="bg-slate-50 font-extrabold text-slate-700 pl-8 rounded-xl uppercase tracking-tag"
                        placeholder={`Palabra Tabú ${index + 1}`}
                      />
                      <span className="absolute left-3 top-2.5 text-[9px] font-bold text-rose-500/60">#{index + 1}</span>
                    </div>
                  ))}
                </div>
              </div>

              {/* Socratic Hint */}
              <div className="space-y-1 pt-2 border-t border-slate-50">
                <Label className="text-[10px] font-black text-slate-500 tracking-widest uppercase pl-1">Explicación o Pista Socrática</Label>
                <Input 
                  value={editingCard.pista || ''} 
                  onChange={(e) => setEditingCard({ ...editingCard, pista: e.target.value })}
                  placeholder="Escribe la pista para el pie de tarjeta..."
                  className="bg-slate-50 text-slate-600 font-extrabold rounded-xl"
                />
              </div>
            </div>

            <div className="flex items-center gap-2 pt-3 border-t border-slate-50 shrink-0">
              <Button 
                variant="destructive"
                onClick={() => {
                  handleDeleteCard(editingCard.id);
                  setEditingCard(null);
                }}
                className="font-black uppercase tracking-wider text-xs px-4 h-10 rounded-xl"
              >
                Eliminar Tarjeta
              </Button>
              <Button 
                onClick={handleUpdateCard}
                className="flex-1 bg-blue-600 hover:bg-blue-700 text-white font-black uppercase text-xs tracking-wider h-10 rounded-xl"
              >
                Guardar Cambios
              </Button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 2: Create card completely */}
      {isAddingCard && (
        <div className="fixed inset-0 z-50 bg-slate-950/40 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl border border-slate-100 shadow-2xl w-full max-w-lg overflow-hidden flex flex-col p-6 animate-in zoom-in-95 duration-200">
            <div className="flex justify-between items-center pb-3 border-b border-slate-50">
              <h3 className="text-sm font-black text-slate-900 uppercase tracking-tight flex items-center gap-1.5">
                <PlusCircle className="text-blue-600" size={16} /> Crear Nueva Tarjeta Tabú
              </h3>
              <Button variant="ghost" size="icon" className="h-7 w-7 rounded-full text-slate-400" onClick={() => setIsAddingCard(false)}>
                <X size={16} />
              </Button>
            </div>

            <div className="space-y-3.5 py-4 flex-1">
              {/* Concept secrecy */}
              <div className="space-y-1">
                <Label className="text-[10px] font-black text-slate-500 uppercase tracking-widest pl-1">Concepto Secreto</Label>
                <Input 
                  value={newCardData.secreto || ''} 
                  onChange={(e) => setNewCardData({ ...newCardData, secreto: e.target.value })}
                  placeholder="Ej: PLANETA TIERRA"
                  className="bg-slate-50 font-black text-slate-900 rounded-xl uppercase tracking-tighter"
                />
              </div>

              {/* Categoría & Nivel */}
              <div className="grid grid-cols-2 gap-3.5">
                <div className="space-y-1">
                  <Label className="text-[10px] font-black text-slate-500 uppercase tracking-widest pl-1">Categoría</Label>
                  <Input 
                    value={newCardData.categoria || ''} 
                    onChange={(e) => setNewCardData({ ...newCardData, categoria: e.target.value })}
                    placeholder="General"
                    className="bg-slate-50 font-black text-slate-900 rounded-xl"
                  />
                </div>
                <div className="space-y-1">
                  <Label className="text-[10px] font-black text-slate-500 uppercase tracking-widest pl-1">Nivel de Logro</Label>
                  <Select 
                    value={newCardData.nivelLogro} 
                    onValueChange={(val) => setNewCardData({ ...newCardData, nivelLogro: val })}
                  >
                    <SelectTrigger className="bg-slate-50 border-slate-200 text-slate-950 font-black rounded-xl">
                      <SelectValue placeholder="Escoger Nivel" />
                    </SelectTrigger>
                    <SelectContent className="rounded-xl font-black text-slate-900">
                      <SelectItem value="Inicio">🔴 Inicio</SelectItem>
                      <SelectItem value="Proceso">🟠 Proceso</SelectItem>
                      <SelectItem value="Logrado">🟢 Logrado</SelectItem>
                      <SelectItem value="Destacado">🔵 Destacado</SelectItem>
                    </SelectContent>
                  </Select>
                </div>
              </div>

              {/* Taboo words */}
              <div className="space-y-2 pt-2 border-t border-slate-50">
                <Label className="text-[10px] font-black text-slate-500 uppercase tracking-widest pl-1">Palabras Prohibidas (Tabú)</Label>
                <div className="grid grid-cols-2 gap-2">
                  {[0, 1, 2, 3].map((index) => (
                    <div key={index} className="relative">
                      <Input 
                        value={newCardData.tabu ? newCardData.tabu[index] || '' : ''} 
                        onChange={(e) => {
                          const tabList = [...(newCardData.tabu || ['', '', '', ''])];
                          tabList[index] = e.target.value;
                          setNewCardData({ ...newCardData, tabu: tabList });
                        }}
                        className="bg-slate-50 font-extrabold text-slate-700 pl-8 rounded-xl uppercase tracking-tag"
                        placeholder={`Palabra Tabú ${index + 1}`}
                      />
                      <span className="absolute left-3 top-2.5 text-[9px] font-bold text-rose-500/60">#{index + 1}</span>
                    </div>
                  ))}
                </div>
              </div>

              {/* Socratic Clue */}
              <div className="space-y-1 pt-2 border-t border-slate-50">
                <Label className="text-[10px] font-black text-slate-500 tracking-widest uppercase pl-1">Explicación o Pista Socrática</Label>
                <Input 
                  value={newCardData.pista || ''} 
                  onChange={(e) => setNewCardData({ ...newCardData, pista: e.target.value })}
                  placeholder="Ej: Esfera celeste donde habitamos con un único satélite natural."
                  className="bg-slate-50 text-slate-600 font-extrabold rounded-xl"
                />
              </div>
            </div>

            <div className="flex items-center gap-2 pt-3 border-t border-slate-50 shrink-0">
              <Button 
                variant="outline" 
                onClick={() => setIsAddingCard(false)}
                className="w-28 font-black uppercase text-xs tracking-wider h-10 rounded-xl"
              >
                Cancelar
              </Button>
              <Button 
                onClick={handleCreateCard}
                className="flex-1 bg-emerald-500 hover:bg-emerald-600 text-white font-black uppercase text-xs tracking-wider h-10 rounded-xl shadow-lg border-b-2"
              >
                Crear Tarjeta
              </Button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
