import React, { useState, useEffect, useMemo, useRef } from 'react';
import { createPortal } from 'react-dom';
import { HexColorPicker } from 'react-colorful';
import {
  Bold, Italic, Underline, ChevronDown, AlignLeft, AlignCenter, AlignRight,
  ArrowUpDown, WrapText, DollarSign, Percent, Lightbulb, Table as TableIcon,
  ArrowRightToLine, Trash2, Wand2, Type, PaintBucket, Baseline,
  Palette, LayoutTemplate, Printer, BringToFront, SendToBack, Layers, RotateCw,
  MoveUp, MoveDown, Focus, Grip, LayoutDashboard, Image as ImageIcon,
  Shapes, Sticker, Box, Network, BarChart2, LineChart, PieChart,
  AreaChart, Activity, TrendingUp, PenTool, MousePointer2,
  Globe, Sigma, Clock, Binary, Calendar, Search, Calculator,
  MoreHorizontal, Tags, Tag, BookmarkPlus, CheckSquare, CornerUpLeft, CornerUpRight,
  Eraser, Eye, CheckCircle2, Settings2, PlaySquare, FileSignature, Hexagon,
  Minus, Square, X,
  Map, LayoutList, BarChart, PanelTop, Radar, LayoutGrid, Sun, SlidersHorizontal, TrendingDown, Filter
} from 'lucide-react';

/* ----- STATE MANAGEMENT ----- */
const cols = Array.from({length: 40}, (_, i) => {
  if (i < 26) return String.fromCharCode(65 + i);
  return 'A' + String.fromCharCode(65 + (i - 26));
});
const rows = Array.from({length: 40}, (_, i) => i + 1);

const getColIdx = (col: string) => cols.indexOf(col);
const getRange = (start: string, end: string) => {
  const startCol = start.match(/[A-Z]+/)?.[0] || 'A';
  const startRow = parseInt(start.match(/[0-9]+/)?.[0] || '1');
  const endCol = end.match(/[A-Z]+/)?.[0] || 'A';
  const endRow = parseInt(end.match(/[0-9]+/)?.[0] || '1');

  return {
    minColIdx: Math.min(getColIdx(startCol), getColIdx(endCol)),
    maxColIdx: Math.max(getColIdx(startCol), getColIdx(endCol)),
    minRow: Math.min(startRow, endRow),
    maxRow: Math.max(startRow, endRow)
  };
};

interface CellData { 
  value: string; 
  bold?: boolean; 
  italic?: boolean; 
  underline?: boolean; 
  align?: 'left' | 'center' | 'right';
  vAlign?: 'top' | 'middle' | 'bottom';
  wrapText?: boolean;
  fontFamily?: string;
  fontSize?: number;
  textColor?: string;
  fillColor?: string;
}

let sharedState = {
  activeCell: 'A1',
  selection: { start: 'A1', end: 'A1' },
  isDragging: false,
  editingCell: null as string | null,
  data: {} as Record<string, CellData>,
  showFloatingMenu: false,
  floatingPos: { top: 0, left: 0 },
  recentColors: [] as string[]
};

const listeners = new Set<() => void>();

const dispatch = (partial: Partial<typeof sharedState>) => {
  sharedState = { ...sharedState, ...partial };
  listeners.forEach(l => l());
};

const useSharedState = () => {
  const [state, setState] = useState(sharedState);
  useEffect(() => {
    const l = () => setState(sharedState);
    listeners.add(l);
    return () => { listeners.delete(l); };
  }, []);
  return state;
};

const updateSelectionStyle = (styleKey: keyof CellData, value: any, toggle: boolean = false) => {
  const { start, end } = sharedState.selection;
  const { minColIdx, maxColIdx, minRow, maxRow } = getRange(start, end);
  const newData = { ...sharedState.data };
  
  let allOn = true;
  if (toggle) {
    for (let r = minRow; r <= maxRow; r++) {
      for (let c = minColIdx; c <= maxColIdx; c++) {
        const id = `${cols[c]}${r}`;
        if (!newData[id]?.[styleKey]) {
          allOn = false;
          break;
        }
      }
    }
  }

  const newValue = toggle ? !allOn : value;

  for (let r = minRow; r <= maxRow; r++) {
    for (let c = minColIdx; c <= maxColIdx; c++) {
      const id = `${cols[c]}${r}`;
      const current = newData[id] || { value: '' };
      newData[id] = { ...current, [styleKey]: newValue };
    }
  }
  dispatch({ data: newData });
};


/* ----- UI COMPONENTS ----- */
const EditableDropdown = ({ items, value, onSelect, placeholder, className, isNumber = false }: any) => {
    const [isOpen, setIsOpen] = useState(false);
    const triggerRef = useRef<HTMLDivElement>(null);
    const [coords, setCoords] = useState({ top: 0, left: 0 });
    
    const currentItem = items.find((i: any) => i.value === value);
    const displayValue = currentItem ? currentItem.label : (value || '');
    const [inputValue, setInputValue] = useState(displayValue);

    useEffect(() => {
        setInputValue(displayValue);
    }, [value, items, displayValue]);

    const handleOpen = () => {
        if (!isOpen && triggerRef.current) {
            const rect = triggerRef.current.getBoundingClientRect();
            setCoords({ top: rect.bottom + window.scrollY, left: rect.left + window.scrollX });
        }
        setIsOpen(true);
    };

    const commitValue = () => {
        let val: any = inputValue;
        if (isNumber) {
            val = parseInt(inputValue, 10);
            if (isNaN(val) || val <= 0) {
                setInputValue(displayValue);
                return;
            }
        }
        onSelect(val);
        setIsOpen(false);
    };

    const handleKeyDown = (e: React.KeyboardEvent) => {
        if (e.key === 'Enter') {
            commitValue();
        }
    };

    return (
        <div className={`relative ${className || ''}`} ref={triggerRef}>
            <div className="flex flex-row items-center border border-slate-200 rounded-sm bg-white hover:bg-slate-50 transition-colors w-full h-[24px]">
                <input
                    type="text"
                    className="w-full px-1.5 py-0 text-[11px] font-medium text-slate-700 outline-none bg-transparent min-w-0"
                    value={inputValue}
                    onChange={(e) => setInputValue(e.target.value)}
                    onFocus={(e) => { e.target.select(); handleOpen(); }}
                    onBlur={() => {
                        setTimeout(() => commitValue(), 150);
                    }}
                    onKeyDown={handleKeyDown}
                    placeholder={placeholder}
                />
                <div 
                    className="px-1 cursor-pointer h-full flex items-center justify-center border-l border-transparent hover:border-slate-200 hover:bg-slate-100"
                    onClick={(e) => {
                        e.stopPropagation();
                        if (isOpen) setIsOpen(false); else handleOpen();
                    }}
                >
                    <ChevronDown size={14} className="text-slate-400" />
                </div>
            </div>
            {isOpen && createPortal(
                <>
                <div className="fixed inset-0 z-[100]" onClick={(e) => { e.stopPropagation(); setIsOpen(false); }} />
                <div 
                   className="absolute bg-white border border-slate-200 rounded-lg shadow-xl z-[101] min-w-max py-1 max-h-[300px] overflow-y-auto" 
                   style={{ top: coords.top + 4, left: coords.left }}
                >
                    {items.map((item: any) => (
                        <div 
                            key={item.value || item.label} 
                            onClick={(e) => { 
                                e.stopPropagation(); 
                                setInputValue(item.label);
                                onSelect(item.value); 
                                setIsOpen(false); 
                            }}
                            className={`px-4 py-1.5 hover:bg-slate-50 cursor-pointer text-sm text-slate-700 flex items-center gap-2 ${value === item.value ? 'bg-blue-50/50' : ''}`}
                        >
                            <span style={item.fontFamily ? { fontFamily: item.fontFamily } : {}}>{item.label}</span>
                        </div>
                    ))}
                </div>
                </>,
                document.body
            )}
        </div>
    );
};

const DropdownMenu = ({ trigger, items, onSelect, className }: any) => {
    const [isOpen, setIsOpen] = useState(false);
    const triggerRef = useRef<HTMLDivElement>(null);
    const [coords, setCoords] = useState({ top: 0, left: 0 });

    const handleOpen = () => {
       if (!isOpen && triggerRef.current) {
          const rect = triggerRef.current.getBoundingClientRect();
          let top = rect.bottom + window.scrollY;
          let left = rect.left + window.scrollX;
          setCoords({ top, left });
       }
       setIsOpen(!isOpen);
    };

    useEffect(() => {
        if (isOpen && triggerRef.current) {
            const rect = triggerRef.current.getBoundingClientRect();
            setCoords({ top: rect.bottom + window.scrollY, left: rect.left + window.scrollX });
        }
    }, [isOpen]);

    return (
        <div className={`relative ${className || ''}`} ref={triggerRef}>
            {React.cloneElement(trigger, { onClick: handleOpen })}
            {isOpen && createPortal(
                <>
                <div className="fixed inset-0 z-[100]" onClick={(e) => { e.stopPropagation(); setIsOpen(false); }} />
                <div 
                   className="absolute bg-white border border-slate-200 rounded-lg shadow-xl z-[101] min-w-max py-1 max-h-[300px] overflow-y-auto" 
                   style={{ top: coords.top + 4, left: coords.left }}
                >
                    {items.map((item: any) => (
                        <div 
                            key={item.value || item.label} 
                            onClick={(e) => { e.stopPropagation(); onSelect(item.value); setIsOpen(false); }}
                            className="px-4 py-1.5 hover:bg-slate-50 cursor-pointer text-sm text-slate-700 flex items-center gap-2"
                        >
                            {item.color && <div className="w-4 h-4 rounded-full border border-slate-200" style={{ backgroundColor: item.color }} />}
                            <span style={{ fontFamily: item.fontFamily }}>{item.label}</span>
                        </div>
                    ))}
                </div>
                </>,
                document.body
            )}
        </div>
    );
};

const ColorPickerMenu = ({ trigger, color, onSelect, className }: any) => {
    const [isOpen, setIsOpen] = useState(false);
    const triggerRef = useRef<HTMLDivElement>(null);
    const [coords, setCoords] = useState({ top: 0, left: 0 });
    const recentColors = useSharedState().recentColors;

    const handleSelect = (c: string) => {
        if (c !== 'transparent' && !recentColors.includes(c)) {
            const newRecent = [c, ...recentColors].slice(0, 10);
            dispatch({ recentColors: newRecent });
        }
        onSelect(c);
    };

    const handleOpen = () => {
       if (!isOpen && triggerRef.current) {
          const rect = triggerRef.current.getBoundingClientRect();
          setCoords({ top: rect.bottom + window.scrollY, left: Math.max(0, rect.left + window.scrollX - 75) }); // Offset to left if too wide
       }
       setIsOpen(!isOpen);
    };

    useEffect(() => {
        if (isOpen && triggerRef.current) {
            const rect = triggerRef.current.getBoundingClientRect();
            setCoords({ top: rect.bottom + window.scrollY, left: Math.max(0, rect.left + window.scrollX - 75) });
        }
    }, [isOpen]);

    return (
        <div className={`relative ${className || ''}`} ref={triggerRef}>
            {React.cloneElement(trigger, { onClick: handleOpen })}
            {isOpen && createPortal(
                <>
                <div className="fixed inset-0 z-[100]" onClick={(e) => { e.stopPropagation(); setIsOpen(false); }} />
                <div 
                   className="absolute bg-white border border-slate-200 rounded-lg shadow-xl z-[101] p-2 min-w-max flex flex-col gap-2" 
                   style={{ top: coords.top + 4, left: coords.left }}
                >
                    {recentColors.length > 0 && (
                        <>
                            <div className="flex flex-col gap-1">
                                <span className="text-[10px] text-slate-500 font-medium px-1">Recent Colors</span>
                                <div className="flex gap-0.5 flex-wrap w-[160px]">
                                    {recentColors.map(c => (
                                        <div 
                                            key={c} 
                                            onClick={(e) => { e.stopPropagation(); handleSelect(c); setIsOpen(false); }}
                                            className="w-5 h-5 cursor-pointer hover:scale-110 transition-transform border border-slate-200/50 rounded-sm"
                                            style={{ backgroundColor: c }} 
                                        />
                                    ))}
                                </div>
                            </div>
                            <div className="w-full h-px bg-slate-200" />
                        </>
                    )}
                    
                    <div className="flex flex-col px-1 gap-2">
                       <span className="text-[11px] text-slate-700 font-medium">Custom Color</span>
                       <HexColorPicker 
                          color={color !== 'transparent' && color ? color : '#ffffff'} 
                          onChange={(newColor) => handleSelect(newColor)} 
                          className="!w-full !h-[120px]"
                       />
                       <input 
                          type="text" 
                          value={color !== 'transparent' && color ? color : '#ffffff'}
                          onChange={(e) => handleSelect(e.target.value)}
                          className="w-full px-2 py-1 text-xs border border-slate-200 rounded"
                       />
                    </div>
                    <div className="w-full h-px bg-slate-200 my-1" />
                    <button 
                       className="text-[11px] text-slate-700 hover:bg-slate-100 py-1 px-2 text-left rounded"
                       onClick={(e) => { e.stopPropagation(); handleSelect('transparent'); setIsOpen(false); }}
                    >
                       No Fill
                    </button>
                </div>
                </>,
                document.body
            )}
        </div>
    );
};

const ColorPickerIcon = ({ icon: Icon, color, onClick, isText = false, isFill=false }: any) => (
  <div onClick={onClick} className="flex flex-col items-center justify-center p-1 rounded-sm hover:bg-slate-100 cursor-pointer h-[28px] min-w-[32px] group">
    <div className="relative flex flex-col items-center gap-[1px]">
      {isText ? (
          <span className="font-serif font-bold text-[14px] leading-none mb-[1px] text-slate-700 group-hover:text-slate-900">A</span>
      ) : (
          <Icon size={14} className={isFill ? "text-emerald-600 mb-[1px]" : "text-slate-700"} strokeWidth={1.5} />
      )}
      <div className="w-[16px] h-[3px] rounded-[1px]" style={{ backgroundColor: color || (isFill ? 'transparent' : '#000000') }} />
    </div>
  </div>
);

const ToolbarItem = ({ icon: Icon, label, active = false, iconColor = "text-slate-600", onClick, small = false }: any) => (
  <div onClick={onClick} className={`flex flex-col items-center justify-start ${small ? 'gap-0.5 min-w-[32px]' : 'gap-1 min-w-[52px]'} cursor-pointer group px-1`}>
    <div className={`p-1.5 rounded-lg transition-colors ${active ? 'bg-[#e5ecf6] text-blue-700' : 'hover:bg-slate-100 ' + iconColor}`}>
      <Icon size={small ? 16 : 18} strokeWidth={active ? 2 : 1.5} />
    </div>
    {label && (
      <span className="text-[10px] text-slate-500 font-medium whitespace-pre-line text-center leading-[1.1] opacity-90 group-hover:opacity-100 transition-opacity">
        {label}
      </span>
    )}
  </div>
);

const SmallToolbarItem = ({icon: Icon, label, iconColor = "text-slate-600", active}: any) => (
    <div className={`flex items-center gap-1.5 cursor-pointer group px-2 py-1 rounded-md transition-colors ${active ? 'bg-slate-100' : 'hover:bg-slate-50'}`}>
        <Icon size={14} className={iconColor} strokeWidth={1.5} />
        <span className="text-[10px] text-slate-600 font-medium group-hover:text-slate-900 leading-none whitespace-nowrap">
            {label}
        </span>
    </div>
);

const RibbonSection = ({title, children, className="", noBorder=false}: any) => (
    <div className={`flex flex-col justify-between items-center relative pr-4 pl-2 shrink-0 ${className}`}>
        <div className="flex gap-1 items-start h-[52px]">
            {children}
        </div>
        <span className="text-[10px] text-slate-400 font-medium tracking-wide mt-2 mb-1">{title}</span>
        {!noBorder && <div className="absolute right-0 top-1 bottom-1 w-[1px] bg-slate-200" />}
    </div>
);

const CheckboxGroup = ({title, items}: any) => (
    <div className="flex flex-col gap-1 w-[70px]">
        <div className="text-[11px] font-medium text-slate-700 mb-0.5">{title}</div>
        {items.map((item: any, i: number) => (
            <label key={i} className="flex items-center gap-1.5 cursor-pointer group">
                <div className={`w-3.5 h-3.5 rounded-sm border ${item.checked ? 'bg-blue-500 border-blue-500' : 'border-slate-300 group-hover:border-slate-400'} flex items-center justify-center transition-colors`}>
                    {item.checked && <svg width="8" height="8" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="4" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg>}
                </div>
                <span className="text-[11px] text-slate-600 group-hover:text-slate-800">{item.label}</span>
            </label>
        ))}
    </div>
);

const DropdownField = ({label, value}: any) => (
    <div className="flex items-center gap-2">
        <span className="text-[11px] text-slate-500 w-[40px] text-right">{label}:</span>
        <button className="flex items-center justify-between border border-slate-200 rounded-md px-2 py-0.5 min-w-[70px] bg-white hover:bg-slate-50 transition-colors">
            <span className="text-[11px] text-slate-700">{value}</span>
            <ChevronDown size={12} className="text-slate-400" />
        </button>
    </div>
);

const MenuTabs = ({ activeTab, setActiveTab }: any) => {
    const tabs = ['File', 'Home', 'Insert', 'Page Layout', 'Formulas', 'Help'];
    return (
        <div className="bg-white/70 backdrop-blur-md px-2 py-3 rounded-[24px] shadow-[0_2px_15px_-4px_rgba(0,0,0,0.05)] border border-white flex items-center justify-between text-sm shrink-0">
           <div className="flex items-center flex-1 overflow-x-auto [&::-webkit-scrollbar]:hidden [-ms-overflow-style:none] [scrollbar-width:none]">
            {tabs.map((tab) => {
                const isActive = tab === activeTab;
                return (
                    <button
                        key={tab}
                        onClick={() => setActiveTab(tab)}
                        className={`relative px-4 sm:px-5 py-1 transition-colors outline-none shrink-0 ${
                        isActive ? 'text-slate-900 font-medium' : 'text-slate-600 hover:text-slate-900'
                        }`}
                    >
                        {tab}
                        {isActive && (
                            <div className="absolute -bottom-1.5 left-1/2 -translate-x-1/2 w-8 h-[3px] bg-gradient-to-r from-teal-400 to-blue-400 rounded-t-full" />
                        )}
                    </button>
                );
            })}
           </div>
           <div className="hidden sm:flex items-center gap-1 pl-4 pr-1 border-l border-slate-200/50 text-slate-500 ml-2">
               <div className="w-8 h-8 flex items-center justify-center rounded-lg hover:bg-slate-200/80 cursor-pointer transition-colors text-slate-500 hover:text-slate-800">
                   <Minus size={16} />
               </div>
               <div className="w-8 h-8 flex items-center justify-center rounded-lg hover:bg-slate-200/80 cursor-pointer transition-colors text-slate-500 hover:text-slate-800">
                   <Square size={13} />
               </div>
               <div className="w-8 h-8 flex items-center justify-center rounded-lg hover:bg-red-500 hover:text-white cursor-pointer transition-colors group text-slate-500">
                   <X size={18} className="group-hover:text-white transition-colors" />
               </div>
           </div>
        </div>
    );
};

const commonFonts = [
  { label: 'Calibri', value: 'Calibri, sans-serif', fontFamily: 'Calibri, sans-serif' },
  { label: 'Arial', value: 'Arial, sans-serif', fontFamily: 'Arial, sans-serif' },
  { label: 'Inter', value: 'Inter, sans-serif', fontFamily: 'Inter, sans-serif' },
  { label: 'Times New Roman', value: '"Times New Roman", serif', fontFamily: '"Times New Roman", serif' },
  { label: 'Courier New', value: '"Courier New", monospace', fontFamily: '"Courier New", monospace' },
];

const presetColors = [
  { label: 'Black', value: '#000000', color: '#000000' },
  { label: 'Red', value: '#EF4444', color: '#EF4444' },
  { label: 'Orange', value: '#F97316', color: '#F97316' },
  { label: 'Yellow', value: '#EAB308', color: '#EAB308' },
  { label: 'Green', value: '#10B981', color: '#10B981' },
  { label: 'Blue', value: '#3B82F6', color: '#3B82F6' },
  { label: 'Purple', value: '#8B5CF6', color: '#8B5CF6' },
  { label: 'Pink', value: '#EC4899', color: '#EC4899' },
  { label: 'White', value: '#ffffff', color: '#ffffff' },
  { label: 'None', value: 'transparent', color: 'transparent'}
];

const HomeRibbon = () => {
  const state = useSharedState();
  const activeCellData = state.data[state.activeCell] || {};
  return (
    <div className="bg-white/80 backdrop-blur-md rounded-[24px] p-4 shadow-[0_4px_20px_-4px_rgba(0,0,0,0.05)] border border-white flex items-stretch gap-2 overflow-x-auto [&::-webkit-scrollbar]:hidden [-ms-overflow-style:none] [scrollbar-width:none] min-w-max h-[116px] shrink-0">
      {/* Font Section */}
      <RibbonSection title="Font">
         <div className="flex gap-2">
             <div className="flex gap-0.5">
                 <ToolbarItem icon={Bold} label="Bold" active={activeCellData.bold} onClick={() => updateSelectionStyle('bold', null, true)} />
                 <ToolbarItem icon={Italic} label="Italic" active={activeCellData.italic} onClick={() => updateSelectionStyle('italic', null, true)} />
                 <ToolbarItem icon={Underline} label="Underline" active={activeCellData.underline} onClick={() => updateSelectionStyle('underline', null, true)} />
             </div>
             <div className="flex flex-col justify-between py-0.5 w-[160px]">
                 <div className="flex gap-1 items-center px-1">
                   <EditableDropdown 
                      items={commonFonts}
                      value={activeCellData.fontFamily || 'Calibri, sans-serif'}
                      onSelect={(val: string) => updateSelectionStyle('fontFamily', val)}
                      placeholder="Font Family"
                      className="flex-1 min-w-0"
                   />
                   <EditableDropdown
                      items={[8,9,10,11,12,14,16,18,20,22,24,28,32,36,48,72].map(s => ({label: s.toString(), value: s}))}
                      value={activeCellData.fontSize || 13}
                      onSelect={(val: number) => updateSelectionStyle('fontSize', val)}
                      placeholder="13"
                      className="w-[45px] shrink-0"
                      isNumber={true}
                   />
                 </div>
                 <div className="flex items-center justify-between px-3 mt-1 text-slate-600">
                    <ColorPickerMenu 
                       color={activeCellData.fillColor}
                       onSelect={(val: string) => updateSelectionStyle('fillColor', val)}
                       trigger={<ColorPickerIcon icon={PaintBucket} isFill color={activeCellData.fillColor} />}
                    />
                    <ColorPickerMenu 
                       color={activeCellData.textColor}
                       onSelect={(val: string) => updateSelectionStyle('textColor', val)}
                       trigger={<ColorPickerIcon icon={Baseline} isText color={activeCellData.textColor} />}
                    />
                 </div>
             </div>
         </div>
      </RibbonSection>

      {/* Alignment Section */}
      <RibbonSection title="Alignment">
         <div className="flex gap-0.5">
             <ToolbarItem icon={AlignLeft} label="Left" active={activeCellData.align === 'left' || !activeCellData.align} onClick={() => updateSelectionStyle('align', 'left')} />
             <ToolbarItem icon={AlignCenter} label={"Center"} active={activeCellData.align === 'center'} onClick={() => updateSelectionStyle('align', 'center')} />
             <ToolbarItem icon={AlignRight} label={"Right"} active={activeCellData.align === 'right'} onClick={() => updateSelectionStyle('align', 'right')} />
             <ToolbarItem icon={ArrowUpDown} label={"Vertical\nAlign"} active={activeCellData.vAlign === 'middle'} onClick={() => updateSelectionStyle('vAlign', activeCellData.vAlign === 'middle' ? 'bottom' : 'middle')} />
             <ToolbarItem icon={WrapText} label={"Wrap\nText"} active={activeCellData.wrapText} onClick={() => updateSelectionStyle('wrapText', null, true)} />
         </div>
      </RibbonSection>

      {/* Number Section */}
      <RibbonSection title="Number">
         <div className="flex flex-col justify-between py-0.5 w-[140px] h-[52px]">
             <button className="flex items-center justify-between border border-slate-200 rounded-full px-4 py-1.5 text-[11px] font-medium text-slate-700 hover:bg-slate-50 transition-colors w-full cursor-pointer outline-none">
                Number Format
                <ChevronDown size={14} className="text-slate-400" />
             </button>
             <div className="flex items-center justify-between px-3 text-slate-600 mt-2">
                <DollarSign size={16} className="cursor-pointer hover:text-slate-900" strokeWidth={1.5} />
                <Percent size={16} className="cursor-pointer hover:text-slate-900" strokeWidth={1.5} />
                <span className="font-serif italic text-[16px] leading-none cursor-pointer hover:text-slate-900 -mt-1">f</span>
                <span className="flex gap-[1px] justify-center cursor-pointer">
                    <span className="w-1.5 h-1.5 rounded-full border-2 border-slate-500"></span>
                    <span className="w-1.5 h-1.5 rounded-full border-2 border-slate-500"></span>
                </span>
             </div>
         </div>
      </RibbonSection>

      {/* Styles Section */}
      <RibbonSection title="Styles">
         <div className="flex gap-1">
             <ToolbarItem icon={Lightbulb} label={"Conditional\nFormatting"} iconColor="text-rose-500" />
             <ToolbarItem icon={TableIcon} label={"Table\nStyles"} iconColor="text-emerald-600" />
         </div>
      </RibbonSection>

      {/* Cells Section */}
      <RibbonSection title="Cells" noBorder>
         <div className="flex gap-1">
             <ToolbarItem icon={ArrowRightToLine} label="Insert" />
             <ToolbarItem icon={Trash2} label="Delete" />
             <ToolbarItem icon={Wand2} label="Format" />
         </div>
      </RibbonSection>
    </div>
  );
}

const PageLayoutRibbon = () => (
    <div className="bg-white/80 backdrop-blur-md rounded-[24px] p-4 shadow-[0_4px_20px_-4px_rgba(0,0,0,0.05)] border border-white flex items-stretch gap-2 overflow-x-auto [&::-webkit-scrollbar]:hidden [-ms-overflow-style:none] [scrollbar-width:none] min-w-max h-[116px] shrink-0">
      <RibbonSection title="Themes">
         <ToolbarItem icon={Palette} label="Themes" iconColor="text-purple-600" active />
         <div className="flex flex-col justify-center gap-0.5 ml-2 h-full">
            <SmallToolbarItem icon={Palette} label="Colors" iconColor="text-emerald-500" />
            <SmallToolbarItem icon={Type} label="Fonts" iconColor="text-blue-500" />
            <SmallToolbarItem icon={Wand2} label="Effects" iconColor="text-slate-500" />
         </div>
      </RibbonSection>

      <RibbonSection title="Page Setup">
         <ToolbarItem icon={LayoutTemplate} label="Margins" iconColor="text-emerald-500" />
         <ToolbarItem icon={ImageIcon} label="Orientation" iconColor="text-emerald-400" />
         <ToolbarItem icon={LayoutList} label="Size" iconColor="text-blue-500" />
         <ToolbarItem icon={Printer} label={"Print Area"} iconColor="text-slate-500" />
      </RibbonSection>

      <RibbonSection title="Scale to Fit">
         <div className="flex flex-col gap-[3px] justify-center h-full">
             <DropdownField label="Width" value="Automatic" />
             <DropdownField label="Height" value="Automatic" />
             <DropdownField label="Scale" value="100%" />
         </div>
      </RibbonSection>

      <RibbonSection title="Sheet Options">
         <div className="flex gap-6 h-full pt-1 px-2">
             <CheckboxGroup title="Gridlines" items={[{label: 'View', checked: true}, {label: 'Print', checked: false}]} />
             <CheckboxGroup title="Headings" items={[{label: 'View', checked: true}, {label: 'Print', checked: false}]} />
         </div>
      </RibbonSection>

      <RibbonSection title="Arrange" noBorder>
         <div className="flex gap-1 h-full">
            <ToolbarItem icon={BringToFront} label={"Bring\nForward"} />
            <ToolbarItem icon={SendToBack} label={"Send\nBackward"} />
            <div className="flex flex-col justify-center gap-0.5 ml-1 h-full">
               <SmallToolbarItem icon={RotateCw} label="Rotate" />
               <SmallToolbarItem icon={MoveUp} label="Reorder Up" />
            </div>
            <div className="flex flex-col justify-center gap-0.5 h-full">
                <SmallToolbarItem icon={Layers} label="Selection Pane" />
                <SmallToolbarItem icon={MoveDown} label="Reorder Down" />
            </div>
            <div className="flex flex-col justify-center gap-0.5 ml-1 h-full">
               <SmallToolbarItem icon={Focus} label="Align" />
               <SmallToolbarItem icon={Grip} label="Snap to Grid" />
            </div>
         </div>
      </RibbonSection>
    </div>
);

const InsertRibbon = () => (
    <div className="flex flex-col gap-2">
        <div className="bg-white/80 backdrop-blur-md rounded-[24px] p-4 shadow-[0_4px_20px_-4px_rgba(0,0,0,0.05)] border border-white flex items-stretch gap-2 overflow-x-auto [&::-webkit-scrollbar]:hidden [-ms-overflow-style:none] [scrollbar-width:none] min-w-max h-[116px] shrink-0">
          <RibbonSection title="Tables">
             <ToolbarItem icon={LayoutDashboard} label={"PivotTable"} iconColor="text-emerald-700" />
             <ToolbarItem icon={CheckSquare} label={"Recommended\nPivotTables"} iconColor="text-emerald-600" />
             <ToolbarItem icon={TableIcon} label={"Table"} iconColor="text-emerald-600" />
          </RibbonSection>

          <RibbonSection title="Illustrations">
             <ToolbarItem icon={ImageIcon} label={"Pictures"} iconColor="text-sky-600" />
             <ToolbarItem icon={Shapes} label={"Shapes"} iconColor="text-blue-500" />
             <ToolbarItem icon={Sticker} label={"Icons"} iconColor="text-sky-500" />
             <ToolbarItem icon={Box} label={"3D\nModels"} iconColor="text-sky-600" />
             <ToolbarItem icon={Network} label={"SmartArt"} iconColor="text-emerald-600" />
          </RibbonSection>

          <RibbonSection title="Charts">
             <ToolbarItem icon={BarChart2} label={"Recommended\nCharts"} iconColor="text-blue-500" />
             <ToolbarItem icon={BarChart} label="Column" iconColor="text-amber-500" />
             <ToolbarItem icon={LineChart} label="Line" iconColor="text-orange-500" />
             <ToolbarItem icon={PieChart} label="Pie" iconColor="text-orange-400" />
             <ToolbarItem icon={AlignLeft} label="Bar" iconColor="text-amber-500" />
             <ToolbarItem icon={AreaChart} label="Area" iconColor="text-sky-500" />
             <ToolbarItem icon={Activity} label="Scatter" iconColor="text-sky-500" />
          </RibbonSection>

          <RibbonSection title="Sparklines" noBorder>
             <ToolbarItem icon={TrendingUp} label={"Line"} iconColor="text-blue-600" />
             <ToolbarItem icon={BarChart2} label={"Column"} iconColor="text-blue-600" />
             <ToolbarItem icon={ArrowUpDown} label={"Win/Loss"} iconColor="text-blue-400" />
          </RibbonSection>
        </div>

        <div className="bg-white/80 backdrop-blur-md rounded-[24px] p-4 shadow-[0_4px_20px_-4px_rgba(0,0,0,0.05)] border border-white flex items-stretch gap-2 overflow-x-auto [&::-webkit-scrollbar]:hidden [-ms-overflow-style:none] [scrollbar-width:none] min-w-max h-[116px] shrink-0">
          <RibbonSection title="Charts">
             <ToolbarItem icon={Map} label="Map" iconColor="text-teal-600" />
             <ToolbarItem icon={PanelTop} label={"Header\n& Footer"} iconColor="text-amber-500" />
             <ToolbarItem icon={Activity} label={"Win/Loss"} iconColor="text-teal-500" />
             <ToolbarItem icon={TrendingUp} label="Stock" iconColor="text-emerald-600" />
             <ToolbarItem icon={Layers} label="Surface" iconColor="text-purple-500" />
             <ToolbarItem icon={Radar} label="Radar" iconColor="text-indigo-400" />
             <ToolbarItem icon={LayoutGrid} label="Treemap" iconColor="text-indigo-500" />
             <ToolbarItem icon={Sun} label="Sunburst" iconColor="text-amber-500" />
             <ToolbarItem icon={BarChart} label="Histogram" iconColor="text-blue-600" />
             <ToolbarItem icon={SlidersHorizontal} label={"Box &\nWhisker"} iconColor="text-orange-500" />
             <ToolbarItem icon={TrendingDown} label="Waterfall" iconColor="text-teal-500" />
             <ToolbarItem icon={Filter} label="Funnel" iconColor="text-blue-400" />
             <ToolbarItem icon={LineChart} label="Combo" iconColor="text-indigo-500" />
          </RibbonSection>

          <RibbonSection title="Text" noBorder>
             <ToolbarItem icon={Type} label={"Text\nBox"} iconColor="text-slate-600" />
             <ToolbarItem icon={Baseline} label={"WordArt"} iconColor="text-blue-500" />
             <ToolbarItem icon={FileSignature} label={"Signature\nLine"} iconColor="text-indigo-500" />
             <ToolbarItem icon={MousePointer2} label="Object" iconColor="text-teal-600" />
             <ToolbarItem icon={Sigma} label="Equation" iconColor="text-indigo-600" />
             <ToolbarItem icon={Globe} label="Symbol" iconColor="text-purple-600" />
          </RibbonSection>
        </div>
    </div>
);

const FormulasRibbon = () => (
    <div className="bg-white/80 backdrop-blur-md rounded-[24px] p-4 shadow-[0_4px_20px_-4px_rgba(0,0,0,0.05)] border border-white flex items-stretch gap-2 overflow-x-auto [&::-webkit-scrollbar]:hidden [-ms-overflow-style:none] [scrollbar-width:none] min-w-max h-[116px] shrink-0">
      <RibbonSection title="Function Library">
         <ToolbarItem icon={Hexagon} label={"Insert\nFunction"} active />
         <ToolbarItem icon={Sigma} label={"AutoSum"} iconColor="text-blue-500" />
         <ToolbarItem icon={Clock} label={"Recently\nUsed"} iconColor="text-blue-500" />
         <ToolbarItem icon={DollarSign} label={"Financial"} iconColor="text-amber-600" />
         <ToolbarItem icon={Binary} label={"Logical"} iconColor="text-indigo-500" />
         <ToolbarItem icon={Type} label={"Text"} iconColor="text-emerald-500" />
         <ToolbarItem icon={Calendar} label={"Date & Time"} iconColor="text-orange-500" />
         <ToolbarItem icon={Search} label={"Lookup"} iconColor="text-purple-500" />
         <ToolbarItem icon={Calculator} label={"Math & Trig"} iconColor="text-teal-500" />
         <ToolbarItem icon={MoreHorizontal} label={"More"} iconColor="text-slate-500" />
      </RibbonSection>

      <RibbonSection title="Defined Names">
         <ToolbarItem icon={Tags} label={"Name\nManager"} iconColor="text-blue-500" />
         <div className="flex flex-col justify-center gap-0.5 ml-1 h-full w-[150px]"> 
            <SmallToolbarItem icon={Tag} label="Define Name:" />
            <SmallToolbarItem icon={BookmarkPlus} label="Use in Formula:" />
            <SmallToolbarItem icon={CheckSquare} label="Create from Selection" />
         </div>
      </RibbonSection>

      <RibbonSection title="Formula Auditing">
         <div className="flex flex-col justify-center gap-0.5 h-full"> 
            <SmallToolbarItem icon={CornerUpLeft} label="Trace Precedents" iconColor="text-red-500" />
            <SmallToolbarItem icon={CornerUpRight} label="Trace Dependents" iconColor="text-blue-500" />
            <SmallToolbarItem icon={Eraser} label="Remove Arrows" iconColor="text-red-600" />
         </div>
         <div className="flex flex-col max-w-[130px] justify-center gap-0.5 ml-2 h-full"> 
            <SmallToolbarItem icon={Eye} label="Show Formulas" active />
            <SmallToolbarItem icon={CheckCircle2} label="Error Checking" iconColor="text-orange-500" />
            <SmallToolbarItem icon={Search} label="Evaluate Formula" iconColor="text-sky-500" />
         </div>
         <div className="ml-2 pt-1 border-l border-slate-200/50 pl-2">
             <ToolbarItem icon={LayoutDashboard} label={"Watch\nWindow"} iconColor="text-emerald-600" />
         </div>
      </RibbonSection>

      <RibbonSection title="Calculation" noBorder>
         <ToolbarItem icon={Calculator} label={"Calculation\nOptions"} iconColor="text-blue-500" />
         <div className="flex flex-col justify-center gap-0.5 ml-2 h-full"> 
            <SmallToolbarItem icon={CheckSquare} label="Calculate Now" iconColor="text-emerald-600" active />
            <SmallToolbarItem icon={PlaySquare} label="Calculate Sheet" iconColor="text-blue-500" />
         </div>
      </RibbonSection>
    </div>
);

const FormulaBar = () => {
    const state = useSharedState();
    const cellValue = state.data[state.activeCell]?.value || '';

    return (
        <div className="bg-white/80 backdrop-blur-md rounded-full p-1.5 pl-6 shadow-[0_2px_10px_-4px_rgba(0,0,0,0.05)] border border-white flex items-center gap-3 w-full shrink-0">
            <div className="flex items-center gap-4">
                <span className="text-sm font-semibold text-slate-800 w-8 text-center">{state.activeCell}</span>
                <div className="w-px h-6 bg-slate-200" />
            </div>
            <div className="flex items-center gap-4 flex-1 pl-1">
                <span className="font-serif italic font-medium tracking-tighter text-slate-500 text-lg leading-none">fx</span>
                <input 
                   type="text" 
                   value={cellValue}
                   onChange={(e) => {
                       const newData = { ...sharedState.data, [state.activeCell]: { ...sharedState.data[state.activeCell], value: e.target.value } };
                       dispatch({ data: newData });
                   }}
                   placeholder="Type a value or formula..."
                   className="flex-1 bg-white border border-slate-100 rounded-full px-5 py-1.5 text-sm outline-none text-slate-700 placeholder-slate-400 font-medium shadow-inner shadow-slate-50/50" 
                />
            </div>
        </div>
    );
};

const Cell = React.memo(({ col, row, colIdx, isActive, isSelected, isEditing, cellData }: any) => {
  const cellId = `${col}${row}`;

  const handleMouseDown = (e: any) => {
    if (e.detail === 2) {
      dispatch({ editingCell: cellId });
      return;
    }
    dispatch({
      activeCell: cellId,
      selection: { start: cellId, end: cellId },
      isDragging: true,
      editingCell: null,
      showFloatingMenu: false
    });
  };

  const handleMouseEnter = () => {
    if (sharedState.isDragging && sharedState.selection.end !== cellId) {
      dispatch({ selection: { ...sharedState.selection, end: cellId }, showFloatingMenu: false });
    }
  };

  return (
          <td 
      className={`w-[100px] min-w-[100px] h-8 border-r border-b border-slate-300 relative outline-none cursor-cell ${isSelected && !isActive ? 'bg-[#e4edfe]' : 'bg-white'} ${isActive ? 'z-20' : 'z-10'}`}
      onMouseDown={handleMouseDown}
      onMouseEnter={handleMouseEnter}
      onDoubleClick={() => { dispatch({ editingCell: cellId, showFloatingMenu: false }); }}
      style={{
         backgroundColor: cellData?.fillColor || (isSelected && !isActive ? '#e4edfe' : 'white')
      }}
    >
      {isSelected && !isActive && (
        <div className="absolute inset-0 bg-blue-500/10 pointer-events-none mix-blend-multiply" />
      )}
      {isActive && !isEditing && (
        <div className="absolute inset-x-[-1px] inset-y-[-1px] border-2 border-blue-500 z-30 pointer-events-none shadow-sm">
           <div className="absolute -bottom-1 -right-1 w-[7px] h-[7px] bg-blue-500 border border-white cursor-crosshair pointer-events-auto" />
        </div>
      )}
      <div className={`w-full h-full px-1.5 flex items-center ${cellData?.wrapText ? 'whitespace-normal break-words' : 'overflow-hidden whitespace-nowrap'} text-[13px] text-slate-800 select-none`}
           style={{
             fontWeight: cellData?.bold ? 600 : 400,
             fontStyle: cellData?.italic ? 'italic' : 'normal',
             textDecoration: cellData?.underline ? 'underline' : 'none',
             justifyContent: cellData?.align === 'center' ? 'center' : (cellData?.align === 'right' ? 'flex-end' : 'flex-start'),
             alignItems: cellData?.vAlign === 'middle' ? 'center' : (cellData?.vAlign === 'bottom' ? 'flex-end' : 'flex-start'),
             fontFamily: cellData?.fontFamily || 'inherit',
             fontSize: cellData?.fontSize ? `${cellData.fontSize}px` : 'inherit',
             color: cellData?.textColor || 'inherit',
           }}>
        {isEditing ? (
          <input 
            autoFocus
            className={`absolute inset-[-1px] px-1.5 outline-none focus:ring-0 focus:outline-none border-2 border-blue-500 bg-white z-40 shadow-sm`}
            value={cellData?.value || ''}
            onChange={(e) => {
              const newData = { ...sharedState.data, [cellId]: { ...cellData, value: e.target.value } };
              dispatch({ data: newData });
            }}
            onBlur={() => dispatch({ editingCell: null })}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                e.preventDefault();
                const nextCell = `${col}${row+1}`;
                dispatch({ editingCell: null, activeCell: nextCell, selection: { start: nextCell, end: nextCell } });
              }
            }}
            style={{
               fontWeight: cellData?.bold ? 600 : 400,
               fontStyle: cellData?.italic ? 'italic' : 'normal',
               textDecoration: cellData?.underline ? 'underline' : 'none',
               textAlign: cellData?.align || 'left',
               fontFamily: cellData?.fontFamily || 'inherit',
               fontSize: cellData?.fontSize ? `${cellData.fontSize}px` : 'inherit',
               color: cellData?.textColor || 'inherit'
            }}
          />
        ) : (
          cellData?.value || ''
        )}
      </div>
    </td>
  );
}, (prev, next) => {
   return prev.isActive === next.isActive && 
          prev.isSelected === next.isSelected && 
          prev.isEditing === next.isEditing && 
          prev.cellData === next.cellData;
});

const SpreadsheetGrid = () => {
  const state = useSharedState();
  const range = state.selection ? getRange(state.selection.start, state.selection.end) : null;

  useEffect(() => {
    const handleMouseUp = (e: MouseEvent) => {
      if (sharedState.isDragging) {
        dispatch({ isDragging: false });
        
        const currentRange = getRange(sharedState.selection.start, sharedState.selection.end);
        if (currentRange) {
           dispatch({ 
              showFloatingMenu: true, 
              floatingPos: { top: Math.max(10, e.clientY + 15), left: Math.max(10, e.clientX + 15) } 
           });
        }
      }
    };
    
    const handleKeyDown = (e: KeyboardEvent) => {
      if (!sharedState.editingCell && e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
        const cellId = sharedState.activeCell;
        const cellData = sharedState.data[cellId] || { value: '' };
        dispatch({ 
          editingCell: cellId,
          data: { ...sharedState.data, [cellId]: { ...cellData, value: e.key } },
          showFloatingMenu: false
        });
      }
      if ((e.key === 'Delete' || e.key === 'Backspace') && !sharedState.editingCell) {
         const { start, end } = sharedState.selection;
         const { minColIdx, maxColIdx, minRow, maxRow } = getRange(start, end);
         const newData = { ...sharedState.data };
         for(let r=minRow; r<=maxRow; r++) {
            for(let c=minColIdx; c<=maxColIdx; c++) {
               const id = `${cols[c]}${r}`;
               if (newData[id]) newData[id] = { ...newData[id], value: '' };
            }
         }
         dispatch({ data: newData, showFloatingMenu: false });
      }
    };

    window.addEventListener('mouseup', handleMouseUp);
    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('mouseup', handleMouseUp);
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, []);

  return (
    <div className="bg-white/90 backdrop-blur-3xl rounded-[20px] shadow-[0_4px_25px_-5px_rgba(0,0,0,0.06)] border border-white w-full flex-1 overflow-hidden flex flex-col min-h-0 relative">
      <div className="overflow-auto [&::-webkit-scrollbar]:hidden [-ms-overflow-style:none] [scrollbar-width:none] flex-1 relative bg-white m-1 rounded-[16px] border border-slate-200">
        <table className="w-max border-collapse table-fixed text-[13px] bg-white">
          <thead className="sticky top-0 z-40 shadow-sm border-b-2 border-slate-300">
            <tr>
              <th className="w-12 min-w-[48px] h-8 border-r border-b border-slate-300 bg-[#f4f7fb] sticky left-0 z-50"></th>
              {cols.map((col, colIdx) => {
                const isActive = range && colIdx >= range.minColIdx && colIdx <= range.maxColIdx;
                return (
                  <th 
                    key={col} 
                    className={`w-[100px] min-w-[100px] border-r border-b border-slate-300 font-medium py-1 text-slate-600 ${isActive ? 'bg-[#e4edfe] text-blue-900 border-b-blue-400' : 'bg-[#f4f7fb]'}`}
                  >
                    {col}
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => {
              const isActiveRow = range && row >= range.minRow && row <= range.maxRow;
              return (
                <tr key={row} className="h-8 group/tr">
                  <td className={`w-12 min-w-[48px] border-r border-b border-slate-300 text-center font-medium sticky left-0 z-30 select-none ${isActiveRow ? 'bg-[#e4edfe] text-blue-900 border-r-blue-400' : 'text-slate-500 bg-[#f4f7fb]'}`}>
                    {row}
                  </td>
                  {cols.map((col, colIdx) => {
                    const cellId = `${col}${row}`;
                    const isSelected = range ? (colIdx >= range.minColIdx && colIdx <= range.maxColIdx && row >= range.minRow && row <= range.maxRow) : false;
                    return (
                       <Cell 
                          key={col}
                          col={col}
                          row={row}
                          colIdx={colIdx}
                          isActive={state.activeCell === cellId}
                          isSelected={isSelected}
                          isEditing={state.editingCell === cellId}
                          cellData={state.data[cellId]}
                       />
                    )
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      
      {state.showFloatingMenu && (
        <div 
          className="fixed flex items-center gap-1.5 p-1 bg-white/95 backdrop-blur-xl border border-slate-200/80 shadow-[0_8px_30px_-4px_rgba(0,0,0,0.15)] rounded-lg z-50 animate-in fade-in zoom-in-95 duration-100"
          style={{ top: state.floatingPos.top, left: state.floatingPos.left }}
          onMouseDown={(e) => {
             e.stopPropagation();
             e.preventDefault();
          }}
        >
          <div className="flex bg-slate-100/60 rounded-md p-0.5">
             <ToolbarItem small icon={Bold} active={state.data[state.activeCell]?.bold} iconColor="text-slate-700" onClick={() => updateSelectionStyle('bold', null, true)} />
             <ToolbarItem small icon={Italic} active={state.data[state.activeCell]?.italic} iconColor="text-slate-700" onClick={() => updateSelectionStyle('italic', null, true)} />
             <ToolbarItem small icon={Underline} active={state.data[state.activeCell]?.underline} iconColor="text-slate-700" onClick={() => updateSelectionStyle('underline', null, true)} />
          </div>
          <div className="w-[1px] h-6 bg-slate-200 mx-0.5" />
          <div className="flex bg-slate-100/60 rounded-md p-0.5 relative z-50">
             <ColorPickerMenu 
                color={state.data[state.activeCell]?.fillColor}
                onSelect={(val: string) => updateSelectionStyle('fillColor', val)}
                trigger={<ColorPickerIcon icon={PaintBucket} isFill color={state.data[state.activeCell]?.fillColor} />}
             />
             <ColorPickerMenu 
                color={state.data[state.activeCell]?.textColor}
                onSelect={(val: string) => updateSelectionStyle('textColor', val)}
                trigger={<ColorPickerIcon icon={Baseline} isText color={state.data[state.activeCell]?.textColor} />}
             />
          </div>
        </div>
      )}
    </div>
  );
};

export default function App() {
  const [activeTab, setActiveTab] = useState('Home');

  return (
    <div className="h-screen w-screen flex flex-col items-center py-6 px-4 md:px-8 gap-4 overflow-hidden font-sans">
        <div className="w-full max-w-[1400px] flex flex-col gap-4 h-full mx-auto">
            <MenuTabs activeTab={activeTab} setActiveTab={setActiveTab} />
            {activeTab === 'Home' && <HomeRibbon />}
            {activeTab === 'Page Layout' && <PageLayoutRibbon />}
            {activeTab === 'Insert' && <InsertRibbon />}
            {activeTab === 'Formulas' && <FormulasRibbon />}
            {!['Home', 'Page Layout', 'Insert', 'Formulas'].includes(activeTab) && <HomeRibbon />}
            <FormulaBar />
            <SpreadsheetGrid />
        </div>
    </div>
  );
}
