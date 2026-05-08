import React, { useState } from 'react';
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

const ToolbarItem = ({ icon: Icon, label, active = false, iconColor = "text-slate-600" }: any) => (
  <div className="flex flex-col items-center justify-start gap-1 cursor-pointer min-w-[52px] group px-1">
    <div className={`p-1.5 rounded-lg transition-colors ${active ? 'bg-[#e5ecf6] text-blue-700' : 'hover:bg-slate-100 ' + iconColor}`}>
      <Icon size={18} strokeWidth={active ? 2 : 1.5} />
    </div>
    <span className="text-[10px] text-slate-500 font-medium whitespace-pre-line text-center leading-[1.1] opacity-90 group-hover:opacity-100 transition-opacity">
      {label}
    </span>
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
           <div className="flex items-center flex-1 overflow-x-auto custom-scrollbar">
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

const HomeRibbon = () => {
  return (
    <div className="bg-white/80 backdrop-blur-md rounded-[24px] p-4 shadow-[0_4px_20px_-4px_rgba(0,0,0,0.05)] border border-white flex items-stretch gap-2 overflow-x-auto custom-scrollbar min-w-max h-[116px] shrink-0">
      {/* Font Section */}
      <RibbonSection title="Font">
         <div className="flex gap-2">
             <div className="flex gap-0.5">
                 <ToolbarItem icon={Bold} label="Bold" active />
                 <ToolbarItem icon={Italic} label="Italic" />
                 <ToolbarItem icon={Underline} label="Underline" />
             </div>
             <div className="flex flex-col justify-between py-0.5 w-[130px]">
                 <button className="flex items-center justify-between border border-slate-200 rounded-full px-4 py-1.5 text-[11px] font-medium text-slate-700 hover:bg-slate-50 transition-colors w-full cursor-pointer outline-none">
                    Font Family
                    <ChevronDown size={14} className="text-slate-400" />
                 </button>
                 <div className="flex items-center justify-between px-3 mt-1 text-slate-600">
                    <div className="flex gap-0.5 items-center px-1 text-slate-600">
                        <span className="font-serif font-bold text-[14px] leading-none hover:text-slate-900 cursor-pointer">A</span>
                        <ChevronDown size={10} className="hover:text-slate-900 cursor-pointer stroke-[3]" />
                    </div>
                    <PaintBucket size={15} className="cursor-pointer hover:text-slate-900" />
                    <div className="relative cursor-pointer group">
                        <Baseline size={16} className="text-slate-600 group-hover:text-slate-900" />
                        <div className="absolute bottom-0 left-0 right-0 h-[2px] bg-red-600" />
                    </div>
                 </div>
             </div>
         </div>
      </RibbonSection>

      {/* Alignment Section */}
      <RibbonSection title="Alignment">
         <div className="flex gap-0.5">
             <ToolbarItem icon={AlignLeft} label="Left" active />
             <ToolbarItem icon={AlignCenter} label={"Center"} />
             <ToolbarItem icon={AlignRight} label={"Right"} />
             <ToolbarItem icon={ArrowUpDown} label={"Vertical\nAlign"} />
             <ToolbarItem icon={WrapText} label={"Wrap\nText"} />
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
    <div className="bg-white/80 backdrop-blur-md rounded-[24px] p-4 shadow-[0_4px_20px_-4px_rgba(0,0,0,0.05)] border border-white flex items-stretch gap-2 overflow-x-auto custom-scrollbar min-w-max h-[116px] shrink-0">
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
        <div className="bg-white/80 backdrop-blur-md rounded-[24px] p-4 shadow-[0_4px_20px_-4px_rgba(0,0,0,0.05)] border border-white flex items-stretch gap-2 overflow-x-auto custom-scrollbar min-w-max h-[116px] shrink-0">
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

        <div className="bg-white/80 backdrop-blur-md rounded-[24px] p-4 shadow-[0_4px_20px_-4px_rgba(0,0,0,0.05)] border border-white flex items-stretch gap-2 overflow-x-auto custom-scrollbar min-w-max h-[116px] shrink-0">
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
    <div className="bg-white/80 backdrop-blur-md rounded-[24px] p-4 shadow-[0_4px_20px_-4px_rgba(0,0,0,0.05)] border border-white flex items-stretch gap-2 overflow-x-auto custom-scrollbar min-w-max h-[116px] shrink-0">
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
    return (
        <div className="bg-white/80 backdrop-blur-md rounded-full p-1.5 pl-6 shadow-[0_2px_10px_-4px_rgba(0,0,0,0.05)] border border-white flex items-center gap-3 w-full shrink-0">
            <div className="flex items-center gap-4">
                <span className="text-sm font-semibold text-slate-800 w-6 text-center">D7</span>
                <div className="w-px h-6 bg-slate-200" />
            </div>
            <div className="flex items-center gap-4 flex-1 pl-1">
                <span className="font-serif italic font-medium tracking-tighter text-slate-500 text-lg leading-none">fx</span>
                <input 
                   type="text" 
                   defaultValue="Column Name" 
                   className="flex-1 bg-white border border-slate-100 rounded-full px-5 py-1.5 text-sm outline-none text-slate-700 placeholder-slate-400 font-medium shadow-inner shadow-slate-50/50" 
                />
            </div>
        </div>
    );
};

const SpreadsheetGrid = () => {
  const cols = Array.from({length: 40}, (_, i) => {
    if (i < 26) return String.fromCharCode(65 + i);
    return 'A' + String.fromCharCode(65 + (i - 26));
  });
  const rows = Array.from({length: 40}, (_, i) => i + 1);

  return (
    <div className="bg-white/90 backdrop-blur-3xl rounded-[20px] shadow-[0_4px_25px_-5px_rgba(0,0,0,0.06)] border border-white w-full flex-1 overflow-hidden flex flex-col min-h-0">
      <div className="overflow-auto custom-scrollbar flex-1 relative bg-white m-1 rounded-[16px] border border-slate-100">
        <table className="w-max border-collapse table-fixed text-[13px] bg-white">
          <thead className="sticky top-0 z-20">
            <tr>
              <th className="w-12 min-w-[48px] h-8 border-r border-b border-slate-200 bg-[#f9fafb] sticky left-0 z-30"></th>
              {cols.map((col) => (
                <th 
                  key={col} 
                  className="w-[100px] min-w-[100px] border-r border-b border-slate-200 font-medium py-1.5 text-slate-500 bg-[#f9fafb]"
                >
                  {col}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => (
              <tr key={row} className="h-8">
                <td className="w-12 min-w-[48px] border-r border-b border-slate-200 text-center font-medium sticky left-0 z-10 text-slate-400 bg-[#f9fafb]">
                  {row}
                </td>
                {cols.map((col) => (
                  <td 
                    key={col} 
                    className="w-[100px] min-w-[100px] border-r border-b border-slate-200 bg-white relative outline-none focus:bg-[#f0f5ff] focus:z-20 cursor-cell group" 
                    tabIndex={0}
                  >
                    <div className="absolute inset-0 border-2 border-transparent group-focus:border-blue-500 pointer-events-none z-20" />
                    <div className="absolute -bottom-1 -right-1 w-2 h-2 bg-blue-500 border border-white opacity-0 group-focus:opacity-100 z-30 pointer-events-none" />
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export default function App() {
  const [activeTab, setActiveTab] = useState('Page Layout');

  return (
    <div className="h-screen w-screen flex flex-col items-center py-6 px-4 md:px-8 gap-4 overflow-hidden font-sans">
        <div className="w-full max-w-[1400px] flex flex-col gap-4 h-full mx-auto">
            <MenuTabs activeTab={activeTab} setActiveTab={setActiveTab} />
            {activeTab === 'Home' && <HomeRibbon />}
            {activeTab === 'Page Layout' && <PageLayoutRibbon />}
            {activeTab === 'Insert' && <InsertRibbon />}
            {activeTab === 'Formulas' && <FormulasRibbon />}
            {/* Fallback for unused tabs */}
            {!['Home', 'Page Layout', 'Insert', 'Formulas'].includes(activeTab) && <HomeRibbon />}
            <FormulaBar />
            <SpreadsheetGrid />
        </div>
    </div>
  );
}

