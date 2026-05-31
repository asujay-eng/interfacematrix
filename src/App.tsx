import React, { useState, useEffect } from 'react';
import { supabase } from './supabaseClient';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import { Document, Packer, Paragraph, Table, TableCell, TableRow, TextRun, WidthType } from 'docx';

interface Subsystem {
  id: number;
  name: string;
  contract_package: string;
  code: string; 
}

interface SavedInterface {
  id: number;
  source_subsystem_id: number;
  target_subsystem_id: number;
  interface_level: string;
  color_accent: string;
  contract_package?: string;
}

export default function App() {
  const [subsystems, setSubsystems] = useState<Subsystem[]>([]);
  const [primarySystem, setPrimarySystem] = useState<string>('');
  const [targetSystems, setTargetSystems] = useState<string[]>([]);
  
  // Matrix Entry States
  const [sourceContract, setSourceContract] = useState<string>('P2427');
  const [targetContract, setTargetContract] = useState<string>('P2427');
  const [interfaceLevel, setInterfaceLevel] = useState<string>('Level 1');
  const [interfaceCategory, setInterfaceCategory] = useState<string>('Physical and Functional');
  const [savedInterfaces, setSavedInterfaces] = useState<SavedInterface[]>([]);
  const [loading, setLoading] = useState<boolean>(true);

  // DATABASE MANAGEMENT PANEL STATES
  const [editingSubsystemId, setEditingSubsystemId] = useState<number | null>(null);
  const [formName, setFormName] = useState<string>('');
  const [formContract, setFormContract] = useState<string>('P2427');

  const colorMap: Record<string, string> = {
    'Physical and Functional': '#3b82f6', // Blue
    'Only Physical': '#eab308',           // Yellow
    'Only Functional': '#10b981'          // Green
  };

  const loadData = async () => {
    try {
      const { data: subs, error: err1 } = await supabase.from('subsystems').select('*').order('id', { ascending: true });
      if (err1) throw err1;
      if (subs) setSubsystems(subs);

      const { data: matrix, error: err2 } = await supabase.from('interface_register').select('*');
      if (err2) throw err2;
      if (matrix) setSavedInterfaces(matrix);
    } catch (err) {
      console.error("Error pulling architecture matrix data tables:", err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadData();
  }, []);

  const refreshMatrix = async () => {
    const { data } = await supabase.from('interface_register').select('*');
    if (data) setSavedInterfaces(data);
  };

  const getContractColor = (pkg: string): string => {
    if (!pkg) return 'text-slate-400';
    const cleanPkg = pkg.trim().toUpperCase();
    if (cleanPkg.includes('2427')) return 'text-blue-400';
    if (cleanPkg.includes('2426')) return 'text-emerald-400';
    if (cleanPkg.includes('2428')) return 'text-purple-400';
    return 'text-slate-400';
  };

  // AUTOMATIC ACRONYM GENERATOR (Extracts bracketed targets or creates compact acronym uppercase tags)
  const generateAcronym = (name: string): string => {
    if (!name) return '';
    const bracketMatch = name.match(/\(([^)]+)\)/);
    if (bracketMatch && bracketMatch[1]) {
      return bracketMatch[1].toUpperCase();
    }
    // Fallback: If no parentheses exist, parse capital letters out of the words
    return name
      .split(/[\s-]+/)
      .filter(word => word.length > 0 && word[0] === word[0].toUpperCase() && !['AND', 'WITH', 'FOR', 'OR', 'TO', 'IN'].includes(word.toUpperCase()))
      .map(word => word[0])
      .join('')
      .replace(/[^A-Z]/g, '');
  };

  // Shows the full subsystem name in dropdowns, while keeping the acronym available as the code.
  const formatSubsystemOption = (sys: Subsystem): string => {
    const fullName = (sys.name || '').trim();
    const acronym = (sys.code || '').trim();

    if (!fullName) return acronym;
    if (!acronym) return fullName;

    // Avoid duplicate display when name is already "Building Management System (BMS)".
    if (fullName.toUpperCase().includes(`(${acronym.toUpperCase()})`)) {
      return fullName;
    }

    // Existing legacy rows may have name = "BMS" and code = "BMS".
    if (fullName.toUpperCase() === acronym.toUpperCase()) {
      return acronym;
    }

    return `${fullName} (${acronym})`;
  };

  // Keeps the matrix compact by displaying only the acronym/code in the grid.
  const getSubsystemCode = (sys: Subsystem): string => {
    return (sys.code || generateAcronym(sys.name) || sys.name || '').trim();
  };

  const handleCommitToMatrix = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!primarySystem || targetSystems.length === 0) {
      alert("Please select a Source System and at least one Target System.");
      return;
    }

    const activeColorHex = colorMap[interfaceCategory] || '#3b82f6';
    const insertRows = targetSystems.map(targetId => ({
      source_subsystem_id: parseInt(primarySystem),
      target_subsystem_id: parseInt(targetId),
      interface_level: interfaceLevel,
      interface_type: interfaceCategory,
      color_accent: activeColorHex,
      contract_package: sourceContract
    }));

    const { error } = await supabase.from('interface_register').upsert(insertRows);
    if (error) {
      alert("Database Matrix Update Error: " + error.message);
    } else {
      await refreshMatrix();
      setTargetSystems([]);
    }
  };

  // CRUD OPERATIONS: SUBSYSTEM ENTRY SUBMISSION (Extracts and commits clean acronym keys)
  const handleSaveSubsystem = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!formName.trim()) {
      alert("Please assign a subsystem name or acronym designation code.");
      return;
    }

    // Process acronym extraction from text input or accept the raw text if already formatted
    const cleanAcronym = generateAcronym(formName) || formName.trim().toUpperCase();

    const payload = {
      // Store the full subsystem title for dropdown display, e.g. "Building Management System (BMS)".
      name: formName.trim(),
      contract_package: formContract,
      // Store the acronym separately for compact matrix display and database keys.
      code: cleanAcronym
    };

    if (editingSubsystemId) {
      const { error } = await supabase.from('subsystems').update(payload).eq('id', editingSubsystemId);
      if (error) {
        alert("Update Failed: " + error.message);
      } else {
        setEditingSubsystemId(null);
        setFormName('');
        await loadData();
      }
    } else {
      const { error } = await supabase.from('subsystems').insert([payload]);
      if (error) {
        alert("Creation Failed: " + error.message);
      } else {
        setFormName('');
        await loadData();
      }
    }
  };

  const startEditSubsystem = (sys: Subsystem) => {
    setEditingSubsystemId(sys.id);
    setFormName(sys.name); // Loads the concise acronym value back into input fields
    setFormContract(sys.contract_package);
  };

  const handleDeleteSubsystem = async (id: number) => {
    if (!confirm("Are you sure you want to remove this subsystem completely? This could invalidate existing matrix relations mapping coordinates.")) return;
    
    const { error } = await supabase.from('subsystems').delete().eq('id', id);
    if (error) {
      alert("Deletion error (Ensure target isn't tied to active interfaces): " + error.message);
    } else {
      if (editingSubsystemId === id) {
        setEditingSubsystemId(null);
        setFormName('');
      }
      await loadData();
    }
  };

  const getCellInterface = (sourceId: number, targetId: number) => {
    return savedInterfaces.find(
      item => item.source_subsystem_id === sourceId && item.target_subsystem_id === targetId
    );
  };



  const buildMatrixExportData = () => {
    const header = ['Source / Target', ...subsystems.map(sys => formatSubsystemOption(sys))];

    const rows = subsystems.map(rowSys => {
      const row: Record<string, string> = {
        'Source / Target': formatSubsystemOption(rowSys)
      };

      subsystems.forEach(colSys => {
        if (rowSys.id === colSys.id) {
          row[formatSubsystemOption(colSys)] = 'N/A';
          return;
        }

        const cellInterface = getCellInterface(rowSys.id, colSys.id);
        row[formatSubsystemOption(colSys)] = cellInterface
          ? `${cellInterface.interface_level}`
          : '';
      });

      return row;
    });

    return { header, rows };
  };

  const handleExportExcel = () => {
    if (subsystems.length === 0) {
      alert('No subsystem data available to export.');
      return;
    }

    const { header, rows } = buildMatrixExportData();
    const worksheet = XLSX.utils.json_to_sheet(rows, { header });
    worksheet['!cols'] = header.map((heading, index) => ({
      wch: index === 0 ? 38 : Math.max(18, heading.length + 2)
    }));

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Interface Matrix');
    XLSX.writeFile(workbook, 'Interface_Grid_Matrix.xlsx');
  };

  const handleExportWord = async () => {
    if (subsystems.length === 0) {
      alert('No subsystem data available to export.');
      return;
    }

    const tableRows = [
      new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Source / Target', bold: true })] })] }),
          ...subsystems.map(sys => new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: getSubsystemCode(sys), bold: true })] })]
          }))
        ]
      }),
      ...subsystems.map(rowSys => new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: getSubsystemCode(rowSys), bold: true })] })] }),
          ...subsystems.map(colSys => {
            const cellInterface = getCellInterface(rowSys.id, colSys.id);
            const value = rowSys.id === colSys.id ? 'N/A' : (cellInterface ? cellInterface.interface_level : '');
            return new TableCell({ children: [new Paragraph(value)] });
          })
        ]
      }))
    ];

    const doc = new Document({
      sections: [{
        children: [
          new Paragraph({
            children: [new TextRun({ text: 'Interface Grid Matrix', bold: true, size: 32 })],
            spacing: { after: 240 }
          }),
          new Paragraph({
            text: 'Generated from the live interface register automation system.',
            spacing: { after: 240 }
          }),
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: tableRows
          })
        ]
      }]
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, 'Interface_Grid_Matrix.docx');
  };

  if (loading) {
    return <div className="p-8 text-center text-slate-400 bg-slate-900 min-h-screen">Compiling Live Engineering Grid Infrastructure...</div>;
  }

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 p-6 flex flex-col gap-6">
      {/* Top Header Row */}
      <header className="flex justify-between items-center border-b border-slate-800 pb-4">
        <div>
          <h1 className="text-xl font-bold tracking-tight text-white">Etihad HSR | Interactive Matrix Panel</h1>
          <p className="text-xs text-emerald-400 flex items-center gap-1.5 mt-0.5">
            <span className="h-1.5 w-1.5 rounded-full bg-emerald-400 animate-pulse"></span>
            Full Name Dropdown Configuration Online
          </p>
        </div>
        <div className="bg-slate-900 border border-slate-800 rounded-lg px-3 py-1.5 text-xs font-mono text-slate-400">
          Total Mapped Coordinates: <span className="text-blue-400 font-bold">{savedInterfaces.length}</span>
        </div>
      </header>

      {/* Main Grid Splitter Layout */}
      <div className="flex flex-col xl:flex-row gap-6 items-start">
        
        {/* Left Control Dashboard Form */}
        <div className="w-full xl:w-80 shrink-0 bg-slate-900 border border-slate-800 p-4 rounded-xl shadow-xl">
          <h2 className="text-sm font-semibold text-white mb-3 flex items-center gap-2">
            <span className="w-2 h-2 rounded-full bg-blue-500"></span> Matrix Entry Controller
          </h2>
          <form onSubmit={handleCommitToMatrix} className="space-y-4">
            
            {/* ROW SYSTEMS CONTROL */}
            <div className="border-b border-slate-800/80 pb-3 space-y-2.5">
              <div>
                <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">Row Contract Package</label>
                <select 
                  value={sourceContract} 
                  onChange={(e) => setSourceContract(e.target.value)}
                  className={`w-full bg-slate-950 border border-slate-800 rounded-lg px-2.5 py-1.5 text-xs font-bold focus:outline-none ${getContractColor(sourceContract)}`}
                >
                  <option value="P2427" className="text-blue-400 bg-slate-950">P2427</option>
                  <option value="P2428" className="text-purple-400 bg-slate-950">P2428</option>
                  <option value="P2426" className="text-emerald-400 bg-slate-950">P2426</option>
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">Source / Row System</label>
                <select 
                  value={primarySystem} 
                  onChange={(e) => setPrimarySystem(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-800 rounded-lg px-2.5 py-1.5 text-xs font-mono text-slate-200 focus:outline-none focus:border-blue-500"
                >
                  <option value="">-- Choose Row Acronym --</option>
                  {subsystems
                    .filter(sys => !sourceContract || sys.contract_package === sourceContract)
                    .map(sys => (
                      <option key={sys.id} value={sys.id}>
                        {formatSubsystemOption(sys)}
                      </option>
                    ))
                  }
                </select>
              </div>
            </div>

            {/* COLUMN SYSTEMS CONTROL */}
            <div className="border-b border-slate-800/80 pb-3 space-y-2.5">
              <div>
                <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">Column Contract Package</label>
                <select 
                  value={targetContract} 
                  onChange={(e) => setTargetContract(e.target.value)}
                  className={`w-full bg-slate-950 border border-slate-800 rounded-lg px-2.5 py-1.5 text-xs font-bold focus:outline-none ${getContractColor(targetContract)}`}
                >
                  <option value="P2427" className="text-blue-400 bg-slate-950">P2427</option>
                  <option value="P2428" className="text-purple-400 bg-slate-950">P2428</option>
                  <option value="P2426" className="text-emerald-400 bg-slate-950">P2426</option>
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">Target / Column Systems</label>
                <select 
                  multiple
                  value={targetSystems}
                  onChange={(e) => setTargetSystems(Array.from(e.target.selectedOptions, o => o.value))}
                  className="w-full bg-slate-950 border border-slate-800 rounded-lg px-2.5 py-1.5 text-xs font-mono text-slate-200 h-28 focus:outline-none"
                >
                  {subsystems
                    .filter(sys => sys.id.toString() !== primarySystem)
                    .filter(sys => !targetContract || sys.contract_package === targetContract)
                    .map(sys => (
                      <option key={sys.id} value={sys.id}>
                        {formatSubsystemOption(sys)}
                      </option>
                    ))
                  }
                </select>
              </div>
            </div>

            {/* Level & Category Dropdown selectors */}
            <div className="space-y-3">
              <div>
                <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">Interface Level</label>
                <select 
                  value={interfaceLevel} 
                  onChange={(e) => setInterfaceLevel(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-800 rounded-lg px-2 py-1.5 text-xs text-slate-300"
                >
                  <option value="Level 1">Level 1</option>
                  <option value="Level 2">Level 2</option>
                  <option value="Level 3">Level 3</option>
                  <option value="Level 4">Level 4</option>
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">Interface Category</label>
                <select 
                  value={interfaceCategory} 
                  onChange={(e) => setInterfaceCategory(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-800 rounded-lg px-2 py-1.5 text-xs font-medium text-slate-200 focus:outline-none"
                >
                  <option value="Physical and Functional">🔵 Physical and Functional</option>
                  <option value="Only Physical">🟡 Only Physical</option>
                  <option value="Only Functional">🟢 Only Functional</option>
                </select>
              </div>
            </div>

            <button type="submit" className="w-full bg-blue-600 hover:bg-blue-500 text-white font-medium py-2 rounded-lg text-xs transition mt-1">
              Commit Grid Coordinates
            </button>
          </form>
        </div>

        {/* Right Side: Legend Block & Matrix Grid View */}
        <div className="flex-1 w-full flex flex-col gap-4">
          
          {/* 4-Tier structural criteria mapping legend */}
          <div className="bg-slate-900 border border-slate-800 rounded-xl p-4 shadow-xl grid grid-cols-1 xl:grid-cols-3 gap-4">
            <div className="xl:col-span-2">
              <h3 className="text-[10px] font-bold uppercase tracking-wider text-slate-400 mb-2">Matrix Hierarchy Definition</h3>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-x-4 gap-y-1.5 text-[11px]">
                <p className="flex items-start gap-2">
                  <span className="font-bold text-blue-400 min-w-[48px] shrink-0 font-mono">Level 1:</span> 
                  <span className="text-slate-300 leading-tight">Interface with the Operators / State Agencies <span className="text-blue-500 font-medium font-mono text-[10px]">(Ext)</span></span>
                </p>
                <p className="flex items-start gap-2">
                  <span className="font-bold text-purple-400 min-w-[48px] shrink-0 font-mono">Level 2:</span> 
                  <span className="text-slate-300 leading-tight">Interfaces between contractors <span className="text-purple-500 font-medium font-mono text-[10px]">(Ext)</span></span>
                </p>
                <p className="flex items-start gap-2">
                  <span className="font-bold text-amber-400 min-w-[48px] shrink-0 font-mono">Level 3:</span> 
                  <span className="text-slate-300 leading-tight">Subsystems within the Contract <span className="text-amber-600 font-medium font-mono text-[10px]">(Int)</span></span>
                </p>
                <p className="flex items-start gap-2">
                  <span className="font-bold text-teal-400 min-w-[48px] shrink-0 font-mono">Level 4:</span> 
                  <span className="text-slate-300 leading-tight">Interfaces within the same Subsystem <span className="text-teal-600 font-medium font-mono text-[10px]">(Int)</span></span>
                </p>
              </div>
            </div>
            
            <div className="border-t xl:border-t-0 xl:border-l border-slate-800 pt-3 xl:pt-0 xl:pl-4 flex flex-col justify-center">
              <h3 className="text-[10px] font-bold uppercase tracking-wider text-slate-400 mb-2">Interface Category Signatures</h3>
              <div className="flex flex-wrap gap-x-4 gap-y-2 text-[11px]">
                <div className="flex items-center gap-1.5">
                  <span className="w-2.5 h-2.5 rounded bg-blue-500 shrink-0"></span>
                  <span className="text-slate-300">Physical & Functional</span>
                </div>
                <div className="flex items-center gap-1.5">
                  <span className="w-2.5 h-2.5 rounded bg-yellow-500 shrink-0"></span>
                  <span className="text-slate-300">Only Physical</span>
                </div>
                <div className="flex items-center gap-1.5">
                  <span className="w-2.5 h-2.5 rounded bg-emerald-500 shrink-0"></span>
                  <span className="text-slate-300">Only Functional</span>
                </div>
              </div>
            </div>
          </div>

          {/* Matrix Layout Grid */}
          <div className="bg-slate-900 border border-slate-800 rounded-xl overflow-hidden shadow-xl flex flex-col">
            <div className="p-3 bg-slate-900/50 border-b border-slate-800 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
              <div>
                <span className="text-xs font-semibold text-slate-300 block">Live Subsystem Interface Grid Matrix</span>
                <span className="text-[10px] text-slate-500 font-mono">Full Name Dropdowns Active</span>
              </div>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={handleExportExcel}
                  className="bg-emerald-600 hover:bg-emerald-500 text-white font-semibold px-3 py-1.5 rounded-lg text-[11px] transition"
                >
                  Download Excel
                </button>
                <button
                  type="button"
                  onClick={handleExportWord}
                  className="bg-blue-600 hover:bg-blue-500 text-white font-semibold px-3 py-1.5 rounded-lg text-[11px] transition"
                >
                  Download Word
                </button>
              </div>
            </div>

            <div className="overflow-auto max-h-[520px] max-w-full bg-slate-950">
              <table className="border-collapse table-fixed select-none">
                <thead>
                  <tr className="bg-slate-900/80 sticky top-0 z-20">
                    <th className="w-24 min-w-[96px] p-2 bg-slate-900 border border-slate-800 text-slate-500 text-[10px] font-mono font-bold sticky left-0 z-30 text-center">
                      SRC ➔ TRG
                    </th>
                    {subsystems.map(colSys => (
                      <th key={colSys.id} className="w-16 min-w-[64px] p-2 border border-slate-800 text-slate-300 text-[11px] font-mono text-center font-bold bg-slate-900 truncate">
                        {getSubsystemCode(colSys)}
                        <span className={`text-[8px] font-bold block tracking-tight opacity-90 mt-0.5 ${getContractColor(colSys.contract_package)}`}>
                          {colSys.contract_package || 'System'}
                        </span>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800">
                  {subsystems.map(rowSys => (
                    <tr key={rowSys.id} className="hover:bg-slate-900/30">
                      <td className="p-2 border border-slate-800 font-mono text-[11px] font-bold text-slate-300 bg-slate-900 sticky left-0 z-10 text-center truncate">
                        {getSubsystemCode(rowSys)}
                        <span className={`text-[8px] font-bold block tracking-tight opacity-90 ${getContractColor(rowSys.contract_package)}`}>
                          {rowSys.contract_package || 'System'}
                        </span>
                      </td>

                      {subsystems.map(colSys => {
                        const isSelf = rowSys.id === colSys.id;
                        const cellInterface = getCellInterface(rowSys.id, colSys.id);

                        return (
                          <td 
                            key={colSys.id} 
                            className={`border border-slate-800 text-center transition-all p-1 h-12 relative text-[9px] font-mono ${
                              isSelf ? 'bg-slate-900/40 diagonal-stripes' : 'bg-slate-950'
                            }`}
                            title={cellInterface ? `${formatSubsystemOption(rowSys)} to ${formatSubsystemOption(colSys)}: ${cellInterface.interface_level}` : ''}
                          >
                            {isSelf ? (
                              <span className="text-slate-700 font-bold block scale-75">\</span>
                            ) : cellInterface ? (
                              <div 
                                className="absolute inset-1 rounded flex flex-col justify-center items-center text-white font-bold text-[9px] shadow-md transition transform hover:scale-105 px-0.5 overflow-hidden"
                                style={{ 
                                  backgroundColor: cellInterface.color_accent,
                                  color: cellInterface.color_accent === '#eab308' ? '#020617' : '#ffffff'
                                }}
                              >
                                <span className="text-[10px]">✔</span>
                                <span className="text-[7px] font-sans tracking-tighter opacity-90 truncate max-w-full leading-none mt-0.5">
                                  {cellInterface.interface_level}
                                </span>
                              </div>
                            ) : null}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

        </div>
      </div>

      {/* SUBSYSTEM CONFIGURATION PANELS */}
      <section className="bg-slate-900 border border-slate-800 rounded-xl p-5 shadow-2xl mt-4">
        <header className="border-b border-slate-800 pb-3 mb-4 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
          <div>
            <h2 className="text-base font-bold text-white flex items-center gap-2">
              <span className="w-2.5 h-2.5 bg-purple-500 rounded-sm"></span> Subsystem Registry
            </h2>
            <p className="text-xs text-slate-400">Add, edit, or delete subsystem records. You can type a full description with parentheses—like <i>Building Management System (BMS)</i>—and the dropdowns will display the full name while the matrix can still use the acronym code.</p>
          </div>
        </header>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6 items-start">
          
          {/* LEFT COLUMN: FORM */}
          <div className="bg-slate-950 border border-slate-800 rounded-xl p-4">
            <h3 className="text-xs font-bold uppercase tracking-wider text-slate-300 mb-3">
              {editingSubsystemId ? '📝 Edit Subsystem' : '➕ Register New Subsystem'}
            </h3>
            
            <form onSubmit={handleSaveSubsystem} className="space-y-4">
              <div>
                <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">
                  Subsystem Full Name / Acronym
                </label>
                <input 
                  type="text"
                  value={formName}
                  onChange={(e) => setFormName(e.target.value)}
                  placeholder="e.g., Building Management System (BMS)"
                  className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-2 text-xs font-mono text-slate-200 placeholder-slate-600 focus:outline-none focus:border-purple-500"
                />
              </div>

              <div>
                <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">Contract Code Assignment</label>
                <select 
                  value={formContract}
                  onChange={(e) => setFormContract(e.target.value)}
                  className={`w-full bg-slate-900 border border-slate-800 rounded-lg px-2.5 py-2 text-xs font-bold focus:outline-none ${getContractColor(formContract)}`}
                >
                  <option value="P2427" className="text-blue-400 bg-slate-950">P2427</option>
                  <option value="P2428" className="text-purple-400 bg-slate-950">P2428</option>
                  <option value="P2426" className="text-emerald-400 bg-slate-950">P2426</option>
                </select>
              </div>

              <div className="flex items-center gap-2 pt-1">
                <button 
                  type="submit" 
                  className="flex-1 bg-purple-600 hover:bg-purple-500 text-white font-semibold py-2 rounded-lg text-xs transition"
                >
                  {editingSubsystemId ? 'Update Subsystem' : 'Insert Subsystem'}
                </button>
                {editingSubsystemId && (
                  <button 
                    type="button" 
                    onClick={() => { setEditingSubsystemId(null); setFormName(''); }}
                    className="bg-slate-800 hover:bg-slate-700 text-slate-300 font-semibold py-2 px-3 rounded-lg text-xs transition"
                  >
                    Cancel
                  </button>
                )}
              </div>
            </form>
          </div>

          {/* RIGHT COLUMN: RE-DESIGNED LIVE ACRONYM LIST PANEL */}
          <div className="lg:col-span-2 bg-slate-950 border border-slate-800 rounded-xl overflow-hidden">
            <div className="p-3 bg-slate-900/50 border-b border-slate-800 font-mono text-[10px] font-bold tracking-wider text-slate-400">
              Active Registered Subsystems ({subsystems.length})
            </div>
            <div className="overflow-auto max-h-[290px]">
              <table className="w-full text-left text-xs border-collapse">
                <thead>
                  <tr className="border-b border-slate-800 bg-slate-900/20 text-slate-400 text-[10px] font-bold uppercase tracking-wider">
                    <th className="p-2.5 font-mono">Subsystem Display Name</th>
                    <th className="p-2.5 font-mono">Acronym Code</th>
                    <th className="p-2.5">Contract Code Assignment</th>
                    <th className="p-2.5 text-right">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800 font-mono">
                  {subsystems.length === 0 ? (
                    <tr>
                      <td colSpan={4} className="text-center p-6 text-slate-600 font-sans">No codes currently initialized in the database registry.</td>
                    </tr>
                  ) : (
                    subsystems.map(sys => (
                      <tr key={sys.id} className="hover:bg-slate-900/40 group">
                        <td className="p-2.5 font-bold text-white text-[12px] tracking-wide">
                          {sys.name}
                        </td>
                        <td className="p-2.5 font-bold text-slate-300 text-[11px]">
                          {getSubsystemCode(sys)}
                        </td>
                        <td className={`p-2.5 font-bold text-[11px] ${getContractColor(sys.contract_package)}`}>
                          {sys.contract_package}
                        </td>
                        <td className="p-2.5 text-right space-x-1.5 whitespace-nowrap font-sans">
                          <button 
                            onClick={() => startEditSubsystem(sys)}
                            className="text-[11px] text-purple-400 hover:text-purple-300 bg-purple-950/40 hover:bg-purple-950 border border-purple-900/50 px-2 py-0.5 rounded transition"
                          >
                            Edit
                          </button>
                          <button 
                            onClick={() => handleDeleteSubsystem(sys.id)}
                            className="text-[11px] text-rose-400 hover:text-rose-300 bg-rose-950/40 hover:bg-rose-950 border border-rose-900/50 px-2 py-0.5 rounded transition"
                          >
                            Delete
                          </button>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>

        </div>
      </section>

    </div>
  );
}