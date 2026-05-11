import React, { useState, useMemo, useEffect } from 'react';
import { 
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, 
  Cell, AreaChart, Area, PieChart, Pie
} from 'recharts';
import { 
  TrendingUp, AlertTriangle, Upload, 
  RefreshCcw, FileSpreadsheet, FileText, ChevronRight
} from 'lucide-react';

/**
 * S&OP ENGINE PRO - Versión Consolidada
 * Herramienta avanzada para pronóstico de ventas y análisis de S&OP.
 */

// Carga de librerías externas vía CDN
const loadExternalLib = (url, globalName) => {
  return new Promise((resolve) => {
    if (window[globalName]) return resolve(window[globalName]);
    const script = document.createElement('script');
    script.src = url;
    script.async = true;
    script.onload = () => resolve(window[globalName]);
    document.head.appendChild(script);
  });
};

const COLORS = ['#2563eb', '#f97316', '#10b981', '#8b5cf6', '#f43f5e'];

const App = () => {
  // --- Estados Principales ---
  const [step, setStep] = useState(1);
  const [activeSegment, setActiveSegment] = useState(1);
  const [isLibLoading, setIsLibLoading] = useState(false);
  const [rawData, setRawData] = useState([]);
  const [columns, setColumns] = useState([]);
  const [useCalculatedSales, setUseCalculatedSales] = useState(false);

  const [mapping, setMapping] = useState({
    fecha: '', ventas: '', cantidad: '', precio: '',   
    producto: '', pais: '', vendedor: '', metodoPago: '', cliente: ''
  });

  const [config, setConfig] = useState({
    modelo: 'lineal',
    horizonte: 12,
    alpha: 0.3,
    filtroProducto: 'Todos',
    filtroPais: 'Todos'
  });

  // Utilidad para parsear fechas de Excel (Serial a Date)
  const parseExcelDate = (val) => {
    if (!val) return null;
    if (typeof val === 'number') {
      return new Date(Math.round((val - 25569) * 86400 * 1000));
    }
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  };

  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setIsLibLoading(true);
    try {
      const XLSX = await loadExternalLib("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js", "XLSX");
      const reader = new FileReader();

      reader.onload = (event) => {
        try {
          const buffer = new Uint8Array(event.target.result);
          const workbook = XLSX.read(buffer, { type: 'array', cellDates: false });
          const worksheet = workbook.Sheets[workbook.SheetNames[0]];
          const data = XLSX.utils.sheet_to_json(worksheet, { defval: null });
          
          if (data.length > 0) {
            setColumns(Object.keys(data[0]));
            setRawData(data);
            setStep(2);
          }
        } catch (err) {
          console.error("Error procesando Excel:", err);
        } finally {
          setIsLibLoading(false);
        }
      };
      reader.readAsArrayBuffer(file);
    } catch (err) {
      console.error("Error cargando librería XLSX:", err);
      setIsLibLoading(false);
    }
  };

  // Motor de Análisis y Pronóstico
  const analysis = useMemo(() => {
    if (rawData.length === 0 || !mapping.fecha) return null;

    let processed = rawData.map(d => {
      let v = 0;
      if (useCalculatedSales) {
        v = (Number(d[mapping.cantidad]) || 0) * (Number(d[mapping.precio]) || 0);
      } else {
        v = Number(d[mapping.ventas]) || 0;
      }
      const dateVal = parseExcelDate(d[mapping.fecha]);
      return { ...d, __val: v, __date: dateVal };
    }).filter(d => d.__date !== null);

    let filtered = processed.filter(d => {
      const matchP = config.filtroProducto === 'Todos' || String(d[mapping.producto]) === config.filtroProducto;
      const matchC = config.filtroPais === 'Todos' || String(d[mapping.pais]) === config.filtroPais;
      return matchP && matchC;
    });

    if (filtered.length === 0) return null;
    filtered.sort((a, b) => a.__date - b.__date);

    const monthlyMap = {};
    filtered.forEach(d => {
      const year = d.__date.getFullYear();
      const month = d.__date.getMonth();
      const key = `${year}-${String(month + 1).padStart(2, '0')}`;
      monthlyMap[key] = (monthlyMap[key] || 0) + d.__val;
    });

    const historicalData = Object.keys(monthlyMap).sort().map(k => ({
      fecha: k,
      ventas: Number(monthlyMap[k].toFixed(2))
    }));

    const vals = historicalData.map(h => h.ventas);
    const n = vals.length;
    if (n === 0) return null;

    let projection = [];
    if (config.modelo === 'lineal') {
      let xSum=0, ySum=0, xySum=0, xxSum=0;
      vals.forEach((y, x) => {
        xSum += x; ySum += y; xySum += x*y; xxSum += x*x;
      });
      const denom = (n * xxSum - xSum * xSum);
      const slope = denom !== 0 ? (n * xySum - xSum * ySum) / denom : 0;
      const intercept = (ySum - slope * xSum) / n;
      for(let i=1; i<=config.horizonte; i++) {
        projection.push(Math.max(0, intercept + slope * (n - 1 + i)));
      }
    } else if (config.modelo === 'promedio') {
      const avg = vals.slice(-3).reduce((a,b) => a+b, 0) / Math.min(n, 3);
      for(let i=1; i<=config.horizonte; i++) projection.push(avg);
    } else if (config.modelo === 'suavizacion') {
      let s = vals[0] || 0;
      const a = config.alpha;
      for(let i=1; i<n; i++) s = a * vals[i] + (1 - a) * s;
      for(let i=1; i<=config.horizonte; i++) projection.push(s);
    }

    const lastDateStr = historicalData[n-1].fecha;
    const [ly, lm] = lastDateStr.split('-').map(Number);
    const chartData = historicalData.map((h, i) => ({
      fecha: h.fecha,
      ventas: h.ventas,
      proyectado: i === n - 1 ? h.ventas : null
    }));

    projection.forEach((p, i) => {
      const d = new Date(ly, lm - 1 + (i + 1), 1);
      chartData.push({
        fecha: `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`,
        ventas: null,
        proyectado: Number(p.toFixed(2))
      });
    });

    const getTop = (key) => {
      if (!mapping[key]) return [];
      const counts = {};
      filtered.forEach(d => {
        const val = d[mapping[key]] || 'Otro';
        counts[val] = (counts[val] || 0) + d.__val;
      });
      return Object.entries(counts)
        .sort((a,b) => b[1] - a[1])
        .slice(0, 5)
        .map(e => ({ name: String(e[0]), value: e[1] }));
    };

    const meanHist = vals.reduce((a, b) => a + b, 0) / n;
    const variance = vals.reduce((a, b) => a + Math.pow(b - meanHist, 2), 0) / n;

    return {
      metrics: {
        totalHist: vals.reduce((a,b)=>a+b, 0),
        totalProj: projection.reduce((a,b)=>a+b, 0),
        avgHist: meanHist,
        avgProj: projection.reduce((a, b) => a + b, 0) / projection.length,
        cv: ((Math.sqrt(variance) / meanHist) * 100 || 0).toFixed(1)
      },
      chartData,
      options: {
        productos: ['Todos', ...new Set(processed.map(d => String(d[mapping.producto])).filter(x => x && x !== 'null'))],
        paises: ['Todos', ...new Set(processed.map(d => String(d[mapping.pais])).filter(x => x && x !== 'null'))]
      },
      topVendedores: getTop('vendedor'),
      topMetodos: getTop('metodoPago'),
      topClientes: getTop('cliente')
    };
  }, [rawData, mapping, config, useCalculatedSales]);

  // Exportación a PDF con jsPDF y AutoTable
  const exportToPDF = async () => {
    if (!analysis) return;
    setIsLibLoading(true);
    try {
      const jspdfLib = await loadExternalLib("https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js", "jspdf");
      await loadExternalLib("https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.28/jspdf.plugin.autotable.min.js", "jspdf-autotable");
      
      const { jsPDF } = jspdfLib;
      const doc = new jsPDF();
      
      doc.setFillColor(15, 23, 42); 
      doc.rect(0, 0, 210, 40, 'F');
      doc.setTextColor(255, 255, 255);
      doc.setFontSize(22);
      doc.text('REPORTE S&OP ENGINE PRO', 15, 25);
      
      doc.setTextColor(15, 23, 42);
      doc.autoTable({
        startY: 50,
        head: [['Métrica de Negocio', 'Valor']],
        body: [
          ['Venta Histórica Total', `$${Math.round(analysis.metrics.totalHist).toLocaleString()}`],
          ['Pronóstico Total', `$${Math.round(analysis.metrics.totalProj).toLocaleString()}`],
          ['Promedio Histórico', `$${Math.round(analysis.metrics.avgHist).toLocaleString()}`],
          ['Promedio Proyectado', `$${Math.round(analysis.metrics.avgProj).toLocaleString()}`],
          ['Modelo Seleccionado', config.modelo.toUpperCase()],
          ['Horizonte', `${config.horizonte} Meses`]
        ],
        theme: 'striped',
        headStyles: { fillColor: [37, 99, 235] }
      });

      doc.save(`SOP_Report_${config.filtroProducto}_${Date.now()}.pdf`);
    } catch (err) {
      console.error("Error al generar PDF:", err);
    } finally {
      setIsLibLoading(false);
    }
  };

  // --- RENDERING ---

  // PASO 1: Upload
  if (step === 1) return (
    <div className="min-h-screen bg-slate-900 flex items-center justify-center p-6 text-white">
      <div className="max-w-xl w-full bg-slate-800 p-10 rounded-[40px] border border-slate-700 shadow-2xl text-center space-y-8 animate-in fade-in zoom-in duration-500">
        <div className="bg-blue-600 w-20 h-20 rounded-3xl flex items-center justify-center mx-auto shadow-lg shadow-blue-900/50">
          <Upload className="w-10 h-10" />
        </div>
        <div>
          <h1 className="text-4xl font-black italic tracking-tighter">S&OP ENGINE PRO</h1>
          <p className="text-slate-400 font-bold text-sm uppercase tracking-widest mt-2">Sistema Universal de Pronóstico</p>
        </div>
        <label className="border-2 border-dashed border-slate-700 rounded-[32px] p-12 block cursor-pointer hover:border-blue-500 hover:bg-slate-700/30 transition-all">
          <input type="file" className="hidden" onChange={handleFileUpload} accept=".csv, .xlsx, .xls" />
          <FileSpreadsheet className="w-14 h-14 text-slate-600 mx-auto mb-4" />
          <span className="text-sm font-black uppercase text-slate-400">Seleccionar Excel o CSV</span>
        </label>
        {isLibLoading && <p className="text-blue-400 font-bold animate-pulse">Iniciando motor de datos...</p>}
      </div>
    </div>
  );

  // PASO 2: Mapping
  if (step === 2) return (
    <div className="min-h-screen bg-slate-50 p-10 flex flex-col items-center">
      <div className="max-w-4xl w-full bg-white rounded-[40px] shadow-2xl border border-slate-200 overflow-hidden animate-in slide-in-from-bottom-8 duration-500">
        <div className="bg-slate-900 p-8 text-white flex justify-between items-center">
          <div>
            <h2 className="text-2xl font-black uppercase italic">Configuración de Datos</h2>
            <p className="text-[10px] text-slate-400 font-bold uppercase tracking-widest">Asigne las columnas de su archivo para el análisis</p>
          </div>
          <div className="flex bg-slate-800 p-1 rounded-2xl">
            <button onClick={() => setUseCalculatedSales(false)} className={`px-6 py-2 rounded-xl text-[10px] font-black transition-all ${!useCalculatedSales ? 'bg-blue-600 text-white' : 'text-slate-400'}`}>VENTA DIRECTA</button>
            <button onClick={() => setUseCalculatedSales(true)} className={`px-6 py-2 rounded-xl text-[10px] font-black transition-all ${useCalculatedSales ? 'bg-blue-600 text-white' : 'text-slate-400'}`}>CALCULAR Q*P</button>
          </div>
        </div>
        
        <div className="p-10 grid grid-cols-1 md:grid-cols-2 gap-10">
          <div className="space-y-6">
            <h3 className="text-xs font-black text-blue-600 uppercase border-b pb-2">Campos Críticos</h3>
            <div className="space-y-4">
              <div className="space-y-2">
                <label className="text-[10px] font-black text-slate-400 uppercase">Columna de Fecha (Indispensable)</label>
                <select className="w-full bg-slate-50 border-2 border-slate-100 p-3 rounded-2xl font-bold text-sm focus:border-blue-500 outline-none" value={mapping.fecha} onChange={e => setMapping({...mapping, fecha: e.target.value})}>
                  <option value="">Seleccionar...</option>
                  {columns.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>
              {!useCalculatedSales ? (
                <div className="space-y-2">
                  <label className="text-[10px] font-black text-slate-400 uppercase">Ventas ($)</label>
                  <select className="w-full bg-slate-50 border-2 border-slate-100 p-3 rounded-2xl font-bold text-sm focus:border-blue-500 outline-none" value={mapping.ventas} onChange={e => setMapping({...mapping, ventas: e.target.value})}>
                    <option value="">Seleccionar...</option>
                    {columns.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>
              ) : (
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-2">
                    <label className="text-[10px] font-black text-slate-400 uppercase">Cantidad</label>
                    <select className="w-full bg-slate-50 border-2 border-slate-100 p-3 rounded-2xl font-bold text-sm" value={mapping.cantidad} onChange={e => setMapping({...mapping, cantidad: e.target.value})}>
                      <option value="">Columna...</option>
                      {columns.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </div>
                  <div className="space-y-2">
                    <label className="text-[10px] font-black text-slate-400 uppercase">Precio</label>
                    <select className="w-full bg-slate-50 border-2 border-slate-100 p-3 rounded-2xl font-bold text-sm" value={mapping.precio} onChange={e => setMapping({...mapping, precio: e.target.value})}>
                      <option value="">Columna...</option>
                      {columns.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </div>
                </div>
              )}
            </div>
          </div>
          
          <div className="space-y-6">
            <h3 className="text-xs font-black text-slate-400 uppercase border-b pb-2">Dimensiones Adicionales</h3>
            <div className="grid grid-cols-1 gap-3">
              {['producto', 'pais', 'vendedor', 'metodoPago', 'cliente'].map(k => (
                <div key={k} className="flex items-center gap-2">
                  <div className="w-2 h-2 rounded-full bg-slate-200"></div>
                  <select className="flex-1 bg-slate-50 border border-slate-200 p-2 rounded-xl font-bold text-[11px]" value={mapping[k]} onChange={e => setMapping({...mapping, [k]: e.target.value})}>
                    <option value="">Ignorar {k.toUpperCase()}</option>
                    {columns.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>
              ))}
            </div>
          </div>
        </div>

        <div className="p-8 bg-slate-50 border-t flex justify-between items-center px-12">
          <button onClick={() => setStep(1)} className="text-slate-400 font-black text-xs uppercase hover:text-slate-600 transition-colors">Volver</button>
          <button disabled={!mapping.fecha} onClick={() => setStep(3)} className="px-20 bg-blue-600 disabled:opacity-50 text-white py-5 rounded-[24px] font-black uppercase shadow-xl hover:scale-105 active:scale-95 transition-all">PROCESAR MODELO</button>
        </div>
      </div>
    </div>
  );

  // DASHBOARD PRINCIPAL
  return (
    <div className="flex h-screen bg-slate-100 overflow-hidden text-slate-900">
      {/* Sidebar de Configuración */}
      <aside className="w-80 bg-white border-r border-slate-200 p-8 space-y-8 shrink-0 flex flex-col shadow-xl z-20">
        <div className="flex items-center gap-3 text-blue-600">
          <TrendingUp className="w-8 h-8" />
          <h1 className="font-black text-xl italic uppercase tracking-tighter text-slate-900">S&OP ENGINE</h1>
        </div>
        
        <div className="space-y-8 flex-1 overflow-y-auto pr-2">
          <section className="space-y-3">
            <p className="text-[10px] font-black text-blue-600 uppercase tracking-widest">Segmentación</p>
            <div className="space-y-2">
              <div className="space-y-1">
                <label className="text-[9px] font-bold text-slate-400 ml-1">PRODUCTO</label>
                <select className="w-full border-2 border-slate-100 p-3 rounded-2xl text-xs font-black bg-slate-50" value={config.filtroProducto} onChange={e => setConfig({...config, filtroProducto: e.target.value})}>
                  {analysis?.options.productos.map(p => <option key={p} value={p}>{p}</option>)}
                </select>
              </div>
              <div className="space-y-1">
                <label className="text-[9px] font-bold text-slate-400 ml-1">PAÍS / REGIÓN</label>
                <select className="w-full border-2 border-slate-100 p-3 rounded-2xl text-xs font-black bg-slate-50" value={config.filtroPais} onChange={e => setConfig({...config, filtroPais: e.target.value})}>
                  {analysis?.options.paises.map(p => <option key={p} value={p}>{p}</option>)}
                </select>
              </div>
            </div>
          </section>

          <section className="space-y-3">
            <p className="text-[10px] font-black text-orange-600 uppercase tracking-widest">Modelo Matemático</p>
            <div className="space-y-2">
              {[
                { id: 'lineal', name: 'Regresión Lineal' },
                { id: 'promedio', name: 'Promedio Móvil' },
                { id: 'suavizacion', name: 'Suavización Exp.' }
              ].map(m => (
                <button 
                  key={m.id} 
                  onClick={() => setConfig({...config, modelo: m.id})} 
                  className={`w-full p-4 rounded-2xl text-[10px] font-black uppercase text-left border-2 transition-all ${config.modelo === m.id ? 'bg-orange-600 text-white border-orange-600 shadow-lg shadow-orange-200' : 'bg-white text-slate-400 border-slate-100 hover:border-orange-200'}`}
                >
                  {m.name}
                </button>
              ))}
            </div>
          </section>

          <section className="pt-4 border-t">
            <div className="flex justify-between items-center mb-4">
              <p className="text-[10px] font-black text-slate-400 uppercase">Horizonte</p>
              <span className="bg-slate-100 px-3 py-1 rounded-full text-[10px] font-black text-slate-600">{config.horizonte} meses</span>
            </div>
            <input type="range" min="1" max="24" className="w-full accent-blue-600 h-2 bg-slate-100 rounded-lg appearance-none cursor-pointer" value={config.horizonte} onChange={e => setConfig({...config, horizonte: Number(e.target.value)})} />
          </section>
        </div>

        <button onClick={() => setStep(1)} className="p-4 border-2 border-slate-100 rounded-2xl text-[10px] font-black text-slate-400 flex items-center justify-center gap-2 hover:bg-red-50 hover:text-red-500 hover:border-red-100 transition-all uppercase tracking-widest mt-auto">
          <RefreshCcw size={14}/> Reiniciar Proceso
        </button>
      </aside>

      {/* Contenido Principal */}
      <main className="flex-1 overflow-y-auto flex flex-col">
        <header className="bg-white border-b border-slate-200 px-12 h-24 flex items-center justify-between sticky top-0 z-10 shrink-0">
          <div className="flex items-center gap-12 h-full">
            {[ {id: 1, label: 'Panel Visual'}, {id: 2, label: 'Tabla de Datos'}, {id: 3, label: 'Distribución'} ].map((t) => (
              <button key={t.id} onClick={() => setActiveSegment(t.id)} className={`h-full border-b-4 text-[11px] font-black uppercase tracking-widest transition-all ${activeSegment === t.id ? 'border-blue-600 text-blue-600' : 'border-transparent text-slate-400 hover:text-slate-600'}`}>{t.label}</button>
            ))}
          </div>
          <button onClick={exportToPDF} disabled={isLibLoading} className="flex items-center gap-2 bg-slate-900 text-white px-6 py-3 rounded-2xl hover:bg-blue-600 disabled:opacity-50 transition-all text-[11px] font-black uppercase">
            <FileText size={16} />
            <span>{isLibLoading ? 'GENERANDO...' : 'EXPORTAR PDF'}</span>
          </button>
        </header>

        <div className="p-12">
          {!analysis ? (
             <div className="py-48 text-center flex flex-col items-center gap-4 text-slate-300">
                <AlertTriangle size={64} />
                <p className="font-black uppercase tracking-widest text-lg">No hay datos suficientes para esta combinación</p>
             </div>
          ) : (
            <div className="animate-in fade-in duration-700">
              {activeSegment === 1 && (
                <div className="space-y-10">
                  {/* KPI Cards */}
                  <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-5 gap-6">
                    {[
                      { label: 'Histórica Total', val: analysis.metrics.totalHist, color: 'text-slate-900', bg: 'bg-white' },
                      { label: 'Pronóstico Total', val: analysis.metrics.totalProj, color: 'text-orange-600', bg: 'bg-orange-50/30' },
                      { label: 'Promedio Hist.', val: analysis.metrics.avgHist, color: 'text-blue-500', bg: 'bg-white' },
                      { label: 'Promedio Proy.', val: analysis.metrics.avgProj, color: 'text-orange-500', bg: 'bg-white' },
                      { label: 'Variabilidad (CV)', val: analysis.metrics.cv, suffix: '%', color: 'text-slate-500', bg: 'bg-slate-50/50' }
                    ].map((k, i) => (
                      <div key={i} className={`${k.bg} p-6 rounded-[32px] border border-slate-200 shadow-sm hover:shadow-md transition-all group`}>
                        <p className="text-[8px] font-black text-slate-400 uppercase mb-3 tracking-widest">{k.label}</p>
                        <h4 className={`text-xl font-black ${k.color}`}>{k.suffix === '%' ? '' : '$'}{Math.round(k.val).toLocaleString()}{k.suffix || ''}</h4>
                        <div className="h-1 w-8 bg-slate-100 mt-4 group-hover:w-full transition-all duration-500 rounded-full"></div>
                      </div>
                    ))}
                  </div>

                  {/* Main Chart */}
                  <div className="bg-white p-10 rounded-[40px] border border-slate-200 shadow-sm">
                    <div className="flex justify-between items-center mb-10">
                      <h3 className="text-xs font-black uppercase tracking-widest flex items-center gap-2">
                        <span className="w-2 h-6 bg-blue-600 rounded-full"></span> 
                        Tendencia vs Pronóstico
                      </h3>
                      <div className="flex gap-4 text-[10px] font-bold">
                        <div className="flex items-center gap-2"><span className="w-3 h-3 rounded-full bg-blue-600"></span> REAL</div>
                        <div className="flex items-center gap-2"><span className="w-3 h-3 rounded-full border-2 border-orange-500 border-dashed"></span> PRONÓSTICO</div>
                      </div>
                    </div>
                    <div className="h-[450px] w-full">
                      <ResponsiveContainer width="100%" height="100%">
                        <AreaChart data={analysis.chartData}>
                          <defs>
                            <linearGradient id="colorReal" x1="0" y1="0" x2="0" y2="1">
                              <stop offset="5%" stopColor="#2563eb" stopOpacity={0.1}/>
                              <stop offset="95%" stopColor="#2563eb" stopOpacity={0}/>
                            </linearGradient>
                            <linearGradient id="colorProj" x1="0" y1="0" x2="0" y2="1">
                              <stop offset="5%" stopColor="#f97316" stopOpacity={0.1}/>
                              <stop offset="95%" stopColor="#f97316" stopOpacity={0}/>
                            </linearGradient>
                          </defs>
                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                          <XAxis dataKey="fecha" tick={{fontSize: 9, fontWeight: 900}} tickLine={false} axisLine={false} />
                          <YAxis tick={{fontSize: 9, fontWeight: 900}} tickFormatter={(v) => `$${(v/1000)}k`} axisLine={false} tickLine={false} />
                          <Tooltip 
                            contentStyle={{ borderRadius: '20px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)', padding: '15px' }}
                            itemStyle={{ fontSize: '12px', fontWeight: 'bold' }}
                          />
                          <Area type="monotone" dataKey="ventas" stroke="#2563eb" fillOpacity={1} fill="url(#colorReal)" strokeWidth={4} connectNulls name="Venta Real" />
                          <Area type="monotone" dataKey="proyectado" stroke="#f97316" fillOpacity={1} fill="url(#colorProj)" strokeWidth={4} strokeDasharray="8 8" connectNulls name="Proyección" />
                        </AreaChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                </div>
              )}

              {activeSegment === 2 && (
                <div className="bg-white rounded-[40px] border border-slate-200 overflow-hidden shadow-sm">
                  <div className="p-8 border-b bg-slate-50 flex justify-between items-center">
                    <h3 className="text-xs font-black uppercase tracking-widest">Histórico Detallado</h3>
                    <span className="text-[10px] font-bold text-slate-400 uppercase">{analysis.chartData.length} Periodos Analizados</span>
                  </div>
                  <table className="w-full text-left text-[11px] font-bold">
                    <thead className="bg-white border-b border-slate-100 text-slate-400 font-black uppercase tracking-widest">
                      <tr>
                        <th className="p-6 text-slate-900">Periodo de Venta</th>
                        <th className="p-6">Venta Real ($)</th>
                        <th className="p-6 text-orange-600">Proyección Modelo ($)</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-50">
                      {analysis.chartData.map((d, i) => (
                        <tr key={i} className="hover:bg-slate-50 transition-colors group">
                          <td className="p-6 text-slate-900 uppercase flex items-center gap-2">
                             <ChevronRight size={14} className="text-slate-300 opacity-0 group-hover:opacity-100 transition-opacity" />
                             {d.fecha}
                          </td>
                          <td className="p-6 text-blue-600">{d.ventas ? `$${d.ventas.toLocaleString()}` : '—'}</td>
                          <td className="p-6 text-orange-600 font-black">{d.proyectado ? `$${d.proyectado.toLocaleString()}` : '—'}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}

              {activeSegment === 3 && (
                <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
                   <div className="bg-white p-10 rounded-[40px] border border-slate-200 shadow-sm space-y-8">
                      <h4 className="text-[10px] font-black uppercase text-slate-400 tracking-widest border-b pb-4">Top 5 Vendedores</h4>
                      <div className="space-y-6">
                         {analysis.topVendedores.map((v, i) => (
                           <div key={i} className="flex flex-col gap-2">
                             <div className="flex justify-between text-[11px] font-black">
                               <span>{v.name}</span>
                               <span className="text-blue-600">${v.value.toLocaleString()}</span>
                             </div>
                             <div className="w-full bg-slate-100 h-2 rounded-full overflow-hidden">
                               <div 
                                 className="bg-blue-600 h-full rounded-full transition-all duration-1000" 
                                 style={{width: `${(v.value / analysis.metrics.totalHist) * 100}%`}}
                               ></div>
                             </div>
                           </div>
                         ))}
                      </div>
                   </div>

                   <div className="bg-white p-10 rounded-[40px] border border-slate-200 shadow-sm flex flex-col">
                      <h4 className="text-[10px] font-black uppercase text-slate-400 w-full mb-8 tracking-widest border-b pb-4">Mezcla de Pago</h4>
                      <div className="flex-1 flex flex-col items-center justify-center">
                        <div className="w-full h-64">
                          <ResponsiveContainer width="100%" height="100%">
                            <PieChart>
                              <Pie data={analysis.topMetodos} innerRadius={65} outerRadius={90} paddingAngle={10} dataKey="value">
                                {analysis.topMetodos.map((_, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} cornerRadius={8} />)}
                              </Pie>
                              <Tooltip 
                                contentStyle={{ borderRadius: '15px', border: 'none', boxShadow: '0 5px 10px rgba(0,0,0,0.1)' }}
                                itemStyle={{ fontSize: '10px', fontWeight: 'bold' }}
                              />
                            </PieChart>
                          </ResponsiveContainer>
                        </div>
                        <div className="grid grid-cols-2 gap-x-6 gap-y-2 mt-4">
                           {analysis.topMetodos.map((m, i) => (
                             <div key={i} className="flex items-center gap-2 text-[10px] font-bold text-slate-500 uppercase">
                               <div className="w-2 h-2 rounded-full" style={{backgroundColor: COLORS[i % COLORS.length]}}></div>
                               <span className="truncate max-w-[80px]">{m.name}</span>
                             </div>
                           ))}
                        </div>
                      </div>
                   </div>

                   <div className="bg-white p-10 rounded-[40px] border border-slate-200 shadow-sm space-y-8">
                      <h4 className="text-[10px] font-black uppercase text-slate-400 tracking-widest border-b pb-4">Principales Clientes</h4>
                      <div className="space-y-3">
                         {analysis.topClientes.length > 0 ? analysis.topClientes.map((c, i) => (
                           <div key={i} className="flex justify-between items-center p-4 bg-slate-50 rounded-2xl border border-slate-100 hover:border-blue-200 hover:bg-white transition-all shadow-sm">
                             <div className="flex items-center gap-3">
                               <span className="w-6 h-6 rounded-full bg-slate-200 flex items-center justify-center text-[9px] font-black text-slate-500">{i+1}</span>
                               <span className="text-[11px] font-black truncate max-w-[120px]">{c.name}</span>
                             </div>
                             <span className="text-[10px] font-black text-blue-600 bg-blue-50 px-3 py-1 rounded-full">${c.value.toLocaleString()}</span>
                           </div>
                         )) : <p className="text-center text-slate-300 py-10">Sin datos de clientes</p>}
                      </div>
                   </div>
                </div>
              )}
            </div>
          )}
        </div>
      </main>
    </div>
  );
};

export default App;
