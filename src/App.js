import React, { useState, useMemo, useEffect } from 'react';
import { 
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, 
  Cell, AreaChart, Area, PieChart, Pie
} from 'recharts';
import { 
  TrendingUp, AlertTriangle, Upload, 
  RefreshCcw, FileSpreadsheet, FileText
} from 'lucide-react';

// Carga de librerías externas (XLSX, jsPDF) desde CDN
const loadExternalLib = (url, globalName) => {
  return new Promise((resolve) => {
    if (window[globalName]) return resolve(window[globalName]);
    const script = document.createElement('script');
    script.src = url;
    script.onload = () => resolve(window[globalName]);
    document.head.appendChild(script);
  });
};

const COLORS = ['#2563eb', '#f97316', '#10b981', '#8b5cf6', '#f43f5e'];

const App = () => {
  // --- INYECCIÓN DE TAILWIND PARA CODESANDBOX ---
  useEffect(() => {
    if (!document.getElementById('tailwind-cdn')) {
      const script = document.createElement('script');
      script.id = 'tailwind-cdn';
      script.src = 'https://cdn.tailwindcss.com';
      document.head.appendChild(script);
    }
  }, []);

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
    const XLSX = await loadExternalLib("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js", "XLSX");
    const reader = new FileReader();

    reader.onload = (event) => {
      try {
        const buffer = new Uint8Array(event.target.result);
        const workbook = XLSX.read(buffer, { type: 'array', cellDates: false });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(worksheet, { defval: null });
        if (data.length > 0) setColumns(Object.keys(data[0]));
        setRawData(data);
        setStep(2);
      } catch (err) {
        console.error("Error cargando archivo:", err);
      } finally {
        setIsLibLoading(false);
      }
    };
    reader.readAsArrayBuffer(file);
  };

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

  const exportToPDF = async () => {
    if (!analysis) return;
    setIsLibLoading(true);
    const { jsPDF } = await loadExternalLib("https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js", "jspdf");
    await loadExternalLib("https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.28/jspdf.plugin.autotable.min.js", "jspdf-autotable");
    
    const doc = new jsPDF();
    doc.setFillColor(15, 23, 42); 
    doc.rect(0, 0, 210, 40, 'F');
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(22);
    doc.text('REPORTE S&OP ENGINE PRO', 15, 25);
    
    doc.setTextColor(15, 23, 42);
    doc.autoTable({
      startY: 50,
      head: [['Metrica de Negocio', 'Valor']],
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

    doc.save(`SOP_Report_${config.filtroProducto}_${new Date().getTime()}.pdf`);
    setIsLibLoading(false);
  };

  if (step === 1) return (
    <div className="min-h-screen bg-slate-900 flex items-center justify-center p-6 text-white font-sans">
      <div className="max-w-xl w-full bg-slate-800 p-10 rounded-[40px] border border-slate-700 shadow-2xl text-center space-y-8">
        <div className="bg-blue-600 w-20 h-20 rounded-3xl flex items-center justify-center mx-auto shadow-lg shadow-blue-900/50">
          <Upload className="w-10 h-10" />
        </div>
        <h1 className="text-4xl font-black italic tracking-tighter">S&OP ENGINE PRO</h1>
        <p className="text-slate-400 font-bold text-sm uppercase tracking-widest">Sistema Universal de Pronóstico</p>
        <label className="border-2 border-dashed border-slate-700 rounded-[32px] p-12 block cursor-pointer hover:border-blue-500 transition-all bg-slate-800/50">
          <input type="file" className="hidden" onChange={handleFileUpload} accept=".csv, .xlsx, .xls" />
          <FileSpreadsheet className="w-14 h-14 text-slate-600 mx-auto mb-4" />
          <span className="text-sm font-black uppercase text-slate-400">Seleccionar Archivo</span>
        </label>
      </div>
    </div>
  );

  if (step === 2) return (
    <div className="min-h-screen bg-slate-50 p-10 flex flex-col items-center font-sans">
      <div className="max-w-4xl w-full bg-white rounded-[40px] shadow-2xl border border-slate-200 overflow-hidden">
        <div className="bg-slate-900 p-8 text-white flex justify-between items-center">
          <div>
            <h2 className="text-2xl font-black uppercase italic">Configuración de Datos</h2>
            <p className="text-[10px] text-slate-400 font-bold uppercase tracking-widest">Asigne las columnas de su archivo</p>
          </div>
          <div className="flex bg-slate-800 p-1 rounded-2xl">
            <button onClick={() => setUseCalculatedSales(false)} className={`px-6 py-2 rounded-xl text-[10px] font-black transition-all ${!useCalculatedSales ? 'bg-blue-600 text-white' : 'text-slate-400'}`}>VENTA DIRECTA</button>
            <button onClick={() => setUseCalculatedSales(true)} className={`px-6 py-2 rounded-xl text-[10px] font-black transition-all ${useCalculatedSales ? 'bg-blue-600 text-white' : 'text-slate-400'}`}>CALCULAR Q*P</button>
          </div>
        </div>
        <div className="p-10 grid grid-cols-2 gap-10">
          <div className="space-y-6">
            <h3 className="text-xs font-black text-blue-600 uppercase">Campos Críticos</h3>
            <div className="space-y-4">
              <div className="space-y-2">
                <label className="text-[10px] font-black text-slate-400 uppercase">Columna de Fecha</label>
                <select className="w-full bg-slate-50 border-2 border-slate-100 p-3 rounded-2xl font-bold text-sm" value={mapping.fecha} onChange={e => setMapping({...mapping, fecha: e.target.value})}>
                  <option value="">Seleccionar...</option>
                  {columns.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>
              {!useCalculatedSales ? (
                <div className="space-y-2">
                  <label className="text-[10px] font-black text-slate-400 uppercase">Ventas ($)</label>
                  <select className="w-full bg-slate-50 border-2 border-slate-100 p-3 rounded-2xl font-bold text-sm" value={mapping.ventas} onChange={e => setMapping({...mapping, ventas: e.target.value})}>
                    <option value="">Seleccionar...</option>
                    {columns.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>
              ) : (
                <div className="grid grid-cols-2 gap-4">
                  <select className="w-full bg-slate-50 border-2 border-slate-100 p-3 rounded-2xl font-bold text-sm" value={mapping.cantidad} onChange={e => setMapping({...mapping, cantidad: e.target.value})}>
                    <option value="">Cantidad</option>
                    {columns.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                  <select className="w-full bg-slate-50 border-2 border-slate-100 p-3 rounded-2xl font-bold text-sm" value={mapping.precio} onChange={e => setMapping({...mapping, precio: e.target.value})}>
                    <option value="">Precio</option>
                    {columns.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>
              )}
            </div>
          </div>
          <div className="space-y-6">
            <h3 className="text-xs font-black text-slate-400 uppercase">Dimensiones de Negocio</h3>
            {['producto', 'pais', 'vendedor', 'metodoPago', 'cliente'].map(k => (
              <select key={k} className="w-full bg-slate-50 border-2 border-slate-100 p-2 rounded-xl font-bold text-[11px]" value={mapping[k]} onChange={e => setMapping({...mapping, [k]: e.target.value})}>
                <option value="">Ignorar {k}</option>
                {columns.map(c => <option key={c} value={c}>{c}</option>)}
              </select>
            ))}
          </div>
        </div>
        <div className="p-8 bg-slate-50 border-t flex justify-center">
          <button disabled={!mapping.fecha} onClick={() => setStep(3)} className="px-20 bg-blue-600 disabled:bg-slate-300 text-white py-5 rounded-[24px] font-black uppercase shadow-xl hover:scale-105 transition-all">PROCESAR MODELO</button>
        </div>
      </div>
    </div>
  );

  return (
    <div className="flex h-screen bg-slate-100 overflow-hidden font-sans">
      <aside className="w-80 bg-white border-r border-slate-200 p-8 space-y-8 shrink-0 flex flex-col shadow-lg z-20">
        <div className="flex items-center gap-3 text-blue-600">
          <TrendingUp className="w-8 h-8" />
          <h1 className="font-black text-xl italic uppercase tracking-tighter text-slate-900">S&OP ENGINE</h1>
        </div>
        <div className="space-y-6">
          <section className="space-y-3">
            <p className="text-[10px] font-black text-blue-600 uppercase tracking-widest">Segmentación</p>
            <div className="space-y-2">
              <select className="w-full border-2 border-slate-100 p-3 rounded-2xl text-xs font-black bg-slate-50" value={config.filtroProducto} onChange={e => setConfig({...config, filtroProducto: e.target.value})}>
                {analysis?.options.productos.map(p => <option key={p} value={p}>{p}</option>)}
              </select>
              <select className="w-full border-2 border-slate-100 p-3 rounded-2xl text-xs font-black bg-slate-50" value={config.filtroPais} onChange={e => setConfig({...config, filtroPais: e.target.value})}>
                {analysis?.options.paises.map(p => <option key={p} value={p}>{p}</option>)}
              </select>
            </div>
          </section>
          <section className="space-y-3">
            <p className="text-[10px] font-black text-orange-600 uppercase tracking-widest">Modelo Matemático</p>
            <div className="space-y-2">
              {['lineal', 'promedio', 'suavizacion'].map(m => (
                <button key={m} onClick={() => setConfig({...config, modelo: m})} className={`w-full p-4 rounded-2xl text-[10px] font-black uppercase text-left border-2 transition-all ${config.modelo === m ? 'bg-orange-600 text-white border-orange-600 shadow-md shadow-orange-200' : 'bg-white text-slate-400 border-slate-100 hover:border-orange-200'}`}>{m === 'lineal' ? 'Regresión Lineal' : m === 'promedio' ? 'Promedio Móvil' : 'Suavización Exp.'}</button>
              ))}
            </div>
          </section>
          <section className="pt-4 border-t">
            <div className="flex justify-between items-center mb-4">
              <p className="text-[10px] font-black text-slate-400 uppercase">Horizonte</p>
              <span className="bg-slate-100 px-3 py-1 rounded-full text-[10px] font-black text-slate-600">{config.horizonte}m</span>
            </div>
            <input type="range" min="1" max="24" className="w-full accent-blue-600" value={config.horizonte} onChange={e => setConfig({...config, horizonte: Number(e.target.value)})} />
          </section>
        </div>
        <button onClick={() => setStep(1)} className="mt-auto p-4 border-2 border-slate-100 rounded-2xl text-[10px] font-black text-slate-400 flex items-center justify-center gap-2 hover:bg-red-50 hover:text-red-500 hover:border-red-100 transition-all uppercase tracking-widest"><RefreshCcw size={14}/> Reiniciar Proceso</button>
      </aside>

      <main className="flex-1 overflow-y-auto flex flex-col">
        <header className="bg-white border-b border-slate-200 px-12 h-24 flex items-center justify-between sticky top-0 z-10 font-black uppercase text-[11px] tracking-widest shrink-0">
          <div className="flex items-center gap-12 h-full">
            {[ {id: 1, label: 'Panel Visual'}, {id: 2, label: 'Tabla de Datos'}, {id: 3, label: 'Insights'} ].map((t) => (
              <button key={t.id} onClick={() => setActiveSegment(t.id)} className={`h-full border-b-4 transition-all ${activeSegment === t.id ? 'border-blue-600 text-blue-600' : 'border-transparent text-slate-400 hover:text-slate-600'}`}>{t.label}</button>
            ))}
          </div>
          <button onClick={exportToPDF} disabled={isLibLoading} className="flex items-center gap-2 bg-slate-900 text-white px-6 py-3 rounded-2xl hover:bg-blue-600 transition-colors">
            <FileText size={16} />
            <span>{isLibLoading ? 'PROCESANDO...' : 'EXPORTAR PDF'}</span>
          </button>
        </header>

        <div className="p-12 space-y-10">
          {!analysis ? (
             <div className="py-32 text-center text-slate-400 font-black uppercase">Sin datos para proyectar</div>
          ) : (
            <>
              {activeSegment === 1 && (
                <div className="space-y-10 animate-in fade-in slide-in-from-bottom-4 duration-500">
                  <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-5 gap-6">
                    {[
                      { label: 'Histórica Total', val: analysis.metrics.totalHist, color: 'text-slate-900' },
                      { label: 'Pronóstico Total', val: analysis.metrics.totalProj, color: 'text-orange-600' },
                      { label: 'Promedio Hist.', val: analysis.metrics.avgHist, color: 'text-blue-500' },
                      { label: 'Promedio Proy.', val: analysis.metrics.avgProj, color: 'text-orange-500' },
                      { label: 'Var (CV)', val: analysis.metrics.cv, suffix: '%', color: 'text-slate-400' }
                    ].map((k, i) => (
                      <div key={i} className="bg-white p-6 rounded-[32px] border border-slate-200 shadow-sm hover:shadow-md transition-shadow">
                        <p className="text-[8px] font-black text-slate-400 uppercase mb-3 tracking-widest">{k.label}</p>
                        <h4 className={`text-xl font-black ${k.color}`}>{k.suffix === '%' ? '' : '$'}{Math.round(k.val).toLocaleString()}{k.suffix || ''}</h4>
                      </div>
                    ))}
                  </div>

                  <div className="bg-white p-10 rounded-[40px] border border-slate-200 shadow-sm">
                    <h3 className="text-sm font-black uppercase tracking-widest mb-10">Gráfica de Tendencia y Pronóstico</h3>
                    <div className="h-[450px] w-full">
                      <ResponsiveContainer width="100%" height="100%">
                        <AreaChart data={analysis.chartData}>
                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                          <XAxis dataKey="fecha" tick={{fontSize: 9, fontWeight: 900}} />
                          <YAxis tick={{fontSize: 9, fontWeight: 900}} tickFormatter={(v) => `$${(v/1000)}k`} />
                          <Tooltip />
                          <Area type="monotone" dataKey="ventas" stroke="#2563eb" fill="#2563eb10" strokeWidth={4} connectNulls name="Real" />
                          <Area type="monotone" dataKey="proyectado" stroke="#f97316" fill="#f9731610" strokeWidth={4} strokeDasharray="8 8" connectNulls name="Proyección" />
                        </AreaChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                </div>
              )}

              {activeSegment === 2 && (
                <div className="bg-white rounded-[40px] border border-slate-200 overflow-hidden shadow-sm animate-in fade-in duration-500">
                  <table className="w-full text-left text-[11px] font-bold">
                    <thead className="bg-slate-50 border-b border-slate-100 text-slate-400 font-black uppercase tracking-widest">
                      <tr>
                        <th className="p-6 text-slate-900">Periodo</th>
                        <th className="p-6">Venta Real ($)</th>
                        <th className="p-6 text-orange-600">Pronóstico ($)</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-50">
                      {analysis.chartData.map((d, i) => (
                        <tr key={i} className="hover:bg-slate-50 transition-colors">
                          <td className="p-6 text-slate-900 uppercase">{d.fecha}</td>
                          <td className="p-6 text-blue-600">{d.ventas ? `$${d.ventas.toLocaleString()}` : '—'}</td>
                          <td className="p-6 text-orange-600 font-black">{d.proyectado ? `$${d.proyectado.toLocaleString()}` : '—'}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}

              {activeSegment === 3 && (
                <div className="grid grid-cols-3 gap-8 animate-in zoom-in-95 duration-500">
                   <div className="bg-white p-10 rounded-[40px] border border-slate-200 shadow-sm space-y-8">
                      <h4 className="text-[10px] font-black uppercase text-slate-400 tracking-widest">Top Vendedores</h4>
                      <div className="space-y-4">
                         {analysis.topVendedores.map((v, i) => (
                           <div key={i} className="flex flex-col gap-1">
                             <div className="flex justify-between text-[11px] font-black"><span>{v.name}</span><span className="text-blue-600">${v.value.toLocaleString()}</span></div>
                             <div className="w-full bg-slate-100 h-1.5 rounded-full overflow-hidden">
                               <div className="bg-blue-600 h-full" style={{width: `${(v.value / analysis.metrics.totalHist) * 100}%`}}></div>
                             </div>
                           </div>
                         ))}
                      </div>
                   </div>
                   <div className="bg-white p-10 rounded-[40px] border border-slate-200 shadow-sm flex flex-col items-center">
                      <h4 className="text-[10px] font-black uppercase text-slate-400 w-full mb-8 tracking-widest">Mezcla de Pago</h4>
                      <div className="w-full h-64">
                        <ResponsiveContainer width="100%" height="100%">
                          <PieChart>
                            <Pie data={analysis.topMetodos} innerRadius={60} outerRadius={85} paddingAngle={8} dataKey="value">
                              {analysis.topMetodos.map((_, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} cornerRadius={10} />)}
                            </Pie>
                            <Tooltip />
                          </PieChart>
                        </ResponsiveContainer>
                      </div>
                   </div>
                   <div className="bg-white p-10 rounded-[40px] border border-slate-200 shadow-sm space-y-8">
                      <h4 className="text-[10px] font-black uppercase text-slate-400 tracking-widest">Top Clientes</h4>
                      <div className="space-y-3">
                         {analysis.topClientes.map((c, i) => (
                           <div key={i} className="flex justify-between items-center p-4 bg-slate-50 rounded-2xl border border-slate-100 hover:border-blue-200 transition-all">
                             <span className="text-[11px] font-black">{c.name}</span>
                             <span className="text-[10px] font-black text-blue-600 bg-blue-50 px-3 py-1 rounded-full">${c.value.toLocaleString()}</span>
                           </div>
                         ))}
                      </div>
                   </div>
                </div>
              )}
            </>
          )}
        </div>
      </main>
    </div>
  );
};

export default App;