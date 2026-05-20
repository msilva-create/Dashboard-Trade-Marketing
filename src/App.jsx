import { useState, useRef, useEffect } from 'react'
import { LayoutDashboard, TrendingUp, DollarSign, ListTodo, PlusCircle, Trash2, Edit2, X, Download, BarChart2, ShoppingCart, BookOpen, Check, MessageCircle, Bot, Send, Search } from 'lucide-react'
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, Legend } from 'recharts'
import * as XLSX from 'xlsx'

const STORAGE_KEY = 'tracker_v3'
const MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
const TIPOS_PLAN = ['Prolub respalda','Sell out Prolub respalda','Apoyo directo','Activación','Otro']
const CONCEPTOS = ['Apoyo a la nomina','CVC','Producto/promocion','Evento','Material POP','Digital','Transporte','Otro']
const ESTADOS_PEND = ['Pendiente','En curso','Listo','Cancelado']
const PRIORIDADES = ['Alta','Media','Baja']
const QUARTERS = ['Q1','Q2','Q3','Q4']
const ESTADOS_PLAN = ['Activo','En negociación','Cerrado','Cancelado']
const COLORES = ['#6c63ff','#3dd68c','#ff9f43','#ff5f5f','#06b6d4','#a78bfa','#f59e0b','#10b981','#ec4899','#14b8a6']

function load() {
  try { const d = localStorage.getItem(STORAGE_KEY); if (d) return JSON.parse(d) } catch {}
  return {
    inversiones: [
      { id:1, fecha:'2026-01-01', anio:2026, mes:'Enero', distribuidor:'CVS- SERVITECAS S.A.S', tipoPlan:'Prolub respalda', concepto:'Apoyo a la nomina', inversion:1109900, galonesPlan:'', notas:'' },
      { id:2, fecha:'2026-01-01', anio:2026, mes:'Enero', distribuidor:'CVS- SERVITECAS S.A.S', tipoPlan:'Apoyo directo', concepto:'CVC', inversion:1004494, galonesPlan:'', notas:'' },
      { id:3, fecha:'2026-01-01', anio:2026, mes:'Enero', distribuidor:'GRUPO MOTOR S.A.S.', tipoPlan:'Prolub respalda', concepto:'Apoyo a la nomina', inversion:1109900, galonesPlan:'', notas:'' },
      { id:4, fecha:'2026-01-01', anio:2026, mes:'Enero', distribuidor:'LUBRICAFE S.A.S.', tipoPlan:'Prolub respalda', concepto:'Apoyo a la nomina', inversion:1664850, galonesPlan:'', notas:'' },
      { id:5, fecha:'2026-01-01', anio:2026, mes:'Enero', distribuidor:'LUBRICAFE S.A.S.', tipoPlan:'Sell out Prolub respalda', concepto:'Producto/promocion', inversion:1364213, galonesPlan:330, notas:'' },
      { id:6, fecha:'2026-01-01', anio:2026, mes:'Enero', distribuidor:'MAQUINAGRO S.A.S', tipoPlan:'Prolub respalda', concepto:'Apoyo a la nomina', inversion:2774750, galonesPlan:'', notas:'' },
      { id:7, fecha:'2026-02-05', anio:2026, mes:'Febrero', distribuidor:'CVS- SERVITECAS S.A.S', tipoPlan:'Prolub respalda', concepto:'Apoyo a la nomina', inversion:950000, galonesPlan:'', notas:'' },
      { id:8, fecha:'2026-02-05', anio:2026, mes:'Febrero', distribuidor:'MAQUINAGRO S.A.S', tipoPlan:'Sell out Prolub respalda', concepto:'Producto/promocion', inversion:1200000, galonesPlan:200, notas:'' },
    ],
    presupuestos: [
      { id:1, anio:2026, mes:'Enero', monto:12000000 },
      { id:2, anio:2026, mes:'Febrero', monto:11500000 },
      { id:3, anio:2026, mes:'Marzo', monto:13000000 },
    ],
    ventas: [
      { id:1, anio:2026, mes:'Enero', distribuidor:'CVS- SERVITECAS S.A.S', galones:450, ventaNeta:8500000 },
      { id:2, anio:2026, mes:'Enero', distribuidor:'GRUPO MOTOR S.A.S.', galones:320, ventaNeta:6200000 },
      { id:3, anio:2026, mes:'Enero', distribuidor:'LUBRICAFE S.A.S.', galones:610, ventaNeta:11800000 },
      { id:4, anio:2026, mes:'Enero', distribuidor:'MAQUINAGRO S.A.S', galones:780, ventaNeta:15200000 },
      { id:5, anio:2026, mes:'Febrero', distribuidor:'CVS- SERVITECAS S.A.S', galones:480, ventaNeta:9100000 },
      { id:6, anio:2026, mes:'Febrero', distribuidor:'MAQUINAGRO S.A.S', galones:820, ventaNeta:16000000 },
    ],
    planes: [
      { id:1, distribuidor:'CVS- SERVITECAS S.A.S', anio:2026, quarter:'Q1', estado:'Activo', tiposPlan:['Prolub respalda'], metaGalones:1500, metaVenta:28000000, condiciones:'Apoyo nomina mensual $1.1M condicionado a compra mínima 400 galones/mes', acuerdos:'Descuento adicional 2% por volumen trimestral', notas:'Cliente prioritario', historial:[] },
      { id:2, distribuidor:'LUBRICAFE S.A.S.', anio:2026, quarter:'Q1', estado:'Activo', tiposPlan:['Prolub respalda','Sell out Prolub respalda'], metaGalones:1800, metaVenta:35000000, condiciones:'Nomina + sell out en producto. Meta 600 galones/mes', acuerdos:'Bonificación en producto al 110% de meta', notas:'Potencial para Q2', historial:[] },
      { id:3, distribuidor:'MAQUINAGRO S.A.S', anio:2026, quarter:'Q2', estado:'En negociación', tiposPlan:['Prolub respalda'], metaGalones:2400, metaVenta:48000000, condiciones:'En revisión — propuesta enviada el 15 de abril', acuerdos:'Pendiente aprobación gerencia', notas:'Solicitan incremento de apoyo vs Q1', historial:[] },
    ],
    gastosPresupuesto: [],
    pendientes: [
      { id:1, distribuidor:'CVS- SERVITECAS S.A.S', tarea:'Revisar cumplimiento meta Q1', categoria:'Seguimiento', fechaLimite:'2026-03-31', prioridad:'Alta', estado:'Pendiente', responsable:'', notas:'' },
      { id:2, distribuidor:'MAQUINAGRO S.A.S', tarea:'Cerrar negociación plan Q2', categoria:'Negociación', fechaLimite:'2026-04-15', prioridad:'Alta', estado:'En curso', responsable:'', notas:'' },
    ],
  }
}
function save(d) { try { localStorage.setItem(STORAGE_KEY, JSON.stringify(d)) } catch {} }

const cop = n => new Intl.NumberFormat('es-CO',{style:'currency',currency:'COP',minimumFractionDigits:0}).format(n||0)
const num = n => new Intl.NumberFormat('es-CO').format(n||0)

const S = {
  card: { background:'var(--bg2)', border:'1px solid var(--border)', borderRadius:'var(--radius)', overflow:'hidden' },
  th: { padding:'10px 14px', textAlign:'left', fontSize:11, fontWeight:600, color:'var(--text3)', textTransform:'uppercase', letterSpacing:'0.05em', background:'var(--bg3)', whiteSpace:'nowrap' },
  td: { padding:'10px 14px', fontSize:13, borderTop:'1px solid var(--border)' },
  btn: (bg,color) => ({ background:bg, color, padding:'7px 14px', borderRadius:8, fontWeight:500, fontSize:13, display:'flex', alignItems:'center', gap:6, cursor:'pointer', border:'none', fontFamily:'var(--font)' }),
}

function Badge({ label }) {
  const map = {
    Alta:{bg:'var(--red-soft)',c:'var(--red)'}, Media:{bg:'var(--yellow-soft)',c:'var(--yellow)'}, Baja:{bg:'var(--green-soft)',c:'var(--green)'},
    Activo:{bg:'var(--green-soft)',c:'var(--green)'}, 'En negociación':{bg:'var(--yellow-soft)',c:'var(--yellow)'},
    Cerrado:{bg:'rgba(90,90,114,0.2)',c:'var(--text3)'}, Cancelado:{bg:'var(--red-soft)',c:'var(--red)'},
    'En curso':{bg:'var(--accent-soft)',c:'var(--accent2)'}, Pendiente:{bg:'var(--yellow-soft)',c:'var(--yellow)'},
    Listo:{bg:'var(--green-soft)',c:'var(--green)'},
    Aprobado:{bg:'var(--green-soft)',c:'var(--green)'},
    Pagado:{bg:'var(--accent-soft)',c:'var(--accent2)'},
    Rechazado:{bg:'var(--red-soft)',c:'var(--red)'},
  }
  const s = map[label]||{bg:'var(--bg4)',c:'var(--text2)'}
  return <span style={{background:s.bg,color:s.c,fontSize:11,fontWeight:500,padding:'3px 8px',borderRadius:6,whiteSpace:'nowrap'}}>{label}</span>
}

function KpiCard({ icon:Icon, label, value, sub, accent }) {
  return (
    <div style={{background:'var(--bg2)',border:'1px solid var(--border)',borderRadius:'var(--radius)',padding:'18px 22px',display:'flex',flexDirection:'column',gap:6}}>
      <div style={{display:'flex',alignItems:'center',gap:7,color:'var(--text2)',fontSize:11,fontWeight:500,textTransform:'uppercase',letterSpacing:'0.06em'}}><Icon size={13}/>{label}</div>
      <div style={{fontSize:22,fontWeight:600,color:accent||'var(--text)',fontFamily:'var(--mono)',letterSpacing:'-0.02em'}}>{value}</div>
      {sub&&<div style={{fontSize:11,color:'var(--text3)'}}>{sub}</div>}
    </div>
  )
}

function Modal({ title, onClose, children, wide }) {
  return (
    <div style={{position:'fixed',inset:0,background:'rgba(0,0,0,0.75)',backdropFilter:'blur(4px)',zIndex:100,display:'flex',alignItems:'center',justifyContent:'center',padding:20}}
      onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div style={{background:'var(--bg2)',border:'1px solid var(--border2)',borderRadius:16,width:'100%',maxWidth:wide?780:540,maxHeight:'92vh',overflowY:'auto'}}>
        <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',padding:'18px 24px',borderBottom:'1px solid var(--border)',position:'sticky',top:0,background:'var(--bg2)',zIndex:1}}>
          <h3 style={{fontWeight:600,fontSize:15}}>{title}</h3>
          <button onClick={onClose} style={{background:'var(--bg3)',color:'var(--text2)',border:'none',borderRadius:8,padding:'5px 7px',cursor:'pointer',display:'flex'}}><X size={15}/></button>
        </div>
        <div style={{padding:24}}>{children}</div>
      </div>
    </div>
  )
}

function Field({ label, children, span }) {
  return (
    <div style={{display:'flex',flexDirection:'column',gap:5,gridColumn:span?'1/-1':''}}>
      <label style={{fontSize:11,color:'var(--text2)',fontWeight:500,textTransform:'uppercase',letterSpacing:'0.04em'}}>{label}</label>
      {children}
    </div>
  )
}

function CT({ active, payload, label }) {
  if(!active||!payload?.length) return null
  return <div style={{background:'var(--bg3)',border:'1px solid var(--border2)',borderRadius:10,padding:'10px 14px',fontSize:12}}>
    <p style={{fontWeight:600,marginBottom:4}}>{label}</p>
    {payload.map((p,i)=><p key={i} style={{color:p.color}}>{p.name}: {cop(p.value)}</p>)}
  </div>
}

// ═══════════════════════════════════════════════════════
// DASHBOARD
// ═══════════════════════════════════════════════════════
function Dashboard({ data }) {
  const [filtroMes, setFiltroMes] = useState('')
  const [filtroAnio, setFiltroAnio] = useState('2026')
  const [filtroDist, setFiltroDist] = useState('')

  const anios = [...new Set([...data.inversiones.map(i=>i.anio),...data.ventas.map(v=>v.anio)])].sort()
  const todosDistribuidores = [...new Set([...data.inversiones.map(i=>i.distribuidor),...data.ventas.map(v=>v.distribuidor)])].sort()

  const invF = data.inversiones.filter(i=>
    (!filtroMes||i.mes===filtroMes)&&
    (!filtroAnio||i.anio===Number(filtroAnio))&&
    (!filtroDist||i.distribuidor===filtroDist)
  )
  const ventF = data.ventas.filter(v=>
    (!filtroMes||v.mes===filtroMes)&&
    (!filtroAnio||v.anio===Number(filtroAnio))&&
    (!filtroDist||v.distribuidor===filtroDist)
  )
  const presF = data.presupuestos.filter(p=>
    (!filtroMes||p.mes===filtroMes)&&
    (!filtroAnio||p.anio===Number(filtroAnio))
  )

  const totalInv = invF.reduce((s,i)=>s+(i.inversion||0),0)
  const totalVenta = ventF.reduce((s,v)=>s+(v.ventaNeta||0),0)
  const totalGalones = ventF.reduce((s,v)=>s+(v.galones||0),0)
  const totalPres = filtroDist ? 0 : presF.reduce((s,p)=>s+(p.monto||0),0)
  const roiPct = totalVenta>0 ? (totalInv/totalVenta)*100 : 0

  // Por distribuidor
  const distribuidores = filtroDist ? [filtroDist] : [...new Set([...invF.map(i=>i.distribuidor),...ventF.map(v=>v.distribuidor)])]
  const porDist = distribuidores.map(d => {
    const inv = invF.filter(i=>i.distribuidor===d).reduce((s,i)=>s+(i.inversion||0),0)
    const venta = ventF.filter(v=>v.distribuidor===d).reduce((s,v)=>s+(v.ventaNeta||0),0)
    const galones = ventF.filter(v=>v.distribuidor===d).reduce((s,v)=>s+(v.galones||0),0)
    const peso = venta>0 ? (inv/venta)*100 : 0
    return { name:d, inv, venta, galones, peso }
  }).sort((a,b)=>b.inv-a.inv)

  // Por mes (respeta filtro dist pero no filtro mes para mostrar tendencia)
  const porMes = MESES.map(m=>({
    name:m.slice(0,3),
    invertido: data.inversiones.filter(i=>i.mes===m&&(!filtroAnio||i.anio===Number(filtroAnio))&&(!filtroDist||i.distribuidor===filtroDist)).reduce((s,i)=>s+(i.inversion||0),0),
    venta: data.ventas.filter(v=>v.mes===m&&(!filtroAnio||v.anio===Number(filtroAnio))&&(!filtroDist||v.distribuidor===filtroDist)).reduce((s,v)=>s+(v.ventaNeta||0),0),
    presupuesto: filtroDist ? 0 : data.presupuestos.filter(p=>p.mes===m&&(!filtroAnio||p.anio===Number(filtroAnio))).reduce((s,p)=>s+(p.monto||0),0),
  })).filter(m=>m.invertido>0||m.venta>0||m.presupuesto>0)

  // Por concepto
  const porConcepto = CONCEPTOS.map(c=>({name:c,value:invF.filter(i=>i.concepto===c).reduce((s,i)=>s+(i.inversion||0),0)})).filter(c=>c.value>0)

  // Ventas por mes para el dist seleccionado
  const ventasPorMes = filtroDist ? MESES.map(m=>({
    name:m.slice(0,3),
    galones: data.ventas.filter(v=>v.mes===m&&v.distribuidor===filtroDist&&(!filtroAnio||v.anio===Number(filtroAnio))).reduce((s,v)=>s+(v.galones||0),0),
    venta: data.ventas.filter(v=>v.mes===m&&v.distribuidor===filtroDist&&(!filtroAnio||v.anio===Number(filtroAnio))).reduce((s,v)=>s+(v.ventaNeta||0),0),
  })).filter(m=>m.galones>0||m.venta>0) : []

  return (
    <div style={{display:'flex',flexDirection:'column',gap:20}}>
      {/* Filtros */}
      <div style={{display:'flex',gap:10,alignItems:'center',flexWrap:'wrap'}}>
        <select value={filtroAnio} onChange={e=>setFiltroAnio(e.target.value)} style={{width:110}}>
          <option value="">Todos los años</option>
          {anios.map(a=><option key={a}>{a}</option>)}
        </select>
        <select value={filtroMes} onChange={e=>setFiltroMes(e.target.value)} style={{width:150}}>
          <option value="">Todos los meses</option>
          {MESES.map(m=><option key={m}>{m}</option>)}
        </select>
        <select value={filtroDist} onChange={e=>setFiltroDist(e.target.value)} style={{width:220}}>
          <option value="">Todos los distribuidores</option>
          {todosDistribuidores.map(d=><option key={d}>{d}</option>)}
        </select>
        {(filtroMes||filtroAnio||filtroDist)&&(
          <button onClick={()=>{setFiltroMes('');setFiltroAnio('2026');setFiltroDist('')}} style={{...S.btn('var(--bg3)','var(--text2)'),fontSize:12,padding:'5px 12px'}}>
            Limpiar filtros
          </button>
        )}
        {filtroDist&&<span style={{fontSize:12,background:'var(--accent-soft)',color:'var(--accent2)',padding:'4px 12px',borderRadius:20,border:'1px solid rgba(108,99,255,0.2)'}}>📍 {filtroDist}</span>}
      </div>

      {/* KPIs */}
      <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(175px,1fr))',gap:13}}>
        <KpiCard icon={TrendingUp} label="Total invertido" value={cop(totalInv)} sub={`${invF.length} registros`} accent="var(--accent2)"/>
        {!filtroDist&&<KpiCard icon={DollarSign} label="Presupuesto" value={cop(totalPres)} sub={totalPres>0?`${((totalInv/totalPres)*100).toFixed(1)}% ejecutado`:'Sin presupuesto'}/>}
        <KpiCard icon={ShoppingCart} label="Venta neta" value={cop(totalVenta)} sub={`${num(totalGalones)} galones`} accent="var(--green)"/>
        <KpiCard icon={BarChart2} label="Inversión / Venta" value={`${roiPct.toFixed(1)}%`} sub="Peso inversión sobre venta" accent={roiPct>15?'var(--red)':roiPct>10?'var(--yellow)':'var(--green)'}/>
        {filtroDist&&<KpiCard icon={DollarSign} label="Precio prom/galón" value={totalGalones>0?cop(totalVenta/totalGalones):'—'} accent="var(--orange)"/>}
      </div>

      {/* Gráficas */}
      <div style={{display:'grid',gridTemplateColumns: filtroDist?'1fr 1fr':'1.3fr 1fr',gap:16}}>
        <div style={{...S.card,padding:20}}>
          <h4 style={{fontSize:11,fontWeight:600,color:'var(--text2)',textTransform:'uppercase',letterSpacing:'0.05em',marginBottom:14}}>
            {filtroDist ? `Inversión vs Venta — ${filtroDist.split(' ')[0]}` : 'Inversión vs Venta Neta por mes'}
          </h4>
          <ResponsiveContainer width="100%" height={200}>
            <BarChart data={porMes} barGap={3}>
              <XAxis dataKey="name" tick={{fill:'var(--text3)',fontSize:11}} axisLine={false} tickLine={false}/>
              <YAxis tickFormatter={v=>`${(v/1000000).toFixed(0)}M`} tick={{fill:'var(--text3)',fontSize:10}} axisLine={false} tickLine={false}/>
              <Tooltip content={<CT/>}/>
              <Bar dataKey="venta" name="Venta neta" fill="var(--green)" radius={[4,4,0,0]} opacity={0.7}/>
              <Bar dataKey="invertido" name="Inversión" fill="var(--accent)" radius={[4,4,0,0]}/>
              {!filtroDist&&<Bar dataKey="presupuesto" name="Presupuesto" fill="var(--bg4)" radius={[4,4,0,0]}/>}
            </BarChart>
          </ResponsiveContainer>
        </div>

        {filtroDist ? (
          <div style={{...S.card,padding:20}}>
            <h4 style={{fontSize:11,fontWeight:600,color:'var(--text2)',textTransform:'uppercase',letterSpacing:'0.05em',marginBottom:14}}>Galones vendidos por mes</h4>
            <ResponsiveContainer width="100%" height={200}>
              <BarChart data={ventasPorMes}>
                <XAxis dataKey="name" tick={{fill:'var(--text3)',fontSize:11}} axisLine={false} tickLine={false}/>
                <YAxis tick={{fill:'var(--text3)',fontSize:10}} axisLine={false} tickLine={false}/>
                <Tooltip contentStyle={{background:'var(--bg3)',border:'1px solid var(--border2)',borderRadius:10,fontSize:12}}/>
                <Bar dataKey="galones" name="Galones" fill="var(--orange)" radius={[4,4,0,0]}/>
              </BarChart>
            </ResponsiveContainer>
          </div>
        ) : (
          <div style={{...S.card,padding:20}}>
            <h4 style={{fontSize:11,fontWeight:600,color:'var(--text2)',textTransform:'uppercase',letterSpacing:'0.05em',marginBottom:14}}>Inversión por concepto</h4>
            <ResponsiveContainer width="100%" height={200}>
              <PieChart>
                <Pie data={porConcepto} cx="50%" cy="50%" innerRadius={45} outerRadius={75} dataKey="value" nameKey="name" paddingAngle={3}>
                  {porConcepto.map((_,i)=><Cell key={i} fill={COLORES[i%COLORES.length]}/>)}
                </Pie>
                <Tooltip formatter={v=>cop(v)} contentStyle={{background:'var(--bg3)',border:'1px solid var(--border2)',borderRadius:10,fontSize:11}}/>
                <Legend iconSize={7} iconType="circle" wrapperStyle={{fontSize:10,color:'var(--text2)'}}/>
              </PieChart>
            </ResponsiveContainer>
          </div>
        )}
      </div>

      {/* Tabla */}
      <div style={S.card}>
        <div style={{padding:'14px 18px',borderBottom:'1px solid var(--border)',display:'flex',alignItems:'center',justifyContent:'space-between'}}>
          <h4 style={{fontSize:11,fontWeight:600,color:'var(--text2)',textTransform:'uppercase',letterSpacing:'0.05em'}}>
            {filtroDist ? `Detalle — ${filtroDist}` : 'Inversión vs Venta por distribuidor'}
          </h4>
          <span style={{fontSize:11,color:'var(--text3)'}}>% = Inversión / Venta Neta</span>
        </div>
        <table style={{width:'100%',borderCollapse:'collapse'}}>
          <thead><tr>
            {['Distribuidor','Inversión','Venta Neta','Galones','% Inv/Venta','Participación'].map(h=><th key={h} style={S.th}>{h}</th>)}
          </tr></thead>
          <tbody>
            {porDist.length===0&&<tr><td colSpan={6} style={{...S.td,textAlign:'center',color:'var(--text3)',padding:32}}>Sin datos en el período seleccionado</td></tr>}
            {porDist.map((d,i)=>(
              <tr key={i} style={{cursor:'pointer'}} onClick={()=>setFiltroDist(filtroDist===d.name?'':d.name)}>
                <td style={{...S.td,fontWeight:500,color:filtroDist===d.name?'var(--accent2)':'var(--text)'}}>{d.name}</td>
                <td style={{...S.td,fontFamily:'var(--mono)'}}>{cop(d.inv)}</td>
                <td style={{...S.td,fontFamily:'var(--mono)',color:'var(--green)'}}>{cop(d.venta)}</td>
                <td style={{...S.td,fontFamily:'var(--mono)',color:'var(--text2)'}}>{num(d.galones)}</td>
                <td style={{...S.td}}>
                  <span style={{fontFamily:'var(--mono)',fontWeight:600,color:d.peso>15?'var(--red)':d.peso>10?'var(--yellow)':'var(--green)'}}>{d.peso.toFixed(1)}%</span>
                </td>
                <td style={{...S.td,minWidth:140}}>
                  <div style={{display:'flex',alignItems:'center',gap:8}}>
                    <div style={{flex:1,height:5,background:'var(--bg4)',borderRadius:3}}>
                      <div style={{width:`${totalInv>0?Math.min((d.inv/totalInv)*100,100):0}%`,height:'100%',background:COLORES[i%COLORES.length],borderRadius:3}}/>
                    </div>
                    <span style={{fontSize:11,color:'var(--text2)',minWidth:34}}>{totalInv>0?((d.inv/totalInv)*100).toFixed(1):0}%</span>
                  </div>
                </td>
              </tr>
            ))}
            {porDist.length>0&&(
              <tr style={{borderTop:'2px solid var(--border2)'}}>
                <td style={{...S.td,fontWeight:700}}>TOTAL</td>
                <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700,color:'var(--accent2)'}}>{cop(totalInv)}</td>
                <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700,color:'var(--green)'}}>{cop(totalVenta)}</td>
                <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:600}}>{num(totalGalones)}</td>
                <td style={{...S.td,fontWeight:700,color:roiPct>15?'var(--red)':roiPct>10?'var(--yellow)':'var(--green)'}}>{roiPct.toFixed(1)}%</td>
                <td style={S.td}/>
              </tr>
            )}
          </tbody>
        </table>
        {!filtroDist&&<div style={{padding:'8px 18px',fontSize:11,color:'var(--text3)',borderTop:'1px solid var(--border)'}}>💡 Clic en una fila para ver el detalle del distribuidor</div>}
      </div>
    </div>
  )
}

// ═══════════════════════════════════════════════════════
// INVERSIONES — Tabla tipo Excel
// ═══════════════════════════════════════════════════════
function Inversiones({ data, setData }) {
  const [filtroMes, setFiltroMes] = useState('')
  const [filtroAnio, setFiltroAnio] = useState('')
  const [filtroDist, setFiltroDist] = useState('')
  const [busqueda, setBusqueda] = useState('')
  const [seleccionados, setSeleccionados] = useState(new Set())
  const [editCell, setEditCell] = useState(null)
  const [editVal, setEditVal] = useState('')
  const inputRef = useRef(null)

  const COLS = [
    {key:'fecha',        label:'Fecha',        w:100},
    {key:'anio',         label:'Año',          w:55},
    {key:'mes',          label:'Mes',          w:90},
    {key:'distribuidor', label:'Distribuidor', w:200},
    {key:'tipoPlan',     label:'Tipo Plan',    w:150},
    {key:'concepto',     label:'Concepto',     w:130},
    {key:'inversion',    label:'Inversión',    w:130},
    {key:'galonesPlan',  label:'Gal. Plan',    w:75},
    {key:'notas',        label:'Notas',        w:150},
  ]

  const anios = [...new Set(data.inversiones.map(i=>i.anio))].sort()
  const distribuidores = [...new Set(data.inversiones.map(i=>i.distribuidor))].sort()

  const lista = data.inversiones.filter(i=>
    (!filtroMes||i.mes===filtroMes) &&
    (!filtroAnio||i.anio===Number(filtroAnio)) &&
    (!filtroDist||i.distribuidor===filtroDist) &&
    (!busqueda||Object.values(i).some(v=>String(v||'').toLowerCase().includes(busqueda.toLowerCase())))
  ).sort((a,b)=>{
    const mi=MESES.indexOf(a.mes), mj=MESES.indexOf(b.mes)
    return mi!==mj ? mi-mj : String(a.distribuidor||'').localeCompare(String(b.distribuidor||''))
  })

  const totalFiltrado = lista.reduce((s,i)=>s+(Number(i.inversion)||0),0)
  const presMes = filtroMes ? data.presupuestos.filter(p=>p.mes===filtroMes&&(!filtroAnio||p.anio===Number(filtroAnio))).reduce((s,p)=>s+(Number(p.monto)||0),0) : 0

  const toggleSel = id => setSeleccionados(prev=>{const n=new Set(prev);n.has(id)?n.delete(id):n.add(id);return n})
  const toggleTodos = () => setSeleccionados(prev=>prev.size===lista.length&&lista.length>0?new Set():new Set(lista.map(i=>i.id)))
  const eliminarSel = () => {
    if(!seleccionados.size||!confirm('¿Eliminar '+seleccionados.size+' registros?')) return
    const nd={...data,inversiones:data.inversiones.filter(i=>!seleccionados.has(i.id))}
    setData(nd);save(nd);setSeleccionados(new Set())
  }

  const startEdit = (id,field,val) => {
    setEditCell({id,field});setEditVal(String(val||''))
    setTimeout(()=>inputRef.current?.focus(),20)
  }
  const commitEdit = () => {
    if(!editCell) return
    const {id,field}=editCell
    const inversiones=data.inversiones.map(i=>{
      if(i.id!==id) return i
      const val=(field==='inversion'||field==='galonesPlan'||field==='anio')?(Number(editVal)||0):editVal
      return {...i,[field]:val}
    })
    const nd={...data,inversiones};setData(nd);save(nd);setEditCell(null)
  }
  const onKD = e => {
    if(e.key==='Enter'){commitEdit();e.preventDefault()}
    if(e.key==='Escape') setEditCell(null)
    if(e.key==='Tab'){commitEdit();e.preventDefault()}
  }

  const addFila = () => {
    const n={id:Date.now(),fecha:'',anio:2026,mes:'',distribuidor:'',tipoPlan:'',concepto:'',inversion:0,galonesPlan:'',notas:''}
    const nd={...data,inversiones:[...data.inversiones,n]};setData(nd);save(nd)
    setTimeout(()=>startEdit(n.id,'distribuidor',''),60)
  }

  const handlePaste = e => {
    const text=e.clipboardData.getData('text')
    if(!text.includes('\t')&&!text.includes('\n')) return
    e.preventDefault()
    const rows=text.trim().split('\n').map(r=>r.split('\t'))
    const nuevas=rows.filter(r=>r.length>=2&&(r[0]||r[3])).map((r,i)=>({
      id:Date.now()+i, fecha:r[0]||'', anio:Number(r[1])||2026, mes:r[2]||'',
      distribuidor:r[3]||'', tipoPlan:r[4]||'', concepto:r[5]||'',
      inversion:parseN(r[6]||0), galonesPlan:r[7]?parseN(r[7]):'', notas:r[8]||''
    })).filter(r=>r.distribuidor||r.mes)
    if(nuevas.length){
      const nd={...data,inversiones:[...data.inversiones,...nuevas]};setData(nd);save(nd)
      alert('✅ '+nuevas.length+' filas pegadas')
    }
  }

  const celda = (id,field) => ({
    padding:'6px 10px', fontSize:12,
    borderTop:'1px solid var(--border)', borderRight:'1px solid var(--border)',
    cursor:'cell', whiteSpace:'nowrap', overflow:'hidden', maxWidth:220,
    background: editCell?.id===id&&editCell?.field===field ? 'rgba(108,99,255,0.18)' : seleccionados.has(id) ? 'rgba(108,99,255,0.07)' : 'transparent',
    outline: editCell?.id===id&&editCell?.field===field ? '2px solid var(--accent)' : 'none',
    outlineOffset:'-1px',
  })

  return (
    <div style={{display:'flex',flexDirection:'column',gap:14}}>
      <div style={{display:'flex',gap:8,flexWrap:'wrap',alignItems:'center'}}>
        <div style={{position:'relative',display:'flex',alignItems:'center'}}>
          <Search size={14} style={{position:'absolute',left:10,color:'var(--text3)',pointerEvents:'none'}}/>
          <input value={busqueda} onChange={e=>setBusqueda(e.target.value)} placeholder="Buscar..." style={{paddingLeft:32,width:190,fontSize:13}}/>
        </div>
        <select value={filtroAnio} onChange={e=>setFiltroAnio(e.target.value)} style={{width:95}}>
          <option value="">Año</option>{anios.map(a=><option key={a}>{a}</option>)}
        </select>
        <select value={filtroMes} onChange={e=>setFiltroMes(e.target.value)} style={{width:130}}>
          <option value="">Todos los meses</option>{MESES.map(m=><option key={m}>{m}</option>)}
        </select>
        <select value={filtroDist} onChange={e=>setFiltroDist(e.target.value)} style={{width:200}}>
          <option value="">Todos los distribuidores</option>{distribuidores.map(d=><option key={d}>{d}</option>)}
        </select>
        {(filtroMes||filtroAnio||filtroDist||busqueda)&&
          <button onClick={()=>{setFiltroMes('');setFiltroAnio('');setFiltroDist('');setBusqueda('')}} style={{...S.btn('var(--bg3)','var(--text2)'),padding:'5px 10px',fontSize:12}}>✕ Limpiar</button>}
        <div style={{marginLeft:'auto',display:'flex',gap:8}}>
          {seleccionados.size>0&&(
            <button onClick={eliminarSel} style={{...S.btn('var(--red-soft)','var(--red)'),fontSize:12}}>
              <Trash2 size={13}/> Eliminar {seleccionados.size}
            </button>
          )}
          <button onClick={addFila} style={{...S.btn('var(--green-soft)','var(--green)'),fontSize:12,border:'1px solid rgba(61,214,140,0.2)'}}>
            <PlusCircle size={14}/> Añadir fila
          </button>
        </div>
      </div>

      {filtroMes&&(
        <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:10}}>
          <KpiCard icon={TrendingUp} label={'Invertido — '+filtroMes} value={cop(totalFiltrado)} sub={lista.length+' registros'} accent="var(--accent2)"/>
          <KpiCard icon={DollarSign} label="Presupuesto del mes" value={cop(presMes)}/>
          <KpiCard icon={BarChart2} label="% Ejecutado" value={presMes>0?((totalFiltrado/presMes)*100).toFixed(1)+'%':'—'} accent={presMes>0&&totalFiltrado/presMes>=1?'var(--red)':'var(--green)'}/>
          <KpiCard icon={DollarSign} label={presMes-totalFiltrado>=0?'Disponible':'Excedido'} value={cop(Math.abs(presMes-totalFiltrado))} accent={presMes-totalFiltrado>=0?'var(--green)':'var(--red)'}/>
        </div>
      )}

      <div style={{fontSize:11,color:'var(--text3)',display:'flex',gap:20,flexWrap:'wrap'}}>
        <span>💡 Clic en celda para editar · Enter confirma · Tab siguiente</span>
        <span>📋 Ctrl+V para pegar desde Excel (mismo orden de columnas)</span>
        <span>☑️ Checkbox para seleccionar y eliminar en bloque</span>
      </div>

      <div style={{...S.card,overflowX:'auto'}} onPaste={handlePaste}>
        <table style={{borderCollapse:'collapse',tableLayout:'fixed',minWidth:'100%'}}>
          <thead>
            <tr style={{background:'var(--bg3)'}}>
              <th style={{...S.th,width:36,textAlign:'center',padding:'8px 6px'}}>
                <input type="checkbox" checked={seleccionados.size===lista.length&&lista.length>0} onChange={toggleTodos} style={{cursor:'pointer',width:14,height:14,accentColor:'var(--accent)'}}/>
              </th>
              <th style={{...S.th,width:36,padding:'8px 4px',textAlign:'center',fontSize:10}}>#</th>
              {COLS.map(c=><th key={c.key} style={{...S.th,width:c.w}}>{c.label}</th>)}
            </tr>
          </thead>
          <tbody>
            {lista.length===0&&(
              <tr><td colSpan={COLS.length+2} style={{padding:48,textAlign:'center',color:'var(--text3)',fontSize:13}}>
                No hay inversiones. Usa <strong>Añadir fila</strong>, importa tu Excel, o pega con <strong>Ctrl+V</strong>.
              </td></tr>
            )}
            {lista.map((inv,idx)=>(
              <tr key={inv.id} style={{background:seleccionados.has(inv.id)?'rgba(108,99,255,0.06)':'transparent'}}>
                <td style={{padding:'6px',textAlign:'center',borderTop:'1px solid var(--border)',borderRight:'1px solid var(--border)'}}>
                  <input type="checkbox" checked={seleccionados.has(inv.id)} onChange={()=>toggleSel(inv.id)} style={{cursor:'pointer',width:14,height:14,accentColor:'var(--accent)'}}/>
                </td>
                <td style={{padding:'6px 4px',textAlign:'center',fontSize:10,color:'var(--text3)',borderTop:'1px solid var(--border)',borderRight:'1px solid var(--border)'}}>{idx+1}</td>
                {COLS.map(col=>(
                  <td key={col.key} style={celda(inv.id,col.key)} onClick={()=>startEdit(inv.id,col.key,inv[col.key])}>
                    {editCell?.id===inv.id&&editCell?.field===col.key ? (
                      col.key==='mes'?(
                        <select value={editVal} onChange={e=>setEditVal(e.target.value)} onBlur={commitEdit} onKeyDown={onKD} ref={inputRef}
                          style={{background:'transparent',color:'var(--text)',border:'none',outline:'none',fontSize:12,width:'100%',fontFamily:'var(--font)'}}>
                          <option value="">—</option>{MESES.map(m=><option key={m}>{m}</option>)}
                        </select>
                      ):col.key==='concepto'?(
                        <select value={editVal} onChange={e=>setEditVal(e.target.value)} onBlur={commitEdit} onKeyDown={onKD} ref={inputRef}
                          style={{background:'transparent',color:'var(--text)',border:'none',outline:'none',fontSize:12,width:'100%',fontFamily:'var(--font)'}}>
                          {CONCEPTOS.map(c=><option key={c}>{c}</option>)}
                        </select>
                      ):(
                        <input ref={inputRef} value={editVal} onChange={e=>setEditVal(e.target.value)} onBlur={commitEdit} onKeyDown={onKD}
                          type={col.key==='inversion'||col.key==='galonesPlan'||col.key==='anio'?'number':'text'}
                          style={{background:'transparent',color:'var(--text)',border:'none',outline:'none',fontSize:12,width:'100%',fontFamily:col.key==='inversion'?'var(--mono)':'var(--font)'}}/>
                      )
                    ):(
                      <span style={{color:col.key==='distribuidor'?'var(--text)':'var(--text2)',fontFamily:col.key==='inversion'?'var(--mono)':'inherit'}}>
                        {col.key==='inversion'?cop(Number(inv[col.key])||0):(inv[col.key]||'')}
                      </span>
                    )}
                  </td>
                ))}
              </tr>
            ))}
            {lista.length>0&&(
              <tr style={{background:'var(--bg3)',borderTop:'2px solid var(--border2)'}}>
                <td colSpan={2} style={{padding:'8px 10px'}}/>
                <td colSpan={5} style={{padding:'8px 12px',fontWeight:700,fontSize:12,color:'var(--text2)'}}>
                  TOTAL {filtroMes&&'— '+filtroMes} · {lista.length} registros {seleccionados.size>0&&'· '+seleccionados.size+' seleccionados'}
                </td>
                <td style={{padding:'8px 12px',fontFamily:'var(--mono)',fontWeight:700,fontSize:13,color:'var(--accent2)'}}>{cop(totalFiltrado)}</td>
                <td colSpan={2} style={{padding:'8px 10px'}}/>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  )
}

// ═══════════════════════════════════════════════════════
// VENTAS
// ═══════════════════════════════════════════════════════
function Ventas({ data, setData }) {
  const [modal, setModal] = useState(false)
  const [filtroMes, setFiltroMes] = useState('')
  const [filtroAnio, setFiltroAnio] = useState('')
  const [editId, setEditId] = useState(null)
  const blank = { anio:2026, mes:'', distribuidor:'', galones:'', ventaNeta:'', notas:'' }
  const [form, setForm] = useState(blank)

  const anios = [...new Set([...data.ventas.map(v=>v.anio),...data.inversiones.map(i=>i.anio)])].sort()
  const distribuidores = [...new Set([...data.ventas.map(v=>v.distribuidor),...data.inversiones.map(i=>i.distribuidor)])].sort()

  const lista = data.ventas.filter(v=>(!filtroMes||v.mes===filtroMes)&&(!filtroAnio||v.anio===Number(filtroAnio))).sort((a,b)=>b.id-a.id)

  const submit = () => {
    if(!form.distribuidor||!form.mes) return
    const entry = {...form, id:editId||Date.now(), galones:Number(form.galones)||0, ventaNeta:Number(form.ventaNeta)||0, anio:Number(form.anio)}
    const ventas = editId ? data.ventas.map(v=>v.id===editId?entry:v) : [...data.ventas,entry]
    const nd = {...data,ventas}; setData(nd); save(nd); setModal(false); setEditId(null); setForm(blank)
  }
  const del = id => { const nd={...data,ventas:data.ventas.filter(v=>v.id!==id)}; setData(nd); save(nd) }
  const edit = v => { setForm({...v,galones:v.galones.toString(),ventaNeta:v.ventaNeta.toString()}); setEditId(v.id); setModal(true) }
  const getInv = (dist,mes,anio) => data.inversiones.filter(i=>i.distribuidor===dist&&i.mes===mes&&i.anio===anio).reduce((s,i)=>s+(i.inversion||0),0)

  const totalGalones = lista.reduce((s,v)=>s+(v.galones||0),0)
  const totalVenta = lista.reduce((s,v)=>s+(v.ventaNeta||0),0)

  return (
    <div style={{display:'flex',flexDirection:'column',gap:18}}>
      <div style={{display:'flex',gap:10,flexWrap:'wrap',alignItems:'center'}}>
        <select value={filtroAnio} onChange={e=>setFiltroAnio(e.target.value)} style={{width:110}}>
          <option value="">Año</option>
          {anios.map(a=><option key={a}>{a}</option>)}
        </select>
        <select value={filtroMes} onChange={e=>setFiltroMes(e.target.value)} style={{width:160}}>
          <option value="">Todos los meses</option>
          {MESES.map(m=><option key={m}>{m}</option>)}
        </select>
        <button onClick={()=>{setEditId(null);setForm(blank);setModal(true)}} style={{...S.btn('var(--accent)','#fff'),marginLeft:'auto'}}>
          <PlusCircle size={15}/> Registrar venta
        </button>
      </div>
      <div style={{display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:12}}>
        <KpiCard icon={ShoppingCart} label="Venta neta total" value={cop(totalVenta)} sub={`${lista.length} registros`} accent="var(--green)"/>
        <KpiCard icon={BarChart2} label="Total galones" value={num(totalGalones)} accent="var(--accent2)"/>
        <KpiCard icon={DollarSign} label="Precio prom./galón" value={totalGalones>0?cop(totalVenta/totalGalones):'—'} accent="var(--orange)"/>
      </div>
      <div style={S.card}>
        <table style={{width:'100%',borderCollapse:'collapse'}}>
          <thead><tr>{['Año','Mes','Distribuidor','Galones','Venta Neta','Precio/Galón','Inversión','% Inv/Venta',''].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
          <tbody>
            {lista.length===0&&<tr><td colSpan={9} style={{...S.td,textAlign:'center',color:'var(--text3)',padding:40}}>No hay ventas registradas</td></tr>}
            {lista.map(v=>{
              const inv=getInv(v.distribuidor,v.mes,v.anio)
              const peso=v.ventaNeta>0?(inv/v.ventaNeta)*100:0
              return (
                <tr key={v.id}>
                  <td style={{...S.td,color:'var(--text2)'}}>{v.anio}</td>
                  <td style={{...S.td,color:'var(--text2)'}}>{v.mes}</td>
                  <td style={{...S.td,fontWeight:500}}>{v.distribuidor}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)'}}>{num(v.galones)}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',color:'var(--green)'}}>{cop(v.ventaNeta)}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',color:'var(--text2)'}}>{v.galones>0?cop(v.ventaNeta/v.galones):'—'}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)'}}>{cop(inv)}</td>
                  <td style={{...S.td}}><span style={{fontFamily:'var(--mono)',fontWeight:600,color:peso>15?'var(--red)':peso>10?'var(--yellow)':'var(--green)'}}>{peso.toFixed(1)}%</span></td>
                  <td style={S.td}>
                    <div style={{display:'flex',gap:6}}>
                      <button onClick={()=>edit(v)} style={{...S.btn('var(--bg3)','var(--text2)'),padding:'4px 8px'}}><Edit2 size={13}/></button>
                      <button onClick={()=>del(v.id)} style={{...S.btn('var(--red-soft)','var(--red)'),padding:'4px 8px'}}><Trash2 size={13}/></button>
                    </div>
                  </td>
                </tr>
              )
            })}
            {lista.length>0&&(
              <tr style={{borderTop:'2px solid var(--border2)'}}>
                <td colSpan={3} style={{...S.td,fontWeight:700,color:'var(--text2)'}}>TOTAL</td>
                <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700}}>{num(totalGalones)}</td>
                <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700,color:'var(--green)'}}>{cop(totalVenta)}</td>
                <td colSpan={4} style={S.td}/>
              </tr>
            )}
          </tbody>
        </table>
      </div>
      {modal&&(
        <Modal title={editId?'Editar venta':'Registrar venta'} onClose={()=>{setModal(false);setEditId(null)}}>
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:14}}>
            <Field label="Año"><input type="number" value={form.anio} onChange={e=>setForm({...form,anio:e.target.value})} placeholder="2026"/></Field>
            <Field label="Mes *">
              <select value={form.mes} onChange={e=>setForm({...form,mes:e.target.value})}>
                <option value="">Selecciona...</option>
                {MESES.map(m=><option key={m}>{m}</option>)}
              </select>
            </Field>
            <Field label="Distribuidor *" span>
              <input list="dist-venta" value={form.distribuidor} onChange={e=>setForm({...form,distribuidor:e.target.value})} placeholder="Nombre del distribuidor"/>
              <datalist id="dist-venta">{distribuidores.map(d=><option key={d} value={d}/>)}</datalist>
            </Field>
            <Field label="Galones vendidos"><input type="number" value={form.galones} onChange={e=>setForm({...form,galones:e.target.value})} placeholder="0"/></Field>
            <Field label="Venta Neta (COP)"><input type="number" value={form.ventaNeta} onChange={e=>setForm({...form,ventaNeta:e.target.value})} placeholder="0"/></Field>
            <Field label="Notas" span><textarea value={form.notas||''} onChange={e=>setForm({...form,notas:e.target.value})} rows={2} style={{resize:'vertical'}} placeholder="Observaciones..."/></Field>
          </div>
          <div style={{display:'flex',gap:10,justifyContent:'flex-end',marginTop:20}}>
            <button onClick={()=>{setModal(false);setEditId(null)}} style={S.btn('var(--bg3)','var(--text2)')}>Cancelar</button>
            <button onClick={submit} style={S.btn('var(--accent)','#fff')}><Check size={15}/> Guardar</button>
          </div>
        </Modal>
      )}
    </div>
  )
}

// ═══════════════════════════════════════════════════════
// PLANES
// ═══════════════════════════════════════════════════════
function Planes({ data, setData }) {
  const [modal, setModal] = useState(false)
  const [modalHoja, setModalHoja] = useState(null)
  const [filtroQ, setFiltroQ] = useState('')
  const [filtroAnio, setFiltroAnio] = useState('2026')
  const [editId, setEditId] = useState(null)
  const blank = { distribuidor:'', anio:2026, quarter:'Q1', estado:'Activo', tiposPlan:[], metaGalones:'', metaVenta:'', condiciones:'', acuerdos:'', notas:'', historial:[] }
  const [form, setForm] = useState(blank)
  const [nuevaNota, setNuevaNota] = useState('')

  const anios = [...new Set([...data.planes.map(p=>p.anio),...data.inversiones.map(i=>i.anio)])].sort()
  const distribuidores = [...new Set([...data.planes.map(p=>p.distribuidor),...data.inversiones.map(i=>i.distribuidor)])].sort()
  const lista = data.planes.filter(p=>(!filtroQ||p.quarter===filtroQ)&&(!filtroAnio||p.anio===Number(filtroAnio)))

  const submit = () => {
    if(!form.distribuidor||!form.quarter) return
    const entry = {...form, id:editId||Date.now(), metaGalones:Number(form.metaGalones)||0, metaVenta:Number(form.metaVenta)||0, anio:Number(form.anio), historial:form.historial||[]}
    const planes = editId ? data.planes.map(p=>p.id===editId?entry:p) : [...data.planes,entry]
    const nd={...data,planes}; setData(nd); save(nd); setModal(false); setEditId(null); setForm(blank)
  }
  const del = id => { const nd={...data,planes:data.planes.filter(p=>p.id!==id)}; setData(nd); save(nd) }
  const edit = p => { setForm({...p,metaGalones:p.metaGalones?.toString()||'',metaVenta:p.metaVenta?.toString()||''}); setEditId(p.id); setModal(true) }

  const agregarNota = planId => {
    if(!nuevaNota.trim()) return
    const nota = { texto:nuevaNota, fecha:new Date().toLocaleDateString('es-CO') }
    const planes = data.planes.map(p=>p.id===planId?{...p,historial:[...(p.historial||[]),nota]}:p)
    const nd={...data,planes}; setData(nd); save(nd)
    setModalHoja(prev=>({...prev,historial:[...(prev.historial||[]),nota]}))
    setNuevaNota('')
  }

  const qResumen = QUARTERS.map(q=>({
    q, total:data.planes.filter(p=>p.quarter===q&&(!filtroAnio||p.anio===Number(filtroAnio))).length,
    activos:data.planes.filter(p=>p.quarter===q&&p.estado==='Activo'&&(!filtroAnio||p.anio===Number(filtroAnio))).length,
  })).filter(q=>q.total>0)

  return (
    <div style={{display:'flex',flexDirection:'column',gap:18}}>
      <div style={{display:'flex',gap:10,flexWrap:'wrap',alignItems:'center'}}>
        <select value={filtroAnio} onChange={e=>setFiltroAnio(e.target.value)} style={{width:110}}>
          <option value="">Año</option>
          {anios.map(a=><option key={a}>{a}</option>)}
        </select>
        <select value={filtroQ} onChange={e=>setFiltroQ(e.target.value)} style={{width:120}}>
          <option value="">Todos los Q</option>
          {QUARTERS.map(q=><option key={q}>{q}</option>)}
        </select>
        <button onClick={()=>{setEditId(null);setForm(blank);setModal(true)}} style={{...S.btn('var(--accent)','#fff'),marginLeft:'auto'}}>
          <PlusCircle size={15}/> Nuevo plan
        </button>
      </div>

      {qResumen.length>0&&(
        <div style={{display:'flex',gap:12,flexWrap:'wrap'}}>
          {qResumen.map(q=>(
            <div key={q.q} onClick={()=>setFiltroQ(filtroQ===q.q?'':q.q)}
              style={{background:filtroQ===q.q?'var(--accent-soft)':'var(--bg2)',border:`1px solid ${filtroQ===q.q?'var(--accent)':'var(--border)'}`,borderRadius:10,padding:'12px 20px',cursor:'pointer',display:'flex',flexDirection:'column',gap:4,minWidth:110}}>
              <span style={{fontSize:18,fontWeight:700,color:filtroQ===q.q?'var(--accent2)':'var(--text)'}}>{q.q}</span>
              <span style={{fontSize:11,color:'var(--text3)'}}>{q.activos} activo{q.activos!==1?'s':''} / {q.total} total</span>
            </div>
          ))}
        </div>
      )}

      <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fill,minmax(340px,1fr))',gap:14}}>
        {lista.length===0&&<div style={{...S.card,padding:40,textAlign:'center',color:'var(--text3)',gridColumn:'1/-1'}}>No hay planes registrados {filtroQ&&`para ${filtroQ}`}</div>}
        {lista.map(p=>{
          const galonesReal = data.ventas.filter(v=>v.distribuidor===p.distribuidor).reduce((s,v)=>s+(v.galones||0),0)
          const ventaTotal = data.ventas.filter(v=>v.distribuidor===p.distribuidor).reduce((s,v)=>s+(v.ventaNeta||0),0)
          const invTotal = data.inversiones.filter(i=>i.distribuidor===p.distribuidor).reduce((s,i)=>s+(i.inversion||0),0)
          const cumpl = p.metaGalones>0?(galonesReal/p.metaGalones)*100:0
          return (
            <div key={p.id} style={{...S.card,display:'flex',flexDirection:'column'}}>
              <div style={{padding:'14px 18px',borderBottom:'1px solid var(--border)',display:'flex',alignItems:'flex-start',justifyContent:'space-between',gap:8}}>
                <div>
                  <div style={{fontWeight:600,fontSize:14,marginBottom:4}}>{p.distribuidor}</div>
                  <div style={{display:'flex',gap:6,flexWrap:'wrap'}}>
                    <span style={{background:'var(--bg4)',color:'var(--accent2)',fontSize:11,fontWeight:600,padding:'2px 8px',borderRadius:5}}>{p.quarter} · {p.anio}</span>
                    <Badge label={p.estado}/>
                  </div>
                </div>
                <div style={{display:'flex',gap:5}}>
                  <button onClick={()=>setModalHoja(p)} style={{...S.btn('var(--bg3)','var(--text2)'),padding:'4px 8px',fontSize:11}}><BookOpen size={12}/> Ver</button>
                  <button onClick={()=>edit(p)} style={{...S.btn('var(--bg3)','var(--text2)'),padding:'4px 8px'}}><Edit2 size={12}/></button>
                  <button onClick={()=>del(p.id)} style={{...S.btn('var(--red-soft)','var(--red)'),padding:'4px 8px'}}><Trash2 size={12}/></button>
                </div>
              </div>
              <div style={{padding:'14px 18px',display:'flex',flexDirection:'column',gap:10,flex:1}}>
                {p.tiposPlan?.length>0&&<div style={{display:'flex',gap:5,flexWrap:'wrap'}}>{p.tiposPlan.map(t=><span key={t} style={{background:'var(--accent-soft)',color:'var(--accent2)',fontSize:10,padding:'2px 7px',borderRadius:5}}>{t}</span>)}</div>}
                <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:8}}>
                  <div style={{background:'var(--bg3)',borderRadius:8,padding:'8px 12px'}}>
                    <div style={{fontSize:10,color:'var(--text3)',marginBottom:2}}>META GALONES</div>
                    <div style={{fontFamily:'var(--mono)',fontWeight:600,fontSize:14}}>{num(p.metaGalones)}</div>
                    {p.metaGalones>0&&<div style={{fontSize:10,color:'var(--text2)',marginTop:2}}>Real: {num(galonesReal)} ({cumpl.toFixed(0)}%)</div>}
                  </div>
                  <div style={{background:'var(--bg3)',borderRadius:8,padding:'8px 12px'}}>
                    <div style={{fontSize:10,color:'var(--text3)',marginBottom:2}}>META VENTA</div>
                    <div style={{fontFamily:'var(--mono)',fontWeight:600,fontSize:13}}>{cop(p.metaVenta)}</div>
                    {p.metaVenta>0&&<div style={{fontSize:10,color:'var(--text2)',marginTop:2}}>Real: {cop(ventaTotal)}</div>}
                  </div>
                </div>
                {p.metaGalones>0&&(
                  <div>
                    <div style={{display:'flex',justifyContent:'space-between',fontSize:10,color:'var(--text3)',marginBottom:4}}>
                      <span>Cumplimiento galones</span>
                      <span style={{color:cumpl>=100?'var(--green)':cumpl>=70?'var(--yellow)':'var(--red)'}}>{cumpl.toFixed(1)}%</span>
                    </div>
                    <div style={{height:5,background:'var(--bg4)',borderRadius:3}}>
                      <div style={{width:`${Math.min(cumpl,100)}%`,height:'100%',background:cumpl>=100?'var(--green)':cumpl>=70?'var(--yellow)':'var(--red)',borderRadius:3}}/>
                    </div>
                  </div>
                )}
                <div style={{display:'flex',justifyContent:'space-between',fontSize:12,color:'var(--text2)'}}>
                  <span>Inversión acumulada:</span>
                  <span style={{fontFamily:'var(--mono)',color:'var(--accent2)',fontWeight:600}}>{cop(invTotal)}</span>
                </div>
                {p.condiciones&&<p style={{fontSize:11,color:'var(--text3)',borderLeft:'2px solid var(--border2)',paddingLeft:8,margin:0}}>{p.condiciones.slice(0,80)}{p.condiciones.length>80?'…':''}</p>}
                {(p.historial?.length||0)>0&&<div style={{fontSize:11,color:'var(--text3)'}}>📝 {p.historial.length} nota{p.historial.length!==1?'s':''} en historial</div>}
              </div>
            </div>
          )
        })}
      </div>

      {modal&&(
        <Modal title={editId?'Editar plan':'Nuevo plan'} onClose={()=>{setModal(false);setEditId(null)}} wide>
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:14}}>
            <Field label="Distribuidor *" span>
              <input list="dist-plan" value={form.distribuidor} onChange={e=>setForm({...form,distribuidor:e.target.value})} placeholder="Nombre del distribuidor"/>
              <datalist id="dist-plan">{distribuidores.map(d=><option key={d} value={d}/>)}</datalist>
            </Field>
            <Field label="Año"><input type="number" value={form.anio} onChange={e=>setForm({...form,anio:e.target.value})} placeholder="2026"/></Field>
            <Field label="Quarter *">
              <select value={form.quarter} onChange={e=>setForm({...form,quarter:e.target.value})}>
                {QUARTERS.map(q=><option key={q}>{q}</option>)}
              </select>
            </Field>
            <Field label="Estado">
              <select value={form.estado} onChange={e=>setForm({...form,estado:e.target.value})}>
                {ESTADOS_PLAN.map(s=><option key={s}>{s}</option>)}
              </select>
            </Field>
            <Field label="Meta galones"><input type="number" value={form.metaGalones} onChange={e=>setForm({...form,metaGalones:e.target.value})} placeholder="0"/></Field>
            <Field label="Meta venta neta (COP)"><input type="number" value={form.metaVenta} onChange={e=>setForm({...form,metaVenta:e.target.value})} placeholder="0"/></Field>
            <Field label="Tipos de plan" span>
              <div style={{display:'flex',flexWrap:'wrap',gap:8}}>
                {TIPOS_PLAN.map(t=>(
                  <label key={t} style={{display:'flex',alignItems:'center',gap:5,fontSize:13,cursor:'pointer',color:'var(--text2)'}}>
                    <input type="checkbox" checked={form.tiposPlan?.includes(t)} onChange={e=>{
                      const arr=e.target.checked?[...(form.tiposPlan||[]),t]:(form.tiposPlan||[]).filter(x=>x!==t)
                      setForm({...form,tiposPlan:arr})
                    }}/>{t}
                  </label>
                ))}
              </div>
            </Field>
            <Field label="Condiciones del plan" span><textarea value={form.condiciones} onChange={e=>setForm({...form,condiciones:e.target.value})} rows={3} style={{resize:'vertical'}} placeholder="Ej: Apoyo nomina mensual $X condicionado a compra mínima de Y galones/mes..."/></Field>
            <Field label="Acuerdos especiales" span><textarea value={form.acuerdos} onChange={e=>setForm({...form,acuerdos:e.target.value})} rows={2} style={{resize:'vertical'}} placeholder="Ej: Bonificación en producto al 110% de meta..."/></Field>
            <Field label="Notas internas" span><textarea value={form.notas} onChange={e=>setForm({...form,notas:e.target.value})} rows={2} style={{resize:'vertical'}} placeholder="Observaciones, alertas, oportunidades..."/></Field>
          </div>
          <div style={{display:'flex',gap:10,justifyContent:'flex-end',marginTop:20}}>
            <button onClick={()=>{setModal(false);setEditId(null)}} style={S.btn('var(--bg3)','var(--text2)')}>Cancelar</button>
            <button onClick={submit} style={S.btn('var(--accent)','#fff')}><Check size={15}/> Guardar plan</button>
          </div>
        </Modal>
      )}

      {modalHoja&&(
        <Modal title={`Hoja de vida — ${modalHoja.distribuidor}`} onClose={()=>{setModalHoja(null);setNuevaNota('')}} wide>
          <div style={{display:'flex',flexDirection:'column',gap:18}}>
            <div style={{display:'flex',gap:8}}><span style={{background:'var(--bg4)',color:'var(--accent2)',fontSize:12,fontWeight:600,padding:'3px 10px',borderRadius:6}}>{modalHoja.quarter} · {modalHoja.anio}</span><Badge label={modalHoja.estado}/></div>
            {modalHoja.tiposPlan?.length>0&&<div><div style={{fontSize:11,color:'var(--text3)',marginBottom:6,textTransform:'uppercase',letterSpacing:'0.05em'}}>Tipos de plan</div><div style={{display:'flex',gap:6,flexWrap:'wrap'}}>{modalHoja.tiposPlan.map(t=><span key={t} style={{background:'var(--accent-soft)',color:'var(--accent2)',fontSize:12,padding:'3px 10px',borderRadius:6}}>{t}</span>)}</div></div>}
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:12}}>
              <div style={{background:'var(--bg3)',borderRadius:10,padding:'14px 18px'}}><div style={{fontSize:11,color:'var(--text3)',marginBottom:4}}>META GALONES</div><div style={{fontFamily:'var(--mono)',fontSize:20,fontWeight:700}}>{num(modalHoja.metaGalones)}</div></div>
              <div style={{background:'var(--bg3)',borderRadius:10,padding:'14px 18px'}}><div style={{fontSize:11,color:'var(--text3)',marginBottom:4}}>META VENTA NETA</div><div style={{fontFamily:'var(--mono)',fontSize:18,fontWeight:700}}>{cop(modalHoja.metaVenta)}</div></div>
            </div>
            {modalHoja.condiciones&&<div><div style={{fontSize:11,color:'var(--text3)',marginBottom:6,textTransform:'uppercase',letterSpacing:'0.05em'}}>Condiciones del plan</div><div style={{background:'var(--bg3)',borderRadius:10,padding:'12px 16px',fontSize:13,lineHeight:1.6}}>{modalHoja.condiciones}</div></div>}
            {modalHoja.acuerdos&&<div><div style={{fontSize:11,color:'var(--text3)',marginBottom:6,textTransform:'uppercase',letterSpacing:'0.05em'}}>Acuerdos especiales</div><div style={{background:'var(--green-soft)',border:'1px solid rgba(61,214,140,0.2)',borderRadius:10,padding:'12px 16px',fontSize:13,lineHeight:1.6}}>{modalHoja.acuerdos}</div></div>}
            {modalHoja.notas&&<div><div style={{fontSize:11,color:'var(--text3)',marginBottom:6,textTransform:'uppercase',letterSpacing:'0.05em'}}>Notas internas</div><div style={{background:'var(--yellow-soft)',border:'1px solid rgba(255,209,102,0.2)',borderRadius:10,padding:'12px 16px',fontSize:13,lineHeight:1.6}}>{modalHoja.notas}</div></div>}
            <div>
              <div style={{fontSize:11,color:'var(--text3)',marginBottom:8,textTransform:'uppercase',letterSpacing:'0.05em'}}>Historial de seguimiento</div>
              <div style={{display:'flex',flexDirection:'column',gap:8,marginBottom:12}}>
                {(modalHoja.historial||[]).length===0&&<div style={{fontSize:12,color:'var(--text3)',fontStyle:'italic'}}>Sin notas aún</div>}
                {(modalHoja.historial||[]).map((n,i)=>(
                  <div key={i} style={{display:'flex',gap:10,padding:'10px 14px',background:'var(--bg3)',borderRadius:8}}>
                    <span style={{fontSize:11,color:'var(--text3)',whiteSpace:'nowrap',marginTop:1}}>{n.fecha}</span>
                    <span style={{fontSize:13}}>{n.texto}</span>
                  </div>
                ))}
              </div>
              <div style={{display:'flex',gap:10}}>
                <input value={nuevaNota} onChange={e=>setNuevaNota(e.target.value)} onKeyDown={e=>e.key==='Enter'&&agregarNota(modalHoja.id)} placeholder="Agregar nota de seguimiento..." style={{flex:1}}/>
                <button onClick={()=>agregarNota(modalHoja.id)} style={S.btn('var(--accent)','#fff')}>Agregar</button>
              </div>
            </div>
          </div>
        </Modal>
      )}
    </div>
  )
}

// ═══════════════════════════════════════════════════════
// PRESUPUESTO — Hoja tipo Excel + Ejecutado vs Presupuesto
// ═══════════════════════════════════════════════════════
function Presupuesto({ data, setData }) {
  const [filtroMes, setFiltroMes] = useState('')
  const [filtroAnio, setFiltroAnio] = useState('2026')
  const [busqueda, setBusqueda] = useState('')
  const [seleccionados, setSeleccionados] = useState(new Set())
  const [editCell, setEditCell] = useState(null)
  const [editVal, setEditVal] = useState('')
  const [modalPres, setModalPres] = useState(false)
  const [formP, setFormP] = useState({ mes:'', anio:2026, monto:'' })
  const inputRef = useRef(null)

  const COLS_G = [
    {key:'mes',         label:'Mes',                    w:90},
    {key:'gasto',       label:'Gasto (Nom. Producto, Cliente)', w:280},
    {key:'valorFactura',label:'Valor Factura',           w:130},
    {key:'canal',       label:'Canal',                   w:80},
    {key:'observacion', label:'Observación (ATJ, otros)',w:200},
    {key:'estado',      label:'Estado',                  w:100},
    {key:'centroCostos',label:'Centro de Costos',        w:140},
  ]
  const ESTADOS_G = ['Pendiente','Ingresado','Aprobado','Pagado','Rechazado']
  const anios = [...new Set([...(data.gastosPresupuesto||[]).map(g=>g.anio),...data.presupuestos.map(p=>p.anio)])].sort()

  const gastos = (data.gastosPresupuesto||[]).filter(g=>
    (!filtroMes||g.mes===filtroMes) &&
    (!filtroAnio||g.anio===Number(filtroAnio)) &&
    (!busqueda||[g.gasto,g.observacion,g.estado,g.centroCostos,g.canal].some(v=>String(v||'').toLowerCase().includes(busqueda.toLowerCase())))
  ).sort((a,b)=>MESES.indexOf(a.mes)-MESES.indexOf(b.mes)||(String(a.gasto||'').localeCompare(String(b.gasto||''))))

  const totalGastado = gastos.reduce((s,g)=>s+(Number(g.valorFactura)||0),0)
  const presAsignado = data.presupuestos.filter(p=>(!filtroMes||p.mes===filtroMes)&&(!filtroAnio||p.anio===Number(filtroAnio))).reduce((s,p)=>s+(Number(p.monto)||0),0)
  const ejec = presAsignado>0?(totalGastado/presAsignado)*100:0
  const disponible = presAsignado - totalGastado

  // Selección
  const toggleSel = id => setSeleccionados(prev=>{const n=new Set(prev);n.has(id)?n.delete(id):n.add(id);return n})
  const toggleTodos = () => setSeleccionados(prev=>prev.size===gastos.length&&gastos.length>0?new Set():new Set(gastos.map(g=>g.id)))
  const eliminarSel = () => {
    if(!seleccionados.size||!confirm('¿Eliminar '+seleccionados.size+' gastos?')) return
    const nd={...data,gastosPresupuesto:(data.gastosPresupuesto||[]).filter(g=>!seleccionados.has(g.id))}
    setData(nd);save(nd);setSeleccionados(new Set())
  }

  // Edición inline
  const startEdit = (id,field,val) => { setEditCell({id,field});setEditVal(String(val||''));setTimeout(()=>inputRef.current?.focus(),20) }
  const commitEdit = () => {
    if(!editCell) return
    const {id,field}=editCell
    const gastosPresupuesto=(data.gastosPresupuesto||[]).map(g=>{
      if(g.id!==id) return g
      const val=field==='valorFactura'?(Number(editVal)||0):editVal
      return {...g,[field]:val}
    })
    const nd={...data,gastosPresupuesto};setData(nd);save(nd);setEditCell(null)
  }
  const onKD = e => { if(e.key==='Enter'){commitEdit();e.preventDefault()} else if(e.key==='Escape') setEditCell(null); else if(e.key==='Tab'){commitEdit();e.preventDefault()} }

  // Añadir fila
  const addFila = () => {
    const n={id:Date.now(),anio:Number(filtroAnio)||2026,mes:filtroMes||'',gasto:'',valorFactura:0,canal:'',observacion:'',estado:'Pendiente',centroCostos:'',notas:''}
    const nd={...data,gastosPresupuesto:[...(data.gastosPresupuesto||[]),n]};setData(nd);save(nd)
    setTimeout(()=>startEdit(n.id,'gasto',''),60)
  }

  // Pegar desde Excel
  const handlePaste = e => {
    const text=e.clipboardData.getData('text')
    if(!text.includes('\t')&&!text.includes('\n')) return
    e.preventDefault()
    const rows=text.trim().split('\n').map(r=>r.split('\t'))
    const nuevas=rows.filter(r=>r[0]||r[1]).map((r,i)=>({
      id:Date.now()+i, anio:Number(filtroAnio)||2026,
      mes:r[0]||filtroMes||'', gasto:r[1]||'',
      valorFactura:parseN(r[2]||0), canal:r[3]||'',
      observacion:r[4]||'', estado:r[5]||'Pendiente',
      centroCostos:r[6]||'', notas:''
    })).filter(r=>r.gasto||r.mes)
    if(nuevas.length){
      const nd={...data,gastosPresupuesto:[...(data.gastosPresupuesto||[]),...nuevas]};setData(nd);save(nd)
      alert('✅ '+nuevas.length+' filas pegadas')
    }
  }

  const celda = (id,field) => ({
    padding:'6px 10px',fontSize:12,
    borderTop:'1px solid var(--border)',borderRight:'1px solid var(--border)',
    cursor:'cell',whiteSpace:'nowrap',overflow:'hidden',maxWidth:300,
    background:editCell?.id===id&&editCell?.field===field?'rgba(108,99,255,0.18)':seleccionados.has(id)?'rgba(108,99,255,0.07)':'transparent',
    outline:editCell?.id===id&&editCell?.field===field?'2px solid var(--accent)':'none',
    outlineOffset:'-1px',
  })

  // Resumen por mes
  const resumenMeses = MESES.map(m=>{
    const gastado=(data.gastosPresupuesto||[]).filter(g=>g.mes===m&&(!filtroAnio||g.anio===Number(filtroAnio))).reduce((s,g)=>s+(Number(g.valorFactura)||0),0)
    const asignado=data.presupuestos.filter(p=>p.mes===m&&(!filtroAnio||p.anio===Number(filtroAnio))).reduce((s,p)=>s+(Number(p.monto)||0),0)
    return {mes:m,gastado,asignado,ejec:asignado>0?(gastado/asignado)*100:0,disponible:asignado-gastado}
  }).filter(r=>r.gastado>0||r.asignado>0)

  const submitPres = () => {
    if(!formP.mes||!formP.monto) return
    const entry={id:Date.now(),mes:formP.mes,anio:Number(formP.anio),monto:Number(formP.monto)}
    const nd={...data,presupuestos:[...data.presupuestos,entry]};setData(nd);save(nd);setModalPres(false);setFormP({mes:'',anio:2026,monto:''})
  }
  const delPres = id => { const nd={...data,presupuestos:data.presupuestos.filter(p=>p.id!==id)};setData(nd);save(nd) }

  return (
    <div style={{display:'flex',flexDirection:'column',gap:18}}>
      {/* Toolbar */}
      <div style={{display:'flex',gap:8,flexWrap:'wrap',alignItems:'center'}}>
        <div style={{position:'relative',display:'flex',alignItems:'center'}}>
          <Search size={13} style={{position:'absolute',left:9,color:'var(--text3)',pointerEvents:'none'}}/>
          <input value={busqueda} onChange={e=>setBusqueda(e.target.value)} placeholder="Buscar gasto..." style={{paddingLeft:30,width:190,fontSize:13}}/>
        </div>
        <select value={filtroAnio} onChange={e=>setFiltroAnio(e.target.value)} style={{width:100}}>
          <option value="">Año</option>{anios.map(a=><option key={a}>{a}</option>)}
        </select>
        <select value={filtroMes} onChange={e=>setFiltroMes(e.target.value)} style={{width:140}}>
          <option value="">Todos los meses</option>{MESES.map(m=><option key={m}>{m}</option>)}
        </select>
        {(filtroMes||busqueda)&&<button onClick={()=>{setFiltroMes('');setBusqueda('')}} style={{...S.btn('var(--bg3)','var(--text2)'),padding:'5px 10px',fontSize:12}}>✕</button>}
        <div style={{marginLeft:'auto',display:'flex',gap:8,flexWrap:'wrap'}}>
          {seleccionados.size>0&&<button onClick={eliminarSel} style={{...S.btn('var(--red-soft)','var(--red)'),fontSize:12}}><Trash2 size={13}/> Eliminar {seleccionados.size}</button>}
          <button onClick={()=>setModalPres(true)} style={{...S.btn('var(--bg3)','var(--text2)'),border:'1px solid var(--border2)',fontSize:12}}><DollarSign size={13}/> Asignar presupuesto</button>
          <button onClick={addFila} style={{...S.btn('var(--green-soft)','var(--green)'),fontSize:12,border:'1px solid rgba(61,214,140,0.2)'}}><PlusCircle size={14}/> Añadir fila</button>
        </div>
      </div>

      {/* KPIs */}
      <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:12}}>
        <KpiCard icon={DollarSign} label="Presupuesto asignado" value={cop(presAsignado)} sub={filtroMes||(filtroAnio||'')+'  Total período'}/>
        <KpiCard icon={TrendingUp} label="Total ejecutado" value={cop(totalGastado)} sub={gastos.length+' registros'} accent="var(--accent2)"/>
        <KpiCard icon={BarChart2} label="% Ejecutado" value={ejec.toFixed(1)+'%'} sub={ejec>100?'Excedido':ejec>80?'Casi al límite':'Dentro del presupuesto'} accent={ejec>100?'var(--red)':ejec>80?'var(--yellow)':'var(--green)'}/>
        <KpiCard icon={DollarSign} label={disponible>=0?'Disponible':'Excedido'} value={cop(Math.abs(disponible))} accent={disponible>=0?'var(--green)':'var(--red)'}/>
      </div>

      {/* Resumen por mes — siempre visible */}
      {resumenMeses.length>0&&(
        <div style={S.card}>
          <div style={{padding:'12px 18px',borderBottom:'1px solid var(--border)',display:'flex',alignItems:'center',justifyContent:'space-between'}}>
            <h4 style={{fontSize:11,fontWeight:600,color:'var(--text2)',textTransform:'uppercase',letterSpacing:'0.05em'}}>Ejecutado vs Presupuesto por mes</h4>
            <span style={{fontSize:11,color:'var(--text3)'}}>Clic en un mes para filtrar el detalle</span>
          </div>
          <table style={{width:'100%',borderCollapse:'collapse'}}>
            <thead><tr>{['Mes','Presupuesto asignado','Ejecutado','Disponible','% Ejec.','Barra',''].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
            <tbody>
              {resumenMeses.map((r,i)=>(
                <tr key={i} onClick={()=>setFiltroMes(filtroMes===r.mes?'':r.mes)} style={{cursor:'pointer',background:filtroMes===r.mes?'rgba(108,99,255,0.07)':'transparent'}}>
                  <td style={{...S.td,fontWeight:600,color:filtroMes===r.mes?'var(--accent2)':'var(--text)'}}>{r.mes}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',color:'var(--text2)'}}>{cop(r.asignado)}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:500}}>{cop(r.gastado)}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',color:r.disponible>=0?'var(--green)':'var(--red)',fontWeight:500}}>{r.disponible>=0?'':'-'}{cop(Math.abs(r.disponible))}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:600,color:r.ejec>100?'var(--red)':r.ejec>80?'var(--yellow)':'var(--green)'}}>{r.ejec.toFixed(1)}%</td>
                  <td style={{...S.td,minWidth:120}}>
                    <div style={{height:6,background:'var(--bg4)',borderRadius:3}}>
                      <div style={{width:Math.min(r.ejec,100)+'%',height:'100%',background:r.ejec>100?'var(--red)':r.ejec>80?'var(--yellow)':'var(--green)',borderRadius:3}}/>
                    </div>
                  </td>
                  <td style={{...S.td,fontSize:11,color:'var(--text3)'}}>{filtroMes===r.mes?'◀ Filtrando':'Ver →'}</td>
                </tr>
              ))}
              <tr style={{borderTop:'2px solid var(--border2)',background:'var(--bg3)'}}>
                <td style={{...S.td,fontWeight:700}}>TOTAL</td>
                <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700,color:'var(--text2)'}}>{cop(resumenMeses.reduce((s,r)=>s+r.asignado,0))}</td>
                <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700,color:'var(--accent2)'}}>{cop(resumenMeses.reduce((s,r)=>s+r.gastado,0))}</td>
                <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700,color:resumenMeses.reduce((s,r)=>s+r.disponible,0)>=0?'var(--green)':'var(--red)'}}>{cop(resumenMeses.reduce((s,r)=>s+r.disponible,0))}</td>
                <td colSpan={3} style={S.td}/>
              </tr>
            </tbody>
          </table>
        </div>
      )}

      <div style={{fontSize:11,color:'var(--text3)',display:'flex',gap:20,flexWrap:'wrap'}}>
        <span>💡 Clic en celda para editar · Enter confirma · Tab siguiente</span>
        <span>📋 Ctrl+V para pegar desde Excel (Mes, Gasto, Valor, Canal, Observación, Estado)</span>
        <span>☑️ Checkbox para seleccionar y eliminar en bloque</span>
      </div>

      {/* Tabla tipo Excel */}
      <div style={S.card}>
        <div style={{padding:'12px 18px',borderBottom:'1px solid var(--border)',display:'flex',alignItems:'center',justifyContent:'space-between'}}>
          <h4 style={{fontSize:11,fontWeight:600,color:'var(--text2)',textTransform:'uppercase',letterSpacing:'0.05em'}}>
            Detalle de gastos {filtroMes&&'— '+filtroMes} · {gastos.length} registros
          </h4>
          {filtroMes&&<button onClick={()=>setFiltroMes('')} style={{...S.btn('var(--bg3)','var(--text2)'),fontSize:11,padding:'3px 10px'}}>← Ver todos</button>}
        </div>
        <div style={{overflowX:'auto'}} onPaste={handlePaste}>
          <table style={{borderCollapse:'collapse',tableLayout:'fixed',minWidth:'100%'}}>
            <thead>
              <tr style={{background:'var(--bg3)'}}>
                <th style={{...S.th,width:36,textAlign:'center',padding:'8px 6px'}}>
                  <input type="checkbox" checked={seleccionados.size===gastos.length&&gastos.length>0} onChange={toggleTodos} style={{cursor:'pointer',width:14,height:14,accentColor:'var(--accent)'}}/>
                </th>
                <th style={{...S.th,width:36,padding:'8px 4px',textAlign:'center',fontSize:10}}>#</th>
                {COLS_G.map(c=><th key={c.key} style={{...S.th,width:c.w}}>{c.label}</th>)}
              </tr>
            </thead>
            <tbody>
              {gastos.length===0&&(
                <tr><td colSpan={COLS_G.length+2} style={{padding:48,textAlign:'center',color:'var(--text3)',fontSize:13}}>
                  No hay gastos. <strong>Añadir fila</strong>, importar Excel, o pegar con <strong>Ctrl+V</strong>.
                </td></tr>
              )}
              {gastos.map((g,idx)=>(
                <tr key={g.id} style={{background:seleccionados.has(g.id)?'rgba(108,99,255,0.06)':'transparent'}}>
                  <td style={{padding:'6px',textAlign:'center',borderTop:'1px solid var(--border)',borderRight:'1px solid var(--border)'}}>
                    <input type="checkbox" checked={seleccionados.has(g.id)} onChange={()=>toggleSel(g.id)} style={{cursor:'pointer',width:14,height:14,accentColor:'var(--accent)'}}/>
                  </td>
                  <td style={{padding:'6px 4px',textAlign:'center',fontSize:10,color:'var(--text3)',borderTop:'1px solid var(--border)',borderRight:'1px solid var(--border)'}}>{idx+1}</td>
                  {COLS_G.map(col=>(
                    <td key={col.key} style={celda(g.id,col.key)} onClick={()=>startEdit(g.id,col.key,g[col.key])}>
                      {editCell?.id===g.id&&editCell?.field===col.key ? (
                        col.key==='mes'?(
                          <select value={editVal} onChange={e=>setEditVal(e.target.value)} onBlur={commitEdit} onKeyDown={onKD} ref={inputRef}
                            style={{background:'transparent',color:'var(--text)',border:'none',outline:'none',fontSize:12,width:'100%',fontFamily:'var(--font)'}}>
                            <option value="">—</option>{MESES.map(m=><option key={m}>{m}</option>)}
                          </select>
                        ):col.key==='estado'?(
                          <select value={editVal} onChange={e=>setEditVal(e.target.value)} onBlur={commitEdit} onKeyDown={onKD} ref={inputRef}
                            style={{background:'transparent',color:'var(--text)',border:'none',outline:'none',fontSize:12,width:'100%',fontFamily:'var(--font)'}}>
                            {ESTADOS_G.map(s=><option key={s}>{s}</option>)}
                          </select>
                        ):(
                          <input ref={inputRef} value={editVal} onChange={e=>setEditVal(e.target.value)} onBlur={commitEdit} onKeyDown={onKD}
                            type={col.key==='valorFactura'?'number':'text'}
                            style={{background:'transparent',color:'var(--text)',border:'none',outline:'none',fontSize:12,width:'100%',fontFamily:col.key==='valorFactura'?'var(--mono)':'var(--font)'}}/>
                        )
                      ):(
                        <span style={{fontFamily:col.key==='valorFactura'?'var(--mono)':'inherit',color:col.key==='valorFactura'?'var(--accent2)':col.key==='gasto'?'var(--text)':'var(--text2)',fontWeight:col.key==='valorFactura'?500:400}}>
                          {col.key==='valorFactura'?cop(Number(g[col.key])||0):col.key==='estado'?(<Badge label={g[col.key]}/>):(g[col.key]||'')}
                        </span>
                      )}
                    </td>
                  ))}
                </tr>
              ))}
              {gastos.length>0&&(
                <tr style={{background:'var(--bg3)',borderTop:'2px solid var(--border2)'}}>
                  <td colSpan={2} style={{padding:'8px 10px'}}/>
                  <td colSpan={2} style={{padding:'8px 12px',fontWeight:700,fontSize:12,color:'var(--text2)'}}>TOTAL {filtroMes&&'— '+filtroMes}</td>
                  <td style={{padding:'8px 12px',fontFamily:'var(--mono)',fontWeight:700,fontSize:13,color:'var(--accent2)'}}>{cop(totalGastado)}</td>
                  <td colSpan={4} style={{padding:'8px 10px'}}/>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Presupuestos asignados */}
      <div style={S.card}>
        <div style={{padding:'12px 18px',borderBottom:'1px solid var(--border)'}}><h4 style={{fontSize:11,fontWeight:600,color:'var(--text2)',textTransform:'uppercase',letterSpacing:'0.05em'}}>Presupuestos asignados por mes</h4></div>
        <table style={{width:'100%',borderCollapse:'collapse'}}>
          <thead><tr>{['Año','Mes','Monto asignado','Ejecutado','Disponible',''].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
          <tbody>
            {data.presupuestos.length===0&&<tr><td colSpan={6} style={{...S.td,textAlign:'center',color:'var(--text3)',padding:28}}>No hay presupuestos. Usa "Asignar presupuesto".</td></tr>}
            {[...data.presupuestos].sort((a,b)=>a.anio!==b.anio?b.anio-a.anio:MESES.indexOf(a.mes)-MESES.indexOf(b.mes)).map(p=>{
              const ejec2=(data.gastosPresupuesto||[]).filter(g=>g.mes===p.mes&&g.anio===p.anio).reduce((s,g)=>s+(Number(g.valorFactura)||0),0)
              return (
                <tr key={p.id}>
                  <td style={{...S.td,color:'var(--text2)'}}>{p.anio}</td>
                  <td style={{...S.td,fontWeight:500}}>{p.mes}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',color:'var(--text2)'}}>{cop(p.monto)}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',color:'var(--accent2)'}}>{cop(ejec2)}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',color:p.monto-ejec2>=0?'var(--green)':'var(--red)'}}>{cop(p.monto-ejec2)}</td>
                  <td style={S.td}><button onClick={()=>delPres(p.id)} style={{...S.btn('var(--red-soft)','var(--red)'),padding:'4px 8px'}}><Trash2 size={13}/></button></td>
                </tr>
              )
            })}
          </tbody>
        </table>
      </div>

      {/* Modal asignar presupuesto */}
      {modalPres&&(
        <Modal title="Asignar presupuesto mensual" onClose={()=>setModalPres(false)}>
          <div style={{display:'flex',flexDirection:'column',gap:14}}>
            <Field label="Año"><input type="number" value={formP.anio} onChange={e=>setFormP({...formP,anio:e.target.value})} placeholder="2026"/></Field>
            <Field label="Mes *"><select value={formP.mes} onChange={e=>setFormP({...formP,mes:e.target.value})}><option value="">Selecciona...</option>{MESES.map(m=><option key={m}>{m}</option>)}</select></Field>
            <Field label="Monto (COP) *"><input type="number" value={formP.monto} onChange={e=>setFormP({...formP,monto:e.target.value})} placeholder="0"/></Field>
          </div>
          <div style={{display:'flex',gap:10,justifyContent:'flex-end',marginTop:20}}>
            <button onClick={()=>setModalPres(false)} style={S.btn('var(--bg3)','var(--text2)')}>Cancelar</button>
            <button onClick={submitPres} style={S.btn('var(--accent)','#fff')}><Check size={15}/> Guardar</button>
          </div>
        </Modal>
      )}
    </div>
  )
}

// ═══════════════════════════════════════════════════════
// PENDIENTES
// ═══════════════════════════════════════════════════════
function Pendientes({ data, setData }) {
  const [modal, setModal] = useState(false)
  const [filtro, setFiltro] = useState('')
  const [editId, setEditId] = useState(null)
  const blank = { distribuidor:'', tarea:'', categoria:'', fechaLimite:'', prioridad:'Media', estado:'Pendiente', responsable:'', notas:'' }
  const [form, setForm] = useState(blank)
  const distribuidores = [...new Set([...data.pendientes.map(p=>p.distribuidor),...data.inversiones.map(i=>i.distribuidor)])].sort()
  const lista = data.pendientes.filter(p=>!filtro||p.estado===filtro).sort((a,b)=>['Alta','Media','Baja'].indexOf(a.prioridad)-['Alta','Media','Baja'].indexOf(b.prioridad))

  const submit = () => {
    if(!form.distribuidor||!form.tarea) return
    const entry={...form,id:editId||Date.now()}
    const pendientes=editId?data.pendientes.map(p=>p.id===editId?entry:p):[...data.pendientes,entry]
    const nd={...data,pendientes}; setData(nd); save(nd); setModal(false); setEditId(null); setForm(blank)
  }
  const del = id => { const nd={...data,pendientes:data.pendientes.filter(p=>p.id!==id)}; setData(nd); save(nd) }
  const edit = p => { setForm({...p}); setEditId(p.id); setModal(true) }
  const toggle = p => {
    const next={'Pendiente':'En curso','En curso':'Listo','Listo':'Cancelado','Cancelado':'Pendiente'}
    const nd={...data,pendientes:data.pendientes.map(x=>x.id===p.id?{...p,estado:next[p.estado]}:x)}; setData(nd); save(nd)
  }
  const counts={P:data.pendientes.filter(p=>p.estado==='Pendiente').length, E:data.pendientes.filter(p=>p.estado==='En curso').length, L:data.pendientes.filter(p=>p.estado==='Listo').length}

  return (
    <div style={{display:'flex',flexDirection:'column',gap:18}}>
      <div style={{display:'flex',gap:10,alignItems:'center',flexWrap:'wrap'}}>
        <div style={{display:'flex',gap:6}}>
          {[['','Todos',data.pendientes.length],['Pendiente','Pendiente',counts.P],['En curso','En curso',counts.E],['Listo','Listo',counts.L]].map(([v,l,c])=>(
            <button key={v} onClick={()=>setFiltro(v)} style={{...S.btn(filtro===v?'var(--accent-soft)':'var(--bg3)',filtro===v?'var(--accent2)':'var(--text2)'),padding:'5px 12px',fontSize:12}}>
              {l} <span style={{background:'var(--bg4)',borderRadius:4,padding:'0 5px',fontSize:11,marginLeft:2}}>{c}</span>
            </button>
          ))}
        </div>
        <button onClick={()=>{setEditId(null);setForm(blank);setModal(true)}} style={{...S.btn('var(--accent)','#fff'),marginLeft:'auto'}}><PlusCircle size={15}/> Nuevo pendiente</button>
      </div>
      <div style={{display:'flex',flexDirection:'column',gap:10}}>
        {lista.length===0&&<div style={{...S.card,padding:40,textAlign:'center',color:'var(--text3)'}}>No hay pendientes</div>}
        {lista.map(p=>(
          <div key={p.id} style={{...S.card,padding:'14px 18px',display:'flex',gap:14,alignItems:'flex-start'}}>
            <button onClick={()=>toggle(p)} style={{marginTop:3,width:22,height:22,borderRadius:'50%',border:`2px solid ${p.estado==='Listo'?'var(--green)':'var(--border2)'}`,background:p.estado==='Listo'?'var(--green-soft)':'transparent',color:'var(--green)',display:'flex',alignItems:'center',justifyContent:'center',cursor:'pointer',flexShrink:0}}>
              {p.estado==='Listo'&&<Check size={11}/>}
            </button>
            <div style={{flex:1,minWidth:0}}>
              <div style={{display:'flex',alignItems:'center',gap:7,flexWrap:'wrap',marginBottom:5}}>
                <span style={{fontWeight:500,fontSize:14,textDecoration:p.estado==='Cancelado'?'line-through':'none',color:p.estado==='Cancelado'?'var(--text3)':'var(--text)'}}>{p.tarea}</span>
                <Badge label={p.prioridad}/><Badge label={p.estado}/>
              </div>
              <div style={{display:'flex',gap:14,flexWrap:'wrap',fontSize:12,color:'var(--text3)'}}>
                <span style={{color:'var(--text2)',fontWeight:500}}>{p.distribuidor}</span>
                {p.categoria&&<span>{p.categoria}</span>}
                {p.fechaLimite&&<span>📅 {p.fechaLimite}</span>}
                {p.responsable&&<span>👤 {p.responsable}</span>}
              </div>
              {p.notas&&<p style={{marginTop:5,fontSize:12,color:'var(--text3)',borderLeft:'2px solid var(--border2)',paddingLeft:8}}>{p.notas}</p>}
            </div>
            <div style={{display:'flex',gap:6}}>
              <button onClick={()=>edit(p)} style={{...S.btn('var(--bg3)','var(--text2)'),padding:'4px 8px'}}><Edit2 size={13}/></button>
              <button onClick={()=>del(p.id)} style={{...S.btn('var(--red-soft)','var(--red)'),padding:'4px 8px'}}><Trash2 size={13}/></button>
            </div>
          </div>
        ))}
      </div>
      {modal&&(
        <Modal title={editId?'Editar pendiente':'Nuevo pendiente'} onClose={()=>{setModal(false);setEditId(null)}}>
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:14}}>
            <Field label="Distribuidor *" span>
              <input list="dist-pend" value={form.distribuidor} onChange={e=>setForm({...form,distribuidor:e.target.value})} placeholder="Nombre del distribuidor"/>
              <datalist id="dist-pend">{distribuidores.map(d=><option key={d} value={d}/>)}</datalist>
            </Field>
            <Field label="Tarea *" span><input value={form.tarea} onChange={e=>setForm({...form,tarea:e.target.value})} placeholder="Describe la tarea..."/></Field>
            <Field label="Categoría"><input value={form.categoria} onChange={e=>setForm({...form,categoria:e.target.value})} placeholder="Ej: Seguimiento, Diseño..."/></Field>
            <Field label="Fecha límite"><input type="date" value={form.fechaLimite} onChange={e=>setForm({...form,fechaLimite:e.target.value})}/></Field>
            <Field label="Prioridad">
              <select value={form.prioridad} onChange={e=>setForm({...form,prioridad:e.target.value})}>
                {PRIORIDADES.map(p=><option key={p}>{p}</option>)}
              </select>
            </Field>
            <Field label="Estado">
              <select value={form.estado} onChange={e=>setForm({...form,estado:e.target.value})}>
                {ESTADOS_PEND.map(s=><option key={s}>{s}</option>)}
              </select>
            </Field>
            <Field label="Responsable"><input value={form.responsable} onChange={e=>setForm({...form,responsable:e.target.value})} placeholder="Nombre..."/></Field>
            <Field label="Notas" span><textarea value={form.notas} onChange={e=>setForm({...form,notas:e.target.value})} rows={2} style={{resize:'vertical'}} placeholder="Detalles adicionales..."/></Field>
          </div>
          <div style={{display:'flex',gap:10,justifyContent:'flex-end',marginTop:20}}>
            <button onClick={()=>{setModal(false);setEditId(null)}} style={S.btn('var(--bg3)','var(--text2)')}>Cancelar</button>
            <button onClick={submit} style={S.btn('var(--accent)','#fff')}><Check size={15}/> Guardar</button>
          </div>
        </Modal>
      )}
    </div>
  )
}

// ═══════════════════════════════════════════════════════
// EXPORTAR EXCEL
// ═══════════════════════════════════════════════════════
function exportarExcel(data) {
  const wb = XLSX.utils.book_new()
  const ws1 = XLSX.utils.json_to_sheet(data.inversiones.map(i=>({ Fecha:i.fecha, 'Año':i.anio, Mes:i.mes, Distribuidor:i.distribuidor, 'Tipo Plan':i.tipoPlan, Concepto:i.concepto, 'Inversión COP':i.inversion, 'Galones Plan':i.galonesPlan||'', Notas:i.notas })))
  ws1['!cols']=[12,8,12,25,22,20,16,12,30].map(w=>({wch:w}))
  XLSX.utils.book_append_sheet(wb,ws1,'Inversiones')
  const ws2 = XLSX.utils.json_to_sheet(data.ventas.map(v=>({ 'Año':v.anio, Mes:v.mes, Distribuidor:v.distribuidor, Galones:v.galones, 'Venta Neta COP':v.ventaNeta, 'Precio por Galón':v.galones>0?Math.round(v.ventaNeta/v.galones):0 })))
  ws2['!cols']=[8,12,25,12,16,16].map(w=>({wch:w}))
  XLSX.utils.book_append_sheet(wb,ws2,'Ventas')
  const ws3 = XLSX.utils.json_to_sheet(data.presupuestos.map(p=>{ const inv=data.inversiones.filter(i=>i.mes===p.mes&&i.anio===p.anio).reduce((s,i)=>s+(i.inversion||0),0); return {'Año':p.anio,Mes:p.mes,'Presupuesto COP':p.monto,'Invertido COP':inv,'Diferencia':inv-p.monto,'% Ejec':p.monto>0?((inv/p.monto)*100).toFixed(1)+'%':'0%'} }))
  ws3['!cols']=[8,12,16,14,14,10].map(w=>({wch:w}))
  XLSX.utils.book_append_sheet(wb,ws3,'Presupuesto')
  const ws4 = XLSX.utils.json_to_sheet(data.planes.map(p=>({ Distribuidor:p.distribuidor, 'Año':p.anio, Quarter:p.quarter, Estado:p.estado, 'Tipos de Plan':(p.tiposPlan||[]).join(', '), 'Meta Galones':p.metaGalones, 'Meta Venta COP':p.metaVenta, Condiciones:p.condiciones, Acuerdos:p.acuerdos, Notas:p.notas })))
  ws4['!cols']=[25,8,8,15,30,14,16,40,40,30].map(w=>({wch:w}))
  XLSX.utils.book_append_sheet(wb,ws4,'Planes')
  const ws5 = XLSX.utils.json_to_sheet(data.pendientes.map(p=>({ Distribuidor:p.distribuidor, Tarea:p.tarea, Categoria:p.categoria, 'Fecha Limite':p.fechaLimite, Prioridad:p.prioridad, Estado:p.estado, Responsable:p.responsable, Notas:p.notas })))
  ws5['!cols']=[22,35,16,13,10,12,16,28].map(w=>({wch:w}))
  XLSX.utils.book_append_sheet(wb,ws5,'Pendientes')
  const ws6 = XLSX.utils.json_to_sheet((data.gastosPresupuesto||[]).map(g=>({ 'Año':g.anio, Mes:g.mes, 'Gasto (Nom. Producto, Cliente)':g.gasto, 'Valor Factura':g.valorFactura, Canal:g.canal, 'Observación (ATJ, otros)':g.observacion, Estado:g.estado, 'Centro de Costos':g.centroCostos, Notas:g.notas })))
  ws6['!cols']=[8,12,35,16,14,25,12,18,25].map(w=>({wch:w}))
  XLSX.utils.book_append_sheet(wb,ws6,'Presupuesto Gastos')
  XLSX.writeFile(wb,`Tracker_${new Date().toISOString().slice(0,10)}.xlsx`)
}

// ═══════════════════════════════════════════════════════
// IMPORTAR EXCEL
// ═══════════════════════════════════════════════════════
function parsearExcel(file, data, setData, onDone) {
  const reader = new FileReader()
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, {type:'array'})
      const importados = {}
      const leer = name => {
        const ws = wb.Sheets[name] || wb.Sheets[wb.SheetNames.find(n=>n.toLowerCase().includes(name.toLowerCase()))]
        return ws ? XLSX.utils.sheet_to_json(ws, {defval:''}) : null
      }
      const primeraRows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {defval:''})
      const cols0 = primeraRows[0] ? Object.keys(primeraRows[0]) : []
      // Detect sheet type by columns
      const esInv   = cols0.some(c=>c.toLowerCase().includes('distribuidor'))
      const esGasto = cols0.some(c=>c.toLowerCase().includes('gasto')||c.toLowerCase().includes('nom. producto')||c.toLowerCase().includes('valor factura'))

      // Inversiones
      const invRows = leer('Inversiones') || (esInv && !esGasto ? primeraRows : null)
      if(invRows?.length) {
        const nuevas=invRows.filter(r=>r['Distribuidor']||r['distribuidor']).map((r,i)=>({
          id:Date.now()+i, fecha:String(r['Fecha']||''),
          anio:Number(r['Año']||r['Ano']||r['año']||2026), mes:String(r['Mes']||'').trim(),
          distribuidor:String(r['Distribuidor']||r['distribuidor']||'').trim(),
          tipoPlan:String(r['Tipo Plan']||'').trim(), concepto:String(r['Concepto']||'').trim(),
          inversion:parseN(r['Inversión COP']||r['Inversion COP']||r['Inversión']||r['Inversion']||0),
          galonesPlan:r['Galones Plan']?parseN(r['Galones Plan']):'',
          notas:String(r['Notas']||'').trim(),
        }))
        if(nuevas.length){data={...data,inversiones:[...data.inversiones,...nuevas]};importados.Inversiones=nuevas.length}
      }
      // Ventas
      const ventRows = leer('Ventas')
      if(ventRows?.length) {
        const nuevas=ventRows.filter(r=>r['Distribuidor']||r['distribuidor']).map((r,i)=>({
          id:Date.now()+10000+i, anio:Number(r['Año']||r['Ano']||2026),
          mes:String(r['Mes']||'').trim(),
          distribuidor:String(r['Distribuidor']||r['distribuidor']||'').trim(),
          galones:parseN(r['Galones']||0), ventaNeta:parseN(r['Venta Neta COP']||r['Venta Neta']||0), notas:'',
        }))
        if(nuevas.length){data={...data,ventas:[...data.ventas,...nuevas]};importados.Ventas=nuevas.length}
      }
      // Presupuesto mensual
      const presRows = leer('Presupuesto')
      if(presRows?.length) {
        const nuevas=presRows.filter(r=>r['Mes']||r['mes']).map((r,i)=>({
          id:Date.now()+20000+i, anio:Number(r['Año']||r['Ano']||2026),
          mes:String(r['Mes']||r['mes']||'').trim(), monto:parseN(r['Presupuesto COP']||r['Presupuesto']||0),
        }))
        if(nuevas.length){data={...data,presupuestos:[...data.presupuestos,...nuevas]};importados.Presupuesto=nuevas.length}
      }
      // Planes
      const planRows = leer('Planes')
      if(planRows?.length) {
        const nuevas=planRows.filter(r=>r['Distribuidor']||r['distribuidor']).map((r,i)=>({
          id:Date.now()+30000+i, distribuidor:String(r['Distribuidor']||r['distribuidor']||'').trim(),
          anio:Number(r['Año']||r['Ano']||2026), quarter:r['Quarter']||'Q1', estado:r['Estado']||'Activo',
          tiposPlan:String(r['Tipos de Plan']||'').split(',').map(s=>s.trim()).filter(Boolean),
          metaGalones:parseN(r['Meta Galones']||0), metaVenta:parseN(r['Meta Venta COP']||0),
          condiciones:String(r['Condiciones']||''), acuerdos:String(r['Acuerdos']||''),
          notas:String(r['Notas']||''), historial:[],
        }))
        if(nuevas.length){data={...data,planes:[...data.planes,...nuevas]};importados.Planes=nuevas.length}
      }
      // Pendientes
      const pendRows = leer('Pendientes')
      if(pendRows?.length) {
        const nuevas=pendRows.filter(r=>r['Distribuidor']||r['distribuidor']).map((r,i)=>({
          id:Date.now()+40000+i, distribuidor:String(r['Distribuidor']||r['distribuidor']||'').trim(),
          tarea:String(r['Tarea']||r['Tarea / Pendiente']||'').trim(),
          categoria:String(r['Categoria']||r['Categoría']||'').trim(),
          fechaLimite:String(r['Fecha Limite']||r['Fecha Límite']||''),
          prioridad:String(r['Prioridad']||'Media'), estado:String(r['Estado']||'Pendiente'),
          responsable:String(r['Responsable']||''), notas:String(r['Notas']||''),
        }))
        if(nuevas.length){data={...data,pendientes:[...data.pendientes,...nuevas]};importados.Pendientes=nuevas.length}
      }
      // Gastos de presupuesto — detecta Hoja1 si tiene columna Gasto/Valor Factura
      const gastosRows = leer('Presupuesto Gastos') || leer('Control Presupuesto') || (esGasto ? primeraRows : null)
      if(gastosRows?.length) {
        const nuevas=gastosRows
          .filter(r=>(r['Gasto (Nom. Producto, Cliente)']||r['Gasto']||r['B'])&&(r['Mes']||r['A']))
          .map((r,i)=>({
            id:Date.now()+50000+i,
            anio:Number(r['Año']||r['Ano']||2026),
            mes:String(r['Mes']||r['A']||'').trim(),
            gasto:String(r['Gasto (Nom. Producto, Cliente)']||r['Gasto']||r['B']||'').trim(),
            valorFactura:parseN(r['Valor Factura']||r['C']||0),
            canal:String(r['Canal']||r['D']||'').trim(),
            observacion:String(r['Observación (ATJ, otros)']||r['Observacion']||r['E']||'').trim(),
            estado:String(r['Estado']||r['F']||'Pendiente').trim(),
            centroCostos:String(r['Centro de Costos']||r['G']||'').trim(),
            notas:'',
          }))
        if(nuevas.length){
          data={...data,gastosPresupuesto:[...(data.gastosPresupuesto||[]),...nuevas]}
          importados['Presupuesto gastos']=nuevas.length
        }
      }

      setData(data);save(data)
      onDone({importados,errores:[]})
    } catch(err) { onDone({importados:{},errores:['Error: '+err.message]}) }
  }
  reader.readAsArrayBuffer(file)
}

// ═══════════════════════════════════════════════════════
// ASISTENTE IA
// ═══════════════════════════════════════════════════════
function Asistente({ data, setData, onClose }) {
  const [msgs, setMsgs] = useState([{ role:'assistant', content:'¡Hola! Soy tu asistente. Puedo crear pendientes, registrar inversiones, consultar datos y más. ¿En qué te ayudo?' }])
  const [input, setInput] = useState('')
  const [loading, setLoading] = useState(false)
  const bottomRef = useRef(null)
  useEffect(()=>{ bottomRef.current?.scrollIntoView({behavior:'smooth'}) },[msgs])

  const resumen = () => JSON.stringify({
    totalInversiones: data.inversiones.reduce((s,i)=>s+(i.inversion||0),0),
    totalVentaNeta: data.ventas.reduce((s,v)=>s+(v.ventaNeta||0),0),
    pendientesAbiertos: data.pendientes.filter(p=>p.estado!=='Listo'&&p.estado!=='Cancelado').length,
    planesActivos: data.planes.filter(p=>p.estado==='Activo').length,
    distribuidores: [...new Set(data.inversiones.map(i=>i.distribuidor))],
    ultimasInversiones: data.inversiones.slice(-5),
    presupuestos: data.presupuestos,
  })

  const ejecutar = (accionStr, currentData) => {
    try {
      const obj = JSON.parse(accionStr)
      if(obj.tipo==='crear_pendiente'){
        const nuevo={id:Date.now(),distribuidor:obj.distribuidor||'',tarea:obj.tarea||'',categoria:obj.categoria||'',fechaLimite:obj.fechaLimite||'',prioridad:obj.prioridad||'Media',estado:'Pendiente',responsable:obj.responsable||'',notas:obj.notas||''}
        const nd={...currentData,pendientes:[...currentData.pendientes,nuevo]}; setData(nd); save(nd)
        return '✅ Pendiente creado: **'+nuevo.tarea+'** para '+nuevo.distribuidor
      }
      if(obj.tipo==='crear_inversion'){
        const nueva={id:Date.now(),fecha:obj.fecha||new Date().toISOString().slice(0,10),anio:Number(obj.anio||new Date().getFullYear()),mes:obj.mes||'',distribuidor:obj.distribuidor||'',tipoPlan:obj.tipoPlan||'',concepto:obj.concepto||'',inversion:Number(obj.inversion||0),galonesPlan:obj.galonesPlan||'',notas:obj.notas||''}
        const nd={...currentData,inversiones:[...currentData.inversiones,nueva]}; setData(nd); save(nd)
        return '✅ Inversión registrada: '+nueva.distribuidor+' — '+cop(nueva.inversion)
      }
      if(obj.tipo==='crear_presupuesto'){
        const nuevo={id:Date.now(),anio:Number(obj.anio||new Date().getFullYear()),mes:obj.mes||'',monto:Number(obj.monto||0)}
        const nd={...currentData,presupuestos:[...currentData.presupuestos,nuevo]}; setData(nd); save(nd)
        return '✅ Presupuesto creado: '+nuevo.mes+' '+nuevo.anio+' — '+cop(nuevo.monto)
      }
    } catch(e){}
    return null
  }

  const enviar = async () => {
    if(!input.trim()||loading) return
    const userMsg=input.trim(); setInput('')
    const newMsgs=[...msgs,{role:'user',content:userMsg}]
    setMsgs(newMsgs); setLoading(true)
    try {
      const sys=`Eres un asistente de Trade Marketing para Prolub. Datos actuales: ${resumen()}

Cuando el usuario pida crear algo, responde normalmente Y agrega al final:
ACCION:{"tipo":"crear_pendiente","distribuidor":"...","tarea":"...","prioridad":"Alta|Media|Baja","fechaLimite":"YYYY-MM-DD","categoria":"","responsable":"","notas":""}
o
ACCION:{"tipo":"crear_inversion","distribuidor":"...","mes":"...","anio":2026,"tipoPlan":"...","concepto":"...","inversion":0}
o
ACCION:{"tipo":"crear_presupuesto","mes":"...","anio":2026,"monto":0}

Responde en español, conciso y útil.`

      const res=await fetch('https://api.anthropic.com/v1/messages',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({model:'claude-sonnet-4-20250514',max_tokens:800,system:sys,messages:newMsgs.map(m=>({role:m.role,content:m.content}))})})
      const result=await res.json()
      let resp=result.content?.[0]?.text||'No pude procesar la respuesta.'
      const match=resp.match(/ACCION:(\{[\s\S]*?\})/)
      if(match){
        const accionResult=ejecutar(match[1],data)
        resp=resp.replace(/ACCION:\{[\s\S]*?\}/,'').trim()
        if(accionResult) resp=resp+'\n\n'+accionResult
      }
      setMsgs([...newMsgs,{role:'assistant',content:resp}])
    } catch(e){
      setMsgs([...newMsgs,{role:'assistant',content:'❌ Error conectando con el asistente.'}])
    }
    setLoading(false)
  }

  const ejemplos=['¿Cuánto invertimos en total?','Crea un pendiente para LUBRICAFE de revisar meta Q2 prioridad alta','¿Qué distribuidor tiene más inversión?','Registra presupuesto 15 millones Junio 2026']

  return (
    <div style={{position:'fixed',bottom:90,right:24,width:380,height:520,background:'var(--bg2)',border:'1px solid var(--border2)',borderRadius:16,display:'flex',flexDirection:'column',zIndex:300,boxShadow:'0 16px 48px rgba(0,0,0,0.5)'}}>
      <div style={{padding:'14px 18px',borderBottom:'1px solid var(--border)',display:'flex',alignItems:'center',justifyContent:'space-between',background:'var(--bg3)',borderRadius:'16px 16px 0 0'}}>
        <div style={{display:'flex',alignItems:'center',gap:9}}>
          <div style={{width:28,height:28,borderRadius:8,background:'var(--accent)',display:'flex',alignItems:'center',justifyContent:'center'}}><Bot size={15} color="#fff"/></div>
          <div><div style={{fontWeight:600,fontSize:13}}>Asistente Prolub</div><div style={{fontSize:10,color:'var(--green)'}}>● En línea</div></div>
        </div>
        <button onClick={onClose} style={{background:'none',border:'none',color:'var(--text3)',cursor:'pointer',display:'flex'}}><X size={16}/></button>
      </div>
      <div style={{flex:1,overflowY:'auto',padding:'14px 16px',display:'flex',flexDirection:'column',gap:10}}>
        {msgs.map((m,i)=>(
          <div key={i} style={{display:'flex',gap:8,flexDirection:m.role==='user'?'row-reverse':'row',alignItems:'flex-end'}}>
            {m.role==='assistant'&&<div style={{width:26,height:26,borderRadius:7,background:'var(--accent)',display:'flex',alignItems:'center',justifyContent:'center',flexShrink:0}}><Bot size={13} color="#fff"/></div>}
            <div style={{maxWidth:'80%',padding:'9px 13px',borderRadius:m.role==='user'?'12px 12px 2px 12px':'12px 12px 12px 2px',background:m.role==='user'?'var(--accent)':'var(--bg3)',color:m.role==='user'?'#fff':'var(--text)',fontSize:13,lineHeight:1.5,whiteSpace:'pre-wrap'}}>{m.content}</div>
          </div>
        ))}
        {loading&&<div style={{display:'flex',gap:8,alignItems:'flex-end'}}>
          <div style={{width:26,height:26,borderRadius:7,background:'var(--accent)',display:'flex',alignItems:'center',justifyContent:'center'}}><Bot size={13} color="#fff"/></div>
          <div style={{padding:'9px 14px',borderRadius:'12px 12px 12px 2px',background:'var(--bg3)',fontSize:13,color:'var(--text3)'}}>Pensando...</div>
        </div>}
        <div ref={bottomRef}/>
      </div>
      {msgs.length<=1&&(
        <div style={{padding:'0 12px 8px',display:'flex',flexWrap:'wrap',gap:5}}>
          {ejemplos.map((e,i)=>(
            <button key={i} onClick={()=>setInput(e)} style={{fontSize:10,color:'var(--accent2)',background:'var(--accent-soft)',border:'1px solid rgba(108,99,255,0.2)',borderRadius:6,padding:'4px 8px',cursor:'pointer',fontFamily:'var(--font)',textAlign:'left'}}>{e}</button>
          ))}
        </div>
      )}
      <div style={{padding:'10px 12px',borderTop:'1px solid var(--border)',display:'flex',gap:8}}>
        <input value={input} onChange={e=>setInput(e.target.value)} onKeyDown={e=>e.key==='Enter'&&!e.shiftKey&&enviar()} placeholder="Escribe o pregunta algo..." style={{flex:1,fontSize:13,padding:'8px 12px'}} disabled={loading}/>
        <button onClick={enviar} disabled={loading||!input.trim()} style={{background:'var(--accent)',color:'#fff',border:'none',borderRadius:8,padding:'8px 12px',cursor:'pointer',display:'flex',alignItems:'center',opacity:loading||!input.trim()?0.5:1}}><Send size={14}/></button>
      </div>
    </div>
  )
}

// ═══════════════════════════════════════════════════════
// APP PRINCIPAL
// ═══════════════════════════════════════════════════════
const TABS = [
  {id:'dashboard',  label:'Dashboard',   icon:LayoutDashboard},
  {id:'inversiones',label:'Inversiones', icon:TrendingUp},
  {id:'ventas',     label:'Ventas',      icon:ShoppingCart},
  {id:'planes',     label:'Planes Q',    icon:BookOpen},
  {id:'presupuesto',label:'Presupuesto', icon:DollarSign},
  {id:'pendientes', label:'Pendientes',  icon:ListTodo},
]

export default function App() {
  const [tab, setTab] = useState('dashboard')
  const [data, setData] = useState(load)
  const [importResult, setImportResult] = useState(null)
  const [importando, setImportando] = useState(false)
  const [chatOpen, setChatOpen] = useState(false)

  const handleFile = e => {
    const file=e.target.files[0]; if(!file) return
    setImportando(true); setImportResult(null)
    parsearExcel(file,data,setData,result=>{setImportResult(result);setImportando(false)})
    e.target.value=''
  }
  const totalImp = importResult?Object.values(importResult.importados).reduce((s,n)=>s+n,0):0

  return (
    <div style={{minHeight:'100vh',display:'flex',flexDirection:'column'}}>
      <header style={{background:'var(--bg2)',borderBottom:'1px solid var(--border)',padding:'0 28px',display:'flex',alignItems:'center',height:56,gap:24,position:'sticky',top:0,zIndex:10}}>
        <div style={{display:'flex',alignItems:'center',gap:9}}>
          <div style={{width:30,height:30,borderRadius:8,background:'var(--accent)',display:'flex',alignItems:'center',justifyContent:'center'}}><BarChart2 size={16} color="#fff"/></div>
          <span style={{fontWeight:600,fontSize:14,letterSpacing:'-0.01em'}}>Prolub</span>
          <span style={{color:'var(--text3)',fontSize:14}}>/ Trade Marketing</span>
        </div>
        <nav style={{display:'flex',gap:2,marginLeft:'auto',overflowX:'auto'}}>
          {TABS.map(t=>{ const Icon=t.icon; const active=tab===t.id; return (
            <button key={t.id} onClick={()=>setTab(t.id)} style={{display:'flex',alignItems:'center',gap:6,padding:'5px 13px',borderRadius:8,fontSize:12,fontWeight:active?500:400,background:active?'var(--accent-soft)':'transparent',color:active?'var(--accent2)':'var(--text2)',border:'none',cursor:'pointer',fontFamily:'var(--font)',whiteSpace:'nowrap'}}>
              <Icon size={14}/>{t.label}
            </button>
          )})}
        </nav>
      </header>

      <main style={{flex:1,padding:'24px 28px',maxWidth:1280,width:'100%',margin:'0 auto'}}>
        <div key={tab}>
          {tab==='dashboard'   &&<Dashboard    data={data}/>}
          {tab==='inversiones' &&<Inversiones  data={data} setData={setData}/>}
          {tab==='ventas'      &&<Ventas       data={data} setData={setData}/>}
          {tab==='planes'      &&<Planes       data={data} setData={setData}/>}
          {tab==='presupuesto' &&<Presupuesto  data={data} setData={setData}/>}
          {tab==='pendientes'  &&<Pendientes   data={data} setData={setData}/>}
        </div>
      </main>

      <footer style={{padding:'12px 28px',borderTop:'1px solid var(--border)',display:'flex',justifyContent:'space-between',alignItems:'center',flexWrap:'wrap',gap:8}}>
        <span style={{fontSize:11,color:'var(--text3)'}}>Datos guardados en tu navegador</span>
        <div style={{display:'flex',gap:10,alignItems:'center',flexWrap:'wrap'}}>
          <label style={{...S.btn('var(--accent-soft)','var(--accent2)'),fontSize:12,padding:'5px 14px',border:'1px solid rgba(108,99,255,0.25)',cursor:'pointer'}}>
            ↑ Importar Excel
            <input type="file" accept=".xlsx,.xls" onChange={handleFile} style={{display:'none'}}/>
          </label>
          <button onClick={()=>exportarExcel(data)} style={{...S.btn('var(--green-soft)','var(--green)'),fontSize:12,padding:'5px 14px',border:'1px solid rgba(61,214,140,0.2)'}}>
            <Download size={13}/> Exportar Excel
          </button>
          <button onClick={()=>{if(confirm('¿Borrar todos los datos?')){localStorage.removeItem(STORAGE_KEY);window.location.reload()}}} style={{fontSize:11,color:'var(--text3)',background:'none',border:'none',cursor:'pointer',padding:'4px 8px',fontFamily:'var(--font)'}}>
            Resetear
          </button>
        </div>
      </footer>

      {importando&&<div style={{position:'fixed',bottom:80,right:28,background:'var(--bg2)',border:'1px solid var(--border2)',borderRadius:12,padding:'14px 20px',zIndex:200,fontSize:13,color:'var(--accent2)'}}>⏳ Importando datos...</div>}
      {importResult&&(
        <div style={{position:'fixed',bottom:80,right:28,background:'var(--bg2)',border:'1px solid '+(importResult.errores?.length?'var(--red)':'var(--green)'),borderRadius:12,padding:'16px 20px',zIndex:200,maxWidth:320,boxShadow:'0 8px 32px rgba(0,0,0,0.4)'}}>
          <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:10}}>
            <span style={{fontWeight:600,fontSize:13,color:importResult.errores?.length?'var(--red)':'var(--green)'}}>{importResult.errores?.length?'⚠️ Error':'✅ '+totalImp+' registros importados'}</span>
            <button onClick={()=>setImportResult(null)} style={{background:'none',border:'none',color:'var(--text3)',cursor:'pointer'}}><X size={14}/></button>
          </div>
          {Object.entries(importResult.importados).map(([k,v])=><div key={k} style={{fontSize:12,color:'var(--text2)',marginBottom:3}}>✓ <strong>{k}</strong>: {v} filas</div>)}
          {importResult.errores?.map((e,i)=><div key={i} style={{fontSize:12,color:'var(--red)',marginTop:4}}>{e}</div>)}
          <div style={{fontSize:11,color:'var(--text3)',marginTop:8,borderTop:'1px solid var(--border)',paddingTop:8}}>Los datos se agregaron a los existentes</div>
        </div>
      )}

      <button onClick={()=>setChatOpen(o=>!o)}
        style={{position:'fixed',bottom:24,right:24,width:52,height:52,borderRadius:'50%',background:chatOpen?'var(--bg3)':'var(--accent)',color:'#fff',border:chatOpen?'1px solid var(--border2)':'none',cursor:'pointer',display:'flex',alignItems:'center',justifyContent:'center',boxShadow:'0 4px 20px rgba(108,99,255,0.4)',zIndex:250,transition:'all 0.2s'}}>
        {chatOpen?<X size={20}/>:<MessageCircle size={22}/>}
      </button>
      {chatOpen&&<Asistente data={data} setData={setData} onClose={()=>setChatOpen(false)}/>}
    </div>
  )
}
