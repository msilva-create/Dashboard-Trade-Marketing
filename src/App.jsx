import { useState, useRef, useEffect } from 'react'
import { LayoutDashboard, TrendingUp, DollarSign, ListTodo, PlusCircle, Trash2, Edit2, X, Download, BarChart2, ShoppingCart, BookOpen, Check, MessageCircle, Bot, Send } from 'lucide-react'
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
// INVERSIONES
// ═══════════════════════════════════════════════════════
function Inversiones({ data, setData }) {
  const [modal, setModal] = useState(false)
  const [filtroMes, setFiltroMes] = useState('')
  const [filtroAnio, setFiltroAnio] = useState('')
  const [filtroDist, setFiltroDist] = useState('')
  const [editId, setEditId] = useState(null)
  const blank = { fecha:'', anio:2026, mes:'', distribuidor:'', tipoPlan:'Prolub respalda', concepto:'Apoyo a la nomina', inversion:'', galonesPlan:'', notas:'' }
  const [form, setForm] = useState(blank)

  const anios = [...new Set(data.inversiones.map(i=>i.anio))].sort()
  const distribuidores = [...new Set(data.inversiones.map(i=>i.distribuidor))].sort()

  const lista = data.inversiones.filter(i=>
    (!filtroMes||i.mes===filtroMes)&&(!filtroAnio||i.anio===Number(filtroAnio))&&(!filtroDist||i.distribuidor===filtroDist)
  ).sort((a,b)=>b.id-a.id)

  const totalFiltrado = lista.reduce((s,i)=>s+(i.inversion||0),0)
  const presMes = filtroMes ? data.presupuestos.filter(p=>p.mes===filtroMes&&(!filtroAnio||p.anio===Number(filtroAnio))).reduce((s,p)=>s+(p.monto||0),0) : 0

  const submit = () => {
    if(!form.distribuidor||!form.mes) return
    const entry = {...form, id:editId||Date.now(), inversion:Number(form.inversion)||0, galonesPlan:form.galonesPlan?Number(form.galonesPlan):'', anio:Number(form.anio)}
    const inversiones = editId ? data.inversiones.map(i=>i.id===editId?entry:i) : [...data.inversiones,entry]
    const nd = {...data,inversiones}; setData(nd); save(nd); setModal(false); setEditId(null); setForm(blank)
  }
  const del = id => { const nd={...data,inversiones:data.inversiones.filter(i=>i.id!==id)}; setData(nd); save(nd) }
  const edit = inv => { setForm({...inv,inversion:inv.inversion.toString(),galonesPlan:inv.galonesPlan?.toString()||''}); setEditId(inv.id); setModal(true) }

  return (
    <div style={{display:'flex',flexDirection:'column',gap:18}}>
      <div style={{display:'flex',gap:10,flexWrap:'wrap',alignItems:'center'}}>
        <select value={filtroAnio} onChange={e=>setFiltroAnio(e.target.value)} style={{width:110}}>
          <option value="">Año</option>
          {anios.map(a=><option key={a}>{a}</option>)}
        </select>
        <select value={filtroMes} onChange={e=>setFiltroMes(e.target.value)} style={{width:150}}>
          <option value="">Todos los meses</option>
          {MESES.map(m=><option key={m}>{m}</option>)}
        </select>
        <select value={filtroDist} onChange={e=>setFiltroDist(e.target.value)} style={{width:220}}>
          <option value="">Todos los distribuidores</option>
          {distribuidores.map(d=><option key={d}>{d}</option>)}
        </select>
        <button onClick={()=>{setEditId(null);setForm(blank);setModal(true)}} style={{...S.btn('var(--accent)','#fff'),marginLeft:'auto'}}>
          <PlusCircle size={15}/> Nueva inversión
        </button>
      </div>

      {filtroMes&&(
        <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:12}}>
          <KpiCard icon={TrendingUp} label={`Invertido — ${filtroMes}`} value={cop(totalFiltrado)} sub={`${lista.length} registros`} accent="var(--accent2)"/>
          <KpiCard icon={DollarSign} label="Presupuesto del mes" value={cop(presMes)}/>
          <KpiCard icon={BarChart2} label="% Ejecutado" value={presMes>0?`${((totalFiltrado/presMes)*100).toFixed(1)}%`:'—'} accent={presMes>0&&totalFiltrado/presMes>=1?'var(--red)':'var(--green)'}/>
          <KpiCard icon={DollarSign} label={presMes-totalFiltrado>=0?'Disponible':'Excedido'} value={cop(Math.abs(presMes-totalFiltrado))} accent={presMes-totalFiltrado>=0?'var(--green)':'var(--red)'}/>
        </div>
      )}

      <div style={S.card}>
        <table style={{width:'100%',borderCollapse:'collapse'}}>
          <thead><tr>{['Fecha','Año','Mes','Distribuidor','Tipo Plan','Concepto','Inversión','Galones Plan',''].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
          <tbody>
            {lista.length===0&&<tr><td colSpan={9} style={{...S.td,textAlign:'center',color:'var(--text3)',padding:40}}>No hay inversiones registradas</td></tr>}
            {lista.map(inv=>(
              <tr key={inv.id}>
                <td style={{...S.td,color:'var(--text2)',whiteSpace:'nowrap'}}>{inv.fecha||'—'}</td>
                <td style={{...S.td,color:'var(--text2)'}}>{inv.anio}</td>
                <td style={{...S.td,color:'var(--text2)'}}>{inv.mes}</td>
                <td style={{...S.td,fontWeight:500}}>{inv.distribuidor}</td>
                <td style={{...S.td,color:'var(--text2)',fontSize:12}}>{inv.tipoPlan}</td>
                <td style={{...S.td}}><Badge label={inv.concepto}/></td>
                <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:500}}>{cop(inv.inversion)}</td>
                <td style={{...S.td,color:'var(--text2)',textAlign:'center'}}>{inv.galonesPlan||'—'}</td>
                <td style={S.td}>
                  <div style={{display:'flex',gap:6}}>
                    <button onClick={()=>edit(inv)} style={{...S.btn('var(--bg3)','var(--text2)'),padding:'4px 8px'}}><Edit2 size={13}/></button>
                    <button onClick={()=>del(inv.id)} style={{...S.btn('var(--red-soft)','var(--red)'),padding:'4px 8px'}}><Trash2 size={13}/></button>
                  </div>
                </td>
              </tr>
            ))}
            {lista.length>0&&(
              <tr style={{borderTop:'2px solid var(--border2)'}}>
                <td colSpan={6} style={{...S.td,fontWeight:700,color:'var(--text2)'}}>TOTAL {filtroMes&&`— ${filtroMes}`}</td>
                <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700,color:'var(--accent2)'}}>{cop(totalFiltrado)}</td>
                <td colSpan={2} style={S.td}/>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {modal&&(
        <Modal title={editId?'Editar inversión':'Nueva inversión'} onClose={()=>{setModal(false);setEditId(null)}} wide>
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:14}}>
            <Field label="Fecha"><input type="date" value={form.fecha} onChange={e=>setForm({...form,fecha:e.target.value})}/></Field>
            <Field label="Año"><input type="number" value={form.anio} onChange={e=>setForm({...form,anio:e.target.value})} placeholder="2026"/></Field>
            <Field label="Mes *">
              <select value={form.mes} onChange={e=>setForm({...form,mes:e.target.value})}>
                <option value="">Selecciona...</option>
                {MESES.map(m=><option key={m}>{m}</option>)}
              </select>
            </Field>
            <Field label="Distribuidor *">
              <input list="dist-inv" value={form.distribuidor} onChange={e=>setForm({...form,distribuidor:e.target.value})} placeholder="Nombre del distribuidor"/>
              <datalist id="dist-inv">{distribuidores.map(d=><option key={d} value={d}/>)}</datalist>
            </Field>
            <Field label="Tipo Plan">
              <select value={form.tipoPlan} onChange={e=>setForm({...form,tipoPlan:e.target.value})}>
                {TIPOS_PLAN.map(t=><option key={t}>{t}</option>)}
              </select>
            </Field>
            <Field label="Concepto">
              <select value={form.concepto} onChange={e=>setForm({...form,concepto:e.target.value})}>
                {CONCEPTOS.map(c=><option key={c}>{c}</option>)}
              </select>
            </Field>
            <Field label="Inversión (COP) *"><input type="number" value={form.inversion} onChange={e=>setForm({...form,inversion:e.target.value})} placeholder="0"/></Field>
            <Field label="Galones Plan"><input type="number" value={form.galonesPlan} onChange={e=>setForm({...form,galonesPlan:e.target.value})} placeholder="Opcional"/></Field>
            <Field label="Notas" span><textarea value={form.notas} onChange={e=>setForm({...form,notas:e.target.value})} rows={2} style={{resize:'vertical'}} placeholder="Observaciones..."/></Field>
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
// PRESUPUESTO
// ═══════════════════════════════════════════════════════
function Presupuesto({ data, setData }) {
  const [modalGasto, setModalGasto] = useState(false)
  const [modalPres, setModalPres] = useState(false)
  const [filtroMes, setFiltroMes] = useState('')
  const [filtroAnio, setFiltroAnio] = useState('2026')
  const [editId, setEditId] = useState(null)
  const blankGasto = { mes:'', anio:2026, gasto:'', valorFactura:'', canal:'', observacion:'', estado:'Pendiente', centroCostos:'', notas:'' }
  const blankPres = { mes:'', anio:2026, monto:'' }
  const [formGasto, setFormGasto] = useState(blankGasto)
  const [formPres, setFormPres] = useState(blankPres)

  const CANALES_PRES = ['Digital','ATL','BTL','Trade','Eventos','POP','Nomina','Otro']
  const ESTADOS_GASTO = ['Pendiente','Aprobado','Pagado','Rechazado']

  const anios = [...new Set([...data.gastosPresupuesto?.map(g=>g.anio)||[],...data.presupuestos.map(p=>p.anio)])].sort()

  const gastosFiltrados = (data.gastosPresupuesto||[]).filter(g=>
    (!filtroMes||g.mes===filtroMes)&&(!filtroAnio||g.anio===Number(filtroAnio))
  ).sort((a,b)=>MESES.indexOf(a.mes)-MESES.indexOf(b.mes))

  const totalGastado = gastosFiltrados.reduce((s,g)=>s+(g.valorFactura||0),0)
  const presAsignado = data.presupuestos.filter(p=>(!filtroMes||p.mes===filtroMes)&&(!filtroAnio||p.anio===Number(filtroAnio))).reduce((s,p)=>s+(p.monto||0),0)
  const disponible = presAsignado - totalGastado
  const ejec = presAsignado>0 ? (totalGastado/presAsignado)*100 : 0

  const submitGasto = () => {
    if(!formGasto.mes||!formGasto.gasto) return
    const entry = {...formGasto, id:editId||Date.now(), valorFactura:Number(formGasto.valorFactura)||0, anio:Number(formGasto.anio)}
    const gastosPresupuesto = editId
      ? (data.gastosPresupuesto||[]).map(g=>g.id===editId?entry:g)
      : [...(data.gastosPresupuesto||[]),entry]
    const nd={...data,gastosPresupuesto}; setData(nd); save(nd); setModalGasto(false); setEditId(null); setFormGasto(blankGasto)
  }

  const submitPres = () => {
    if(!formPres.mes||!formPres.monto) return
    const entry = {...formPres, id:Date.now(), monto:Number(formPres.monto), anio:Number(formPres.anio)}
    const presupuestos = [...data.presupuestos, entry]
    const nd={...data,presupuestos}; setData(nd); save(nd); setModalPres(false); setFormPres(blankPres)
  }

  const delGasto = id => { const nd={...data,gastosPresupuesto:(data.gastosPresupuesto||[]).filter(g=>g.id!==id)}; setData(nd); save(nd) }
  const delPres = id => { const nd={...data,presupuestos:data.presupuestos.filter(p=>p.id!==id)}; setData(nd); save(nd) }
  const editGasto = g => { setFormGasto({...g,valorFactura:g.valorFactura.toString()}); setEditId(g.id); setModalGasto(true) }

  // Resumen por mes
  const resumenMeses = MESES.map(m=>{
    const gastado = (data.gastosPresupuesto||[]).filter(g=>g.mes===m&&(!filtroAnio||g.anio===Number(filtroAnio))).reduce((s,g)=>s+(g.valorFactura||0),0)
    const asignado = data.presupuestos.filter(p=>p.mes===m&&(!filtroAnio||p.anio===Number(filtroAnio))).reduce((s,p)=>s+(p.monto||0),0)
    return { mes:m, gastado, asignado, ejec:asignado>0?(gastado/asignado)*100:0 }
  }).filter(r=>r.gastado>0||r.asignado>0)

  return (
    <div style={{display:'flex',flexDirection:'column',gap:18}}>
      {/* Filtros y botones */}
      <div style={{display:'flex',gap:10,flexWrap:'wrap',alignItems:'center'}}>
        <select value={filtroAnio} onChange={e=>setFiltroAnio(e.target.value)} style={{width:110}}>
          <option value=''>Año</option>
          {anios.map(a=><option key={a}>{a}</option>)}
        </select>
        <select value={filtroMes} onChange={e=>setFiltroMes(e.target.value)} style={{width:150}}>
          <option value=''>Todos los meses</option>
          {MESES.map(m=><option key={m}>{m}</option>)}
        </select>
        <div style={{display:'flex',gap:8,marginLeft:'auto'}}>
          <button onClick={()=>{setModalPres(true)}} style={{...S.btn('var(--bg3)','var(--text2)'),border:'1px solid var(--border2)'}}>
            <DollarSign size={14}/> Asignar presupuesto
          </button>
          <button onClick={()=>{setEditId(null);setFormGasto(blankGasto);setModalGasto(true)}} style={S.btn('var(--accent)','#fff')}>
            <PlusCircle size={15}/> Nuevo gasto
          </button>
        </div>
      </div>

      {/* KPIs */}
      <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:12}}>
        <KpiCard icon={DollarSign} label='Presupuesto asignado' value={cop(presAsignado)} sub={filtroMes||'Total período'}/>
        <KpiCard icon={TrendingUp} label='Total gastado' value={cop(totalGastado)} sub={gastosFiltrados.length+' registros'} accent='var(--accent2)'/>
        <KpiCard icon={BarChart2} label='% Ejecutado' value={ejec.toFixed(1)+'%'} accent={ejec>100?'var(--red)':ejec>80?'var(--yellow)':'var(--green)'}/>
        <KpiCard icon={DollarSign} label={disponible>=0?'Disponible':'Excedido'} value={cop(Math.abs(disponible))} accent={disponible>=0?'var(--green)':'var(--red)'}/>
      </div>

      {/* Resumen por mes */}
      {!filtroMes&&resumenMeses.length>0&&(
        <div style={S.card}>
          <div style={{padding:'12px 18px',borderBottom:'1px solid var(--border)'}}><h4 style={{fontSize:11,fontWeight:600,color:'var(--text2)',textTransform:'uppercase',letterSpacing:'0.05em'}}>Resumen por mes</h4></div>
          <table style={{width:'100%',borderCollapse:'collapse'}}>
            <thead><tr>{['Mes','Presupuesto','Gastado','Disponible','% Ejec.',''].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
            <tbody>
              {resumenMeses.map((r,i)=>(
                <tr key={i} style={{cursor:'pointer'}} onClick={()=>setFiltroMes(r.mes)}>
                  <td style={{...S.td,fontWeight:500}}>{r.mes}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',color:'var(--text2)'}}>{cop(r.asignado)}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)'}}>{cop(r.gastado)}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',color:r.asignado-r.gastado>=0?'var(--green)':'var(--red)'}}>{cop(r.asignado-r.gastado)}</td>
                  <td style={{...S.td}}>
                    <div style={{display:'flex',alignItems:'center',gap:8}}>
                      <div style={{flex:1,height:5,background:'var(--bg4)',borderRadius:3,minWidth:80}}>
                        <div style={{width:Math.min(r.ejec,100)+'%',height:'100%',background:r.ejec>100?'var(--red)':r.ejec>80?'var(--yellow)':'var(--green)',borderRadius:3}}/>
                      </div>
                      <span style={{fontFamily:'var(--mono)',fontSize:12,fontWeight:600,color:r.ejec>100?'var(--red)':r.ejec>80?'var(--yellow)':'var(--green)',minWidth:45}}>{r.ejec.toFixed(1)}%</span>
                    </div>
                  </td>
                  <td style={{...S.td,fontSize:11,color:'var(--text3)'}}>Ver detalle →</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {/* Detalle de gastos */}
      <div style={S.card}>
        <div style={{padding:'12px 18px',borderBottom:'1px solid var(--border)',display:'flex',alignItems:'center',justifyContent:'space-between'}}>
          <h4 style={{fontSize:11,fontWeight:600,color:'var(--text2)',textTransform:'uppercase',letterSpacing:'0.05em'}}>
            Detalle de gastos {filtroMes&&'— '+filtroMes}
          </h4>
          {filtroMes&&<button onClick={()=>setFiltroMes('')} style={{...S.btn('var(--bg3)','var(--text2)'),fontSize:11,padding:'3px 10px'}}>← Todos los meses</button>}
        </div>
        <table style={{width:'100%',borderCollapse:'collapse'}}>
          <thead><tr>{['Mes','Gasto (Producto/Cliente)','Valor Factura','Canal','Observación','Estado','Centro Costos',''].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
          <tbody>
            {gastosFiltrados.length===0&&<tr><td colSpan={8} style={{...S.td,textAlign:'center',color:'var(--text3)',padding:40}}>No hay gastos registrados {filtroMes&&'en '+filtroMes}</td></tr>}
            {gastosFiltrados.map(g=>(
              <tr key={g.id}>
                <td style={{...S.td,color:'var(--text2)',whiteSpace:'nowrap'}}>{g.mes}</td>
                <td style={{...S.td,fontWeight:500,maxWidth:200}}>{g.gasto}</td>
                <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:500,color:'var(--accent2)'}}>{cop(g.valorFactura)}</td>
                <td style={{...S.td,fontSize:12,color:'var(--text2)'}}>{g.canal}</td>
                <td style={{...S.td,fontSize:12,color:'var(--text2)',maxWidth:180}}>{g.observacion}</td>
                <td style={{...S.td}}><Badge label={g.estado}/></td>
                <td style={{...S.td,fontSize:12,color:'var(--text2)'}}>{g.centroCostos}</td>
                <td style={S.td}>
                  <div style={{display:'flex',gap:6}}>
                    <button onClick={()=>editGasto(g)} style={{...S.btn('var(--bg3)','var(--text2)'),padding:'4px 8px'}}><Edit2 size={13}/></button>
                    <button onClick={()=>delGasto(g.id)} style={{...S.btn('var(--red-soft)','var(--red)'),padding:'4px 8px'}}><Trash2 size={13}/></button>
                  </div>
                </td>
              </tr>
            ))}
            {gastosFiltrados.length>0&&(
              <tr style={{borderTop:'2px solid var(--border2)'}}>
                <td colSpan={2} style={{...S.td,fontWeight:700,color:'var(--text2)'}}>TOTAL {filtroMes&&filtroMes}</td>
                <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700,color:'var(--accent2)'}}>{cop(totalGastado)}</td>
                <td colSpan={5} style={S.td}/>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {/* Presupuestos asignados */}
      <div style={S.card}>
        <div style={{padding:'12px 18px',borderBottom:'1px solid var(--border)',display:'flex',alignItems:'center',justifyContent:'space-between'}}>
          <h4 style={{fontSize:11,fontWeight:600,color:'var(--text2)',textTransform:'uppercase',letterSpacing:'0.05em'}}>Presupuestos asignados</h4>
        </div>
        <table style={{width:'100%',borderCollapse:'collapse'}}>
          <thead><tr>{['Año','Mes','Monto asignado',''].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
          <tbody>
            {data.presupuestos.length===0&&<tr><td colSpan={4} style={{...S.td,textAlign:'center',color:'var(--text3)',padding:32}}>No hay presupuestos asignados</td></tr>}
            {[...data.presupuestos].sort((a,b)=>a.anio!==b.anio?b.anio-a.anio:MESES.indexOf(a.mes)-MESES.indexOf(b.mes)).map(p=>(
              <tr key={p.id}>
                <td style={{...S.td,color:'var(--text2)'}}>{p.anio}</td>
                <td style={{...S.td,fontWeight:500}}>{p.mes}</td>
                <td style={{...S.td,fontFamily:'var(--mono)',color:'var(--green)'}}>{cop(p.monto)}</td>
                <td style={S.td}><button onClick={()=>delPres(p.id)} style={{...S.btn('var(--red-soft)','var(--red)'),padding:'4px 8px'}}><Trash2 size={13}/></button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Modal nuevo gasto */}
      {modalGasto&&(
        <Modal title={editId?'Editar gasto':'Nuevo gasto de presupuesto'} onClose={()=>{setModalGasto(false);setEditId(null)}} wide>
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:14}}>
            <Field label='Año'><input type='number' value={formGasto.anio} onChange={e=>setFormGasto({...formGasto,anio:e.target.value})} placeholder='2026'/></Field>
            <Field label='Mes *'>
              <select value={formGasto.mes} onChange={e=>setFormGasto({...formGasto,mes:e.target.value})}>
                <option value=''>Selecciona...</option>
                {MESES.map(m=><option key={m}>{m}</option>)}
              </select>
            </Field>
            <Field label='Gasto (Producto/Cliente) *' span>
              <input value={formGasto.gasto} onChange={e=>setFormGasto({...formGasto,gasto:e.target.value})} placeholder='Ej: GORRASCIDENTEX, Evento lanzamiento...'/>
            </Field>
            <Field label='Valor Factura (COP)'>
              <input type='number' value={formGasto.valorFactura} onChange={e=>setFormGasto({...formGasto,valorFactura:e.target.value})} placeholder='0'/>
            </Field>
            <Field label='Canal'>
              <select value={formGasto.canal} onChange={e=>setFormGasto({...formGasto,canal:e.target.value})}>
                <option value=''>Selecciona...</option>
                {CANALES_PRES.map(c=><option key={c}>{c}</option>)}
              </select>
            </Field>
            <Field label='Observación (ATJ, otros)'>
              <input value={formGasto.observacion} onChange={e=>setFormGasto({...formGasto,observacion:e.target.value})} placeholder='Ej: ATJ, aprobado por...'/>
            </Field>
            <Field label='Estado'>
              <select value={formGasto.estado} onChange={e=>setFormGasto({...formGasto,estado:e.target.value})}>
                {ESTADOS_GASTO.map(s=><option key={s}>{s}</option>)}
              </select>
            </Field>
            <Field label='Centro de Costos'>
              <input value={formGasto.centroCostos} onChange={e=>setFormGasto({...formGasto,centroCostos:e.target.value})} placeholder='Ej: Katherine, Marketing...'/>
            </Field>
            <Field label='Notas' span>
              <textarea value={formGasto.notas} onChange={e=>setFormGasto({...formGasto,notas:e.target.value})} rows={2} style={{resize:'vertical'}} placeholder='Observaciones adicionales...'/>
            </Field>
          </div>
          <div style={{display:'flex',gap:10,justifyContent:'flex-end',marginTop:20}}>
            <button onClick={()=>{setModalGasto(false);setEditId(null)}} style={S.btn('var(--bg3)','var(--text2)')}>Cancelar</button>
            <button onClick={submitGasto} style={S.btn('var(--accent)','#fff')}><Check size={15}/> Guardar</button>
          </div>
        </Modal>
      )}

      {/* Modal asignar presupuesto */}
      {modalPres&&(
        <Modal title='Asignar presupuesto mensual' onClose={()=>setModalPres(false)}>
          <div style={{display:'flex',flexDirection:'column',gap:14}}>
            <Field label='Año'><input type='number' value={formPres.anio} onChange={e=>setFormPres({...formPres,anio:e.target.value})} placeholder='2026'/></Field>
            <Field label='Mes *'>
              <select value={formPres.mes} onChange={e=>setFormPres({...formPres,mes:e.target.value})}>
                <option value=''>Selecciona...</option>
                {MESES.map(m=><option key={m}>{m}</option>)}
              </select>
            </Field>
            <Field label='Monto presupuestado (COP) *'>
              <input type='number' value={formPres.monto} onChange={e=>setFormPres({...formPres,monto:e.target.value})} placeholder='0'/>
            </Field>
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
      const wb = XLSX.read(e.target.result,{type:'array'})
      const leer = name => { 
        const ws=wb.Sheets[name] || wb.Sheets[wb.SheetNames.find(n=>n.toLowerCase().includes(name.toLowerCase()))]
        return ws?XLSX.utils.sheet_to_json(ws,{defval:''}):null 
      }
      const leerPrimera = () => {
        const ws=wb.Sheets[wb.SheetNames[0]]
        return ws?XLSX.utils.sheet_to_json(ws,{defval:''}):null
      }
      const importados = {}

      const inv = leer('Inversiones') || leerPrimera()
      if(inv?.length){ const nuevas=inv.filter(r=>r['Distribuidor']||r['distribuidor']).map((r,i)=>({ id:Date.now()+i, fecha:r['Fecha']||'', anio:Number(r['Año']||r['Ano']||new Date().getFullYear()), mes:r['Mes']||'', distribuidor:r['Distribuidor']||r['distribuidor']||'', tipoPlan:r['Tipo Plan']||'', concepto:r['Concepto']||'', inversion:(()=>{ const v=r['Inversión COP']||r['Inversion COP']||r['Inversión']||r['Inversion']||r['inversion']||0; if(typeof v==='number') return v; const clean=String(v).replace(/[$s]/g,'').replace(/./g,'').replace(',','.'); return parseFloat(clean)||0 })(), galonesPlan:r['Galones Plan']?String(r['Galones Plan']).replace(',','.')*1||'':'', notas:r['Notas']||'' })); if(nuevas.length){data={...data,inversiones:[...data.inversiones,...nuevas]};importados.Inversiones=nuevas.length} }

      const vent = leer('Ventas')
      if(vent?.length){ const nuevas=vent.filter(r=>r['Distribuidor']||r['distribuidor']).map((r,i)=>({ id:Date.now()+10000+i, anio:Number(r['Año']||r['Ano']||new Date().getFullYear()), mes:r['Mes']||'', distribuidor:r['Distribuidor']||r['distribuidor']||'', galones:Number(r['Galones']||0), ventaNeta:Number(r['Venta Neta COP']||r['Venta Neta']||0), notas:r['Notas']||'' })); if(nuevas.length){data={...data,ventas:[...data.ventas,...nuevas]};importados.Ventas=nuevas.length} }

      const pres = leer('Presupuesto')
      if(pres?.length){ const nuevas=pres.filter(r=>r['Mes']||r['mes']).map((r,i)=>({ id:Date.now()+20000+i, anio:Number(r['Año']||r['Ano']||new Date().getFullYear()), mes:r['Mes']||r['mes']||'', monto:Number(r['Presupuesto COP']||r['Presupuesto']||0) })); if(nuevas.length){data={...data,presupuestos:[...data.presupuestos,...nuevas]};importados.Presupuesto=nuevas.length} }

      const plan = leer('Planes')
      if(plan?.length){ const nuevas=plan.filter(r=>r['Distribuidor']||r['distribuidor']).map((r,i)=>({ id:Date.now()+30000+i, distribuidor:r['Distribuidor']||r['distribuidor']||'', anio:Number(r['Año']||r['Ano']||new Date().getFullYear()), quarter:r['Quarter']||'Q1', estado:r['Estado']||'Activo', tiposPlan:(r['Tipos de Plan']||'').split(',').map(s=>s.trim()).filter(Boolean), metaGalones:Number(r['Meta Galones']||0), metaVenta:Number(r['Meta Venta COP']||0), condiciones:r['Condiciones']||'', acuerdos:r['Acuerdos']||'', notas:r['Notas']||'', historial:[] })); if(nuevas.length){data={...data,planes:[...data.planes,...nuevas]};importados.Planes=nuevas.length} }

      const pend = leer('Pendientes')
      if(pend?.length){ const nuevas=pend.filter(r=>r['Distribuidor']||r['distribuidor']).map((r,i)=>({ id:Date.now()+40000+i, distribuidor:r['Distribuidor']||r['distribuidor']||'', tarea:r['Tarea']||r['Tarea / Pendiente']||'', categoria:r['Categoria']||r['Categoría']||'', fechaLimite:r['Fecha Limite']||r['Fecha Límite']||'', prioridad:r['Prioridad']||'Media', estado:r['Estado']||'Pendiente', responsable:r['Responsable']||'', notas:r['Notas']||'' })); if(nuevas.length){data={...data,pendientes:[...data.pendientes,...nuevas]};importados.Pendientes=nuevas.length} }

      // Presupuesto Gastos sheet
      const presGast = leer('Presupuesto Gastos') || leer('Control Presupuesto') || leer('Hoja2')
      if(presGast?.length){ const nuevas=presGast.filter(r=>r['Gasto (Nom. Producto, Cliente)']||r['Gasto']||r['B']).map((r,i)=>({ id:Date.now()+50000+i, anio:Number(r['Año']||r['Ano']||r['A']||new Date().getFullYear()), mes:r['Mes']||r['A']||'', gasto:r['Gasto (Nom. Producto, Cliente)']||r['Gasto']||r['B']||'', valorFactura:(()=>{ const v=r['Valor Factura']||r['Valor Factu']||r['C']||0; if(typeof v==='number') return v; const c=String(v).replace(/[$s]/g,'').replace(/./g,'').replace(',','.'); return parseFloat(c)||0 })(), canal:r['Canal']||r['D']||'', observacion:r['Observación (ATJ, otros)']||r['Observacion']||r['E']||'', estado:r['Estado']||r['F']||'Pendiente', centroCostos:r['Centro de Costos']||r['G']||'', notas:r['Notas']||'' })); if(nuevas.length){data={...data,gastosPresupuesto:[...(data.gastosPresupuesto||[]),...nuevas]};importados['Presupuesto Gastos']=nuevas.length} }

      setData(data); save(data)
      onDone({importados, errores:[]})
    } catch(err){ onDone({importados:{}, errores:['Error: '+err.message]}) }
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
