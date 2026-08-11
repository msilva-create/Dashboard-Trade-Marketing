import { useState, useRef, useEffect } from 'react'
import { LayoutDashboard, TrendingUp, DollarSign, ListTodo, PlusCircle, Trash2, Edit2, X, Download, BarChart2, Bell, ShoppingCart, BookOpen, Check, MessageCircle, Bot, Send, Search } from 'lucide-react'
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, Legend } from 'recharts'
import * as XLSX from 'xlsx'

const STORAGE_KEY = 'tracker_v3'
let _currentUserKey = STORAGE_KEY

const SHEETS_CONFIG = {
  distribucion: {
    url: 'https://script.google.com/macros/s/AKfycbxzqhKZr7ue545BDwpgqQrNsognsVOtxxZNPd6P4Jm_wtdj1_Sqkd9FdswrdtfM2d8S/exec',
    spreadsheetId: '15X2WrJGDw__xFvPpe3Gcm5iR9KSKgJP20Q_kyJ0Zn0U',
    enabled: true,
  },
  industria: {
    url: 'https://script.google.com/macros/s/AKfycbxzqhKZr7ue545BDwpgqQrNsognsVOtxxZNPd6P4Jm_wtdj1_Sqkd9FdswrdtfM2d8S/exec',
    spreadsheetId: '1YOqpzgEyLWdqAvXLIQBW9YrQ1KpnvUnEjZvHkY5Ce4A',
    enabled: true,
  },
  zonas: {
    url: 'https://script.google.com/macros/s/AKfycbxzqhKZr7ue545BDwpgqQrNsognsVOtxxZNPd6P4Jm_wtdj1_Sqkd9FdswrdtfM2d8S/exec',
    spreadsheetId: '1OHpuBQ5Aj-NV9n-wQBlkUoT8mgC3DgZhWyAVSAfFL7Y',
    enabled: true,
  },
  juan: {
    url: 'https://script.google.com/macros/s/AKfycbxzqhKZr7ue545BDwpgqQrNsognsVOtxxZNPd6P4Jm_wtdj1_Sqkd9FdswrdtfM2d8S/exec',
    spreadsheetId: '1s6fkyyCihzOki9Fan9EKIIvxwlra-QizwI7fy210GUI',
    enabled: true,
  },
  camilo: {
    url: 'https://script.google.com/macros/s/AKfycbxzqhKZr7ue545BDwpgqQrNsognsVOtxxZNPd6P4Jm_wtdj1_Sqkd9FdswrdtfM2d8S/exec',
    spreadsheetId: '1jVKNc8krqeClgfzaXr-i7YIHwuYRmjkoU35NEjz4oDk',
    enabled: true,
  },
  julian: {
    url: 'https://script.google.com/macros/s/AKfycbxzqhKZr7ue545BDwpgqQrNsognsVOtxxZNPd6P4Jm_wtdj1_Sqkd9FdswrdtfM2d8S/exec',
    spreadsheetId: '1nEdGAcy45hMckYxOC_SRfrYEqW6ifgRv9fFoJA75YlY',
    enabled: true,
  },
}

const EMAILS_POR_CANAL = {
  distribucion: {
    'Cristian': 'cblanco@prolub.com.co',
    'Mauricio': 'oramirez@prolub.com.co',
    'Gonzalo':  'grodriguez@prolub.com.co',
  },
  industria: {
    'Angela Pastor':        'apastor@prolub.com.co',
    'Maria Fernanda':       'mquiroga@prolub.com.co',
    'Aylin Cepeda':         'acepeda@prolub.com.co',
    'Juan Pablo Rodriguez': 'jrodriguez@prolub.com.co',
  },
  zonas: {
    'Andres Hernandez': 'ahernandez@prolub.com.co',
    'Maria Fajardo':    'mfajardo@prolub.com.co',
    'Luis Ferro':       'lferro@prolub.com.co',
    'Ruben Escarraga':  'rescarraga@prolub.com.co',
    'Juan Cubillos':    'jcubillos@prolub.com.co',
    'Johan Santibanez': 'jsantibanez@prolub.com.co',
    'Jhonn Florez':     'jflorez@prolub.com.co',
    'Sandra Ceron':     'sceron@gulfcolombia.com',
    'Hans Munoz':       'hdmunoz@prolub.com.co',
    'Victor Molina':    'vmolina@prolub.com.co',
  },
}

const JEFES_POR_CANAL = {
  distribucion: { 'Maria Fernanda': 'mquiroga@prolub.com.co' },
  industria:    { 'Maria Fernanda': 'mquiroga@prolub.com.co' },
  zonas:        { 'Victor': 'vmolina@prolub.com.co', 'Carlos': 'cmercado@prolub.com.co' },
}

const HOJA_MAP = {
  inversiones:       'INVERSIONES',
  ventas:            'VENTAS',
  presupuestos:      'PRESUPUESTO',
  gastosPresupuesto: 'GASTOS_PRESUPUESTO',
  planes:            'PLANES',
  apoyoCierre:       'APOYO_CIERRE',
  redenciones:       'REDENCIONES_APOYO',
  pendientes:        'PENDIENTES',
  tareasAsignadas:   'TAREAS_ASIGNADAS',
  misTareas:          'MIS_TAREAS',
  proyectos:          'PROYECTOS',
  subtareasProyectos: 'SUBTAREAS_PROYECTOS',
}

const CAMPOS_MAP = {
  inversiones:  { id:'id', fecha:'Fecha', anio:'Año', mes:'Mes', distribuidor:'Distribuidor', tipoPlan:'Tipo Plan', concepto:'Concepto', inversion:'Inversión', galonesPlan:'Gal. Plan', notas:'Notas' },
  ventas:       { id:'id', anio:'Año', mes:'Mes', distribuidor:'Distribuidor', galones:'Galones', ventaNeta:'Venta Neta', notas:'Notas' },
  presupuestos: { id:'id', anio:'Año', mes:'Mes', monto:'Monto' },
  gastosPresupuesto: { id:'id', anio:'año', mes:'Mes', gasto:'Gasto (Nom. Producto, Cliente)', valorFactura:'Valor Factura', canal:'Canal', observacion:'Observación (ATJ)', estado:'Estado', centroCostos:'Centro de Costos', ordenCompra:'Orden de Compra', nitProveedor:'NIT Proveedor', nomProveedor:'Nom. Proveedor', docCargado:'Doc. Cargado', obsKatherine:'Obs. Katherine', notas:'notas' },
  planes:       { id:'id', distribuidor:'distribuidor', anio:'año', quarter:'quarter', estado:'estado', tiposPlan:'tipoPlan', metaGalones:'metaGalones', metaVenta:'metaVenta', condiciones:'condiciones', acuerdos:'acuerdos', notas:'notas', metaMes1:'metaMes1', metaMes2:'metaMes2', metaMes3:'metaMes3', metaGalonesMes1:'metaGalonesMes1', metaGalonesMes2:'metaGalonesMes2', metaGalonesMes3:'metaGalonesMes3', tieneCVC:'tieneCVC', montoCVC:'montoCVC', detalleCVC:'detalleCVC', tienePromotora:'tienePromotora', nombrePromotora:'nombrePromotora', montoPromotora:'montoPromotora', condicionesGenerales:'condicionesGenerales' },
  apoyoCierre:  { id:'ID', anio:'AÑO', mes:'MES', comercial:'COMERCIAL', monto:'MONTO_ASIGNADO', distribuidor:'CLIENTE' },
  redenciones:  { id:'ID', fecha:'FECHA', anio:'AÑO', mes:'MES', comercial:'COMERCIAL', cliente:'CLIENTE', producto:'PRODUCTO', valor:'VALOR_REDIMIDO', notas:'OBSERVACION' },
  pendientes:   { id:'id', distribuidor:'distribuidor', tarea:'tarea', categoria:'categoria', fechaLimite:'fechaLimite', prioridad:'prioridad', estado:'estado', responsable:'responsable', notas:'notas', asignadoPor:'asignadoPor', fechaCreacion:'fechaCreacion' },
  tareasAsignadas: {
    id:'id',
    nombre:'nombre',
    descripcion:'descripcion',
    responsable:'responsable',
    areas:'areas',
    fechaLimite:'fechaLimite',
    estado:'estado',
    asignadoPor:'asignadoPor',
    fechaCreacion:'fechaCreacion',
    fechaFinalizacion:'fechaFinalizacion',
    tiempoValor:'tiempoValor',
    tiempoUnidad:'tiempoUnidad',
    tiempoMinutos:'tiempoMinutos'
  },
  misTareas: {
    id:'id',
    nombre:'nombre',
    descripcion:'descripcion',
    responsable:'responsable',
    areas:'areas',
    fechaLimite:'fechaLimite',
    estado:'estado',
    asignadoPor:'asignadoPor',
    fechaCreacion:'fechaCreacion',
    fechaFinalizacion:'fechaFinalizacion',
    tiempoValor:'tiempoValor',
    tiempoUnidad:'tiempoUnidad',
    tiempoMinutos:'tiempoMinutos'
  },
  proyectos:    { id:'id', nombre:'nombre', descripcion:'descripcion', responsable:'responsable', areas:'areas', fechaEntrega:'fechaEntrega', estado:'estado', creadoPor:'creadoPor', subtareas:'subtareas', fechaCreacion:'fechaCreacion', correoFinalEnviado:'correoFinalEnviado' },
  subtareasProyectos: { id:'id', proyecto:'proyecto', nombre:'nombre', descripcion:'descripcion', responsable:'responsable', fechaEntrega:'fechaEntrega', estado:'estado', areas:'areas', creadoPor:'creadoPor', fechaCreacion:'fechaCreacion' },
}

// ── NORMALIZAR MES ── "MAYO" / "mayo" → "Mayo"
const normMes = m => {
  const s = String(m||'').trim()
  if(!s) return s
  return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase()
}

function toSheetRow(tipo, obj) {
  const map = CAMPOS_MAP[tipo] || {}
  const row = {}
  Object.entries(map).forEach(([appKey, sheetKey]) => {
    let val = obj[appKey]
    if(val === undefined) val = ''
    if(Array.isArray(val)) {
      // Arrays de strings → "a, b, c"
      // Arrays de objetos → JSON string
      if(val.length > 0 && typeof val[0] === 'object') {
        val = JSON.stringify(val)
      } else {
        val = val.join(', ')
      }
    }
    row[sheetKey] = val
  })
  // Para planes: serializar planesItems como texto legible en columna existente
  if(tipo === 'planes' && obj.planesItems && obj.planesItems.length > 0) {
    const resumen = obj.planesItems
      .filter(p=>p.nombre)
      .map(p=>{
        let s = p.nombre
        if(p.tienePago==='Sí' && p.descripcionPago) s += ': '+p.descripcionPago
        if(p.condiciones) s += ' ['+p.condiciones+']'
        return s
      }).join(' | ')
    row['condiciones'] = resumen || row['condiciones']
  }
  return row
}

// ── FIX: fromSheetRow normaliza el mes al leer de Sheets ──
function fromSheetRow(tipo, row) {
  const map = CAMPOS_MAP[tipo] || {}
  const obj = {}
  // Índice case-insensitive para encontrar columnas sin importar capitalización
  const rowCI = {}
  Object.keys(row).forEach(k => { rowCI[k.toLowerCase()] = row[k] })

  Object.entries(map).forEach(([appKey, sheetKey]) => {
    let val = row[sheetKey]
    if(val === undefined) val = rowCI[sheetKey.toLowerCase()]
    // Manejo especial para 'año' / 'ano' / 'Año' / 'year'
    if(val === undefined && appKey === 'anio') {
      val = row['año'] ?? row['Año'] ?? row['ANO'] ?? row['ano'] ?? row['year'] ?? rowCI['año'] ?? rowCI['ano']
    }
    if(val === undefined || val === '') val = appKey === 'anio' ? 2026 : ''
    // Normalizar mes: "MAYO" / "mayo" → "Mayo"
    if(appKey === 'mes' && typeof val === 'string' && val.trim()) {
      val = normMes(val)
    }
    if(['anio','inversion','galonesPlan','galones','ventaNeta','monto','metaGalones','metaVenta','valor','tiempoValor','tiempoMinutos'].includes(appKey)) {
      val = Number(val) || 0
    }
    if(appKey === 'tiposPlan' && typeof val === 'string') {
      val = val.split(',').map(s=>s.trim()).filter(Boolean)
    }
    obj[appKey] = val
  })
  // Para planes: restaurar planesItems desde JSON si existe en la fila
  if(tipo === 'planes') {
    try {
      const pi = row['planesItems'] || rowCI['planesitems']
      if(pi && typeof pi === 'string' && pi.startsWith('[')) {
        obj.planesItems = JSON.parse(pi)
      }
    } catch(e) {}
    // Asegurar valores numéricos en metas
    ;['metaMes1','metaMes2','metaMes3','metaGalonesMes1','metaGalonesMes2','metaGalonesMes3','montoCVC','montoPromotora'].forEach(k=>{
      if(obj[k]!==undefined) obj[k] = obj[k]===''?'':Number(obj[k])||0
    })
  }
  if(tipo === 'proyectos') {
    if(typeof obj.areas === 'string') {
      obj.areas = obj.areas.split(',').map(a=>a.trim()).filter(Boolean)
    }
    try {
      if(typeof obj.subtareas === 'string' && obj.subtareas.trim()) {
        obj.subtareas = JSON.parse(obj.subtareas)
      }
    } catch(e) {
      obj.subtareas = []
    }
    if(!Array.isArray(obj.subtareas)) obj.subtareas = []
  }
  if(!obj.id) obj.id = Date.now() + Math.random()
  return obj
}

async function cargarDesdeSheets(uid) {
  const cfg = SHEETS_CONFIG[uid]
  if(!cfg?.enabled) return null
  try {
    const readUrl = cfg.spreadsheetId
      ? `${cfg.url}?spreadsheetId=${encodeURIComponent(cfg.spreadsheetId)}`
      : cfg.url
    const res = await fetch(readUrl)
    const json = await res.json()
    if(!json.ok) return null
    const d = json.data
    return {
      inversiones:      (d.inversiones      ||[]).map(r=>fromSheetRow('inversiones',r)),
      ventas:           (d.ventas           ||[]).map(r=>fromSheetRow('ventas',r)),
      presupuestos:     (d.presupuestos || d.presupuesto || d.PRESUPUESTO || []).map(r=>fromSheetRow('presupuestos',r)),
      planes:           (d.planes           ||[]).map(r=>fromSheetRow('planes',r)),
      apoyoCierre:      (d.apoyo_cierre     ||[]).map(r=>fromSheetRow('apoyoCierre',r)),
      redenciones:      (d.redenciones_apoyo||[]).map(r=>fromSheetRow('redenciones',r)),
      gastosPresupuesto:(d.gastos_presupuesto||[]).map(r=>fromSheetRow('gastosPresupuesto',r)),
      pendientes:       (d.pendientes       ||[]).map(r=>fromSheetRow('pendientes',r)),
      tareasAsignadas:  (d.tareas_asignadas || d.TAREAS_ASIGNADAS || d.tareasAsignadas || []).map(r=>fromSheetRow('tareasAsignadas',r)),
      misTareas:         (d.mis_tareas || d.MIS_TAREAS || d.misTareas || []).map(r=>fromSheetRow('misTareas',r)),
      proyectos:         (d.proyectos || d.PROYECTOS || []).map(r=>fromSheetRow('proyectos',r)),
    }
  } catch(e) {
    console.error('Error cargando Sheets:', e)
    return null
  }
}

async function insertarFilaSheets(uid, tipo, item) {
  const cfg = SHEETS_CONFIG[uid]
  if(!cfg?.enabled) return
  const hoja = HOJA_MAP[tipo]
  if(!hoja) return
  const fila = toSheetRow(tipo, item)
  const cb = 'gs_cb_' + Date.now() + '_' + Math.random().toString(36).slice(2)
  return new Promise(resolve => {
    window[cb] = () => { resolve(); delete window[cb] }
    const s = document.createElement('script')
    s.src = cfg.url + '?callback=' + cb + '&data=' + encodeURIComponent(JSON.stringify({
      accion:'agregar_lote',
      hoja,
      filas:[fila],
      spreadsheetId: cfg.spreadsheetId || ''
    }))
    s.onerror = () => resolve()
    s.onload = () => { setTimeout(resolve, 200); s.remove() }
    document.head.appendChild(s)
    setTimeout(resolve, 4000)
  })
}

async function eliminarFilaSheets(uid, tipo, id) {
  const cfg = SHEETS_CONFIG[uid]
  if(!cfg?.enabled) return
  const hoja = HOJA_MAP[tipo]
  if(!hoja) return
  const cb = 'gs_cb_' + Date.now() + '_' + Math.random().toString(36).slice(2)
  return new Promise(resolve => {
    window[cb] = () => { resolve(); delete window[cb] }
    const s = document.createElement('script')
    s.src = cfg.url + '?callback=' + cb + '&data=' + encodeURIComponent(JSON.stringify({
      accion:'eliminar',
      hoja,
      id,
      spreadsheetId: cfg.spreadsheetId || ''
    }))
    s.onerror = () => resolve()
    s.onload = () => { setTimeout(resolve, 200); s.remove() }
    document.head.appendChild(s)
    setTimeout(resolve, 4000)
  })
}

async function guardarEnSheets(uid, tipo, items) {
  const cfg = SHEETS_CONFIG[uid]
  if(!cfg?.enabled) return
  const hoja = HOJA_MAP[tipo]
  if(!hoja) return
  const filas = items.map(i=>toSheetRow(tipo, i))
  const enviarScript = (payload) => new Promise(resolve => {
    const cb = 'gs_cb_' + Date.now() + '_' + Math.random().toString(36).slice(2)
    window[cb] = () => { resolve(); delete window[cb] }
    const s = document.createElement('script')
    s.src = cfg.url + '?callback=' + cb + '&data=' + encodeURIComponent(JSON.stringify(payload))
    s.onerror = () => { resolve() }
    s.onload = () => { setTimeout(resolve, 200); s.remove() }
    document.head.appendChild(s)
    setTimeout(resolve, 3000)
  })
  try {
    await enviarScript({ accion:'limpiar', hoja })
    await new Promise(r=>setTimeout(r,800))
    for(let i=0; i<filas.length; i++) {
      await enviarScript({ accion:'agregar_lote', hoja, filas:[filas[i]] })
      await new Promise(r=>setTimeout(r,300))
    }
  } catch(e) {
    console.error('Error guardando en Sheets:', e)
  }
}

const MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
const TIPOS_PLAN = ['Prolub respalda','Sell out Prolub respalda','Apoyo directo','Activación','Otro']
const CONCEPTOS = ['Apoyo a la nomina','CVC','Producto/promocion','Evento','Material POP','Digital','Transporte','Otro']
const ESTADOS_PEND = ['Pendiente','En curso','Listo','Cancelado']
const PRIORIDADES = ['Alta','Media','Baja']
const QUARTERS = ['Q1','Q2','Q3','Q4']
const ESTADOS_PLAN = ['Activo','En negociación','Cerrado','Cancelado']
const COLORES = ['#6c63ff','#3dd68c','#ff9f43','#ff5f5f','#06b6d4','#a78bfa','#f59e0b','#10b981','#ec4899','#14b8a6']

function load() {
  try { const d = localStorage.getItem(_currentUserKey); if (d) return JSON.parse(d) } catch {}
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
    tareasAsignadas: [],
    misTareas: [],
    proyectos: [],
    pendientes: [
      { id:1, distribuidor:'CVS- SERVITECAS S.A.S', tarea:'Revisar cumplimiento meta Q1', categoria:'Seguimiento', fechaLimite:'2026-03-31', prioridad:'Alta', estado:'Pendiente', responsable:'', notas:'' },
      { id:2, distribuidor:'MAQUINAGRO S.A.S', tarea:'Cerrar negociación plan Q2', categoria:'Negociación', fechaLimite:'2026-04-15', prioridad:'Alta', estado:'En curso', responsable:'', notas:'' },
    ],
  }
}
function save(d, opts={}) { 
  try { localStorage.setItem(_currentUserKey, JSON.stringify(d)) } catch {} 
  if(window._sheetsSyncFn) try { window._sheetsSyncFn(d, opts) } catch {}
}

const parseN = v => { if(typeof v==='number') return Math.abs(v); const c=String(v).replace(/[$\s ]/g,'').replace(/\./g,'').replace(',','.'); return parseFloat(c)||0 }
const cop = n => new Intl.NumberFormat('es-CO',{style:'currency',currency:'COP',minimumFractionDigits:0}).format(n||0)
const num = n => new Intl.NumberFormat('es-CO').format(n||0)

const S = {
  card: { background:'var(--bg2)', border:'1px solid var(--border)', borderRadius:'var(--radius)', overflow:'hidden', boxShadow:'0 2px 12px rgba(30,16,64,0.06)' },
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

const KPI_GRADIENTS = [
  'linear-gradient(135deg,#6c3fc4,#a855f7)',
  'linear-gradient(135deg,#0891b2,#22d3ee)',
  'linear-gradient(135deg,#16a25a,#4ade80)',
  'linear-gradient(135deg,#d97706,#fbbf24)',
  'linear-gradient(135deg,#db2777,#f472b6)',
  'linear-gradient(135deg,#e03c3c,#f87171)',
]
let _kpiIdx = 0

function KpiCard({ icon:Icon, label, value, sub, accent, gradient }) {
  const bg = gradient || (accent ? null : KPI_GRADIENTS[0])
  const useGrad = !accent
  return (
    <div style={{
      background: useGrad ? bg : 'var(--bg2)',
      border: useGrad ? 'none' : '1px solid var(--border)',
      borderRadius:'var(--radius)',
      padding:'20px 22px',
      display:'flex',flexDirection:'column',gap:8,
      boxShadow: useGrad ? '0 4px 20px rgba(0,0,0,0.12)' : 'none',
      position:'relative', overflow:'visible', minWidth:0
    }}>
      {useGrad && <div style={{position:'absolute',top:-20,right:-20,width:80,height:80,borderRadius:'50%',background:'rgba(255,255,255,0.08)'}}/>}
      <div style={{display:'flex',alignItems:'center',justifyContent:'space-between'}}>
        <div style={{fontSize:11,fontWeight:600,color:useGrad?'rgba(255,255,255,0.75)':'var(--text2)',textTransform:'uppercase',letterSpacing:'0.06em'}}>{label}</div>
        <div style={{width:32,height:32,borderRadius:8,background:useGrad?'rgba(255,255,255,0.15)':'var(--accent-soft)',display:'flex',alignItems:'center',justifyContent:'center'}}>
          <Icon size={15} color={useGrad?'#fff':accent||'var(--accent)'}/>
        </div>
      </div>
      <div style={{fontSize:22,fontWeight:700,color:useGrad?'#fff':(accent||'var(--text)'),fontFamily:'var(--mono)',letterSpacing:'-0.02em',lineHeight:1.1,wordBreak:'break-all',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{value}</div>
      {sub&&<div style={{fontSize:11,color:useGrad?'rgba(255,255,255,0.65)':'var(--text3)'}}>{sub}</div>}
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
  const todosDistribuidores = [...new Set([...data.inversiones.map(i=>i.distribuidor),...data.ventas.map(v=>v.distribuidor)])].filter(Boolean).sort()

  // ── Todos los tipos que realmente existen en los datos ──
  const todosTipos = [...new Set(data.inversiones.map(i=>i.tipoPlan).filter(Boolean))].sort()

  const invF = data.inversiones.filter(i=>
    (!filtroMes||normMes(i.mes)===normMes(filtroMes))&&
    (!filtroAnio||i.anio===Number(filtroAnio))&&
    (!filtroDist||i.distribuidor===filtroDist)
  )
  const ventF = data.ventas.filter(v=>
    (!filtroMes||normMes(v.mes)===normMes(filtroMes))&&
    (!filtroAnio||v.anio===Number(filtroAnio))&&
    (!filtroDist||v.distribuidor===filtroDist)
  )
  const presF = data.presupuestos.filter(p=>
    (!filtroMes||normMes(p.mes)===normMes(filtroMes))&&
    (!filtroAnio||p.anio===Number(filtroAnio))
  )

  const totalInv   = invF.reduce((s,i)=>s+(i.inversion||0),0)
  const totalVenta = ventF.reduce((s,v)=>s+(v.ventaNeta||0),0)
  const totalGalones = ventF.reduce((s,v)=>s+(v.galones||0),0)
  const totalPres  = presF.reduce((s,p)=>s+(p.monto||0),0)
  const roiPct     = totalVenta>0 ? (totalInv/totalVenta)*100 : 0

  // Tabla resumen por distribuidor
  const distLista = filtroDist
    ? [filtroDist]
    : [...new Set([...invF.map(i=>i.distribuidor),...ventF.map(v=>v.distribuidor)])].filter(Boolean)

  const porDist = distLista.map(d=>{
    const inv    = invF.filter(i=>i.distribuidor===d).reduce((s,i)=>s+(i.inversion||0),0)
    const venta  = ventF.filter(v=>v.distribuidor===d).reduce((s,v)=>s+(v.ventaNeta||0),0)
    const galones= ventF.filter(v=>v.distribuidor===d).reduce((s,v)=>s+(v.galones||0),0)
    return { name:d, inv, venta, galones, peso: venta>0?(inv/venta)*100:0 }
  }).sort((a,b)=>b.inv-a.inv)

  // Gráfica por mes
  const porMes = MESES.map(m=>({
    name:m.slice(0,3),
    invertido: data.inversiones.filter(i=>normMes(i.mes)===m&&(!filtroAnio||i.anio===Number(filtroAnio))&&(!filtroDist||i.distribuidor===filtroDist)).reduce((s,i)=>s+(i.inversion||0),0),
    venta:     data.ventas.filter(v=>normMes(v.mes)===m&&(!filtroAnio||v.anio===Number(filtroAnio))&&(!filtroDist||v.distribuidor===filtroDist)).reduce((s,v)=>s+(v.ventaNeta||0),0),
    presupuesto: data.presupuestos.filter(p=>normMes(p.mes)===m&&(!filtroAnio||p.anio===Number(filtroAnio))).reduce((s,p)=>s+(p.monto||0),0),
  })).filter(m=>m.invertido>0||m.venta>0||m.presupuesto>0)

  // ── PieChart: TODOS los tipos reales en los datos filtrados ──
  const tiposEnFiltro = [...new Set(invF.map(i=>i.tipoPlan).filter(Boolean))].sort()
  const porTipo = tiposEnFiltro.map((t,i)=>({
    name: t,
    value: invF.filter(i2=>i2.tipoPlan===t).reduce((s,i2)=>s+(i2.inversion||0),0),
    color: COLORES[i%COLORES.length],
  })).filter(t=>t.value>0).sort((a,b)=>b.value-a.value)

  // ── Navegación cliente por cliente ──
  const distribuidoresOrdenados = [...new Set(invF.map(i=>i.distribuidor))].filter(Boolean).sort()
  const idxActual = filtroDist ? distribuidoresOrdenados.indexOf(filtroDist) : -1
  const irAnterior = () => {
    if(idxActual<=0) setFiltroDist(distribuidoresOrdenados[distribuidoresOrdenados.length-1])
    else setFiltroDist(distribuidoresOrdenados[idxActual-1])
  }
  const irSiguiente = () => {
    if(idxActual>=distribuidoresOrdenados.length-1) setFiltroDist(distribuidoresOrdenados[0])
    else setFiltroDist(distribuidoresOrdenados[idxActual+1])
  }

  // ── Desglose del distribuidor seleccionado (barra horizontal multicolor) ──
  // Muestra todos los tipos reales de ese distribuidor
  const tiposDelDist = filtroDist
    ? [...new Set(invF.filter(i=>i.distribuidor===filtroDist).map(i=>i.tipoPlan).filter(Boolean))].sort()
    : tiposEnFiltro
  const invDistTotal = invF.filter(i=>!filtroDist||i.distribuidor===filtroDist).reduce((s,i)=>s+(i.inversion||0),0)

  // Datos para el barchart de desglose (un registro por distribuidor con campos por tipo)
  const distribuidoresDesglose = filtroDist
    ? [filtroDist]
    : [...new Set(invF.map(i=>i.distribuidor))].filter(Boolean).sort()

  const datosDesglose = distribuidoresDesglose.map(d=>({
    name: d.length>18 ? d.slice(0,18)+'…' : d,
    fullName: d,
    ...Object.fromEntries(todosTipos.map(t=>[t,
      invF.filter(i=>i.distribuidor===d&&i.tipoPlan===t).reduce((s,i)=>s+(i.inversion||0),0)
    ]))
  }))

  return (
    <div style={{display:'flex',flexDirection:'column',gap:20}}>
      {/* ── Filtros con navegación cliente a cliente ── */}
      <div style={{display:'flex',gap:10,alignItems:'center',flexWrap:'wrap'}}>
        <select value={filtroAnio} onChange={e=>setFiltroAnio(e.target.value)} style={{width:110}}>
          <option value="">Todos los años</option>
          {anios.map(a=><option key={a}>{a}</option>)}
        </select>
        <select value={filtroMes} onChange={e=>setFiltroMes(e.target.value)} style={{width:150}}>
          <option value="">Todos los meses</option>
          {MESES.map(m=><option key={m}>{m}</option>)}
        </select>
        <select value={filtroDist} onChange={e=>setFiltroDist(e.target.value)} style={{width:240}}>
          <option value="">Todos los distribuidores</option>
          {todosDistribuidores.map(d=><option key={d}>{d}</option>)}
        </select>

        {/* Navegación ← → cliente a cliente */}
        {filtroDist&&(
          <div style={{display:'flex',gap:4,alignItems:'center'}}>
            <button onClick={irAnterior}
              style={{...S.btn('var(--bg3)','var(--text2)'),padding:'5px 10px',fontSize:13,fontWeight:700}}>‹</button>
            <span style={{fontSize:11,color:'var(--text3)',minWidth:60,textAlign:'center'}}>
              {idxActual+1}/{distribuidoresOrdenados.length}
            </span>
            <button onClick={irSiguiente}
              style={{...S.btn('var(--bg3)','var(--text2)'),padding:'5px 10px',fontSize:13,fontWeight:700}}>›</button>
          </div>
        )}

        {(filtroMes||filtroAnio!=='2026'||filtroDist)&&(
          <button onClick={()=>{setFiltroMes('');setFiltroAnio('2026');setFiltroDist('')}}
            style={{...S.btn('var(--bg3)','var(--text2)'),fontSize:12,padding:'5px 12px'}}>
            Limpiar filtros
          </button>
        )}
        {filtroDist&&(
          <span style={{fontSize:12,background:'var(--accent-soft)',color:'var(--accent2)',padding:'4px 12px',borderRadius:20,border:'1px solid rgba(108,99,255,0.2)'}}>
            📍 {filtroDist}
          </span>
        )}
      </div>

      {/* KPIs */}
      <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(175px,1fr))',gap:13}}>
        <KpiCard icon={TrendingUp} label="Total invertido"    value={cop(totalInv)}    sub={`${invF.length} registros`} accent="var(--accent2)"/>
        <KpiCard icon={DollarSign} label="Presupuesto"        value={cop(totalPres)}   sub={totalPres>0?`${((totalInv/totalPres)*100).toFixed(1)}% ejecutado`:'Sin presupuesto'}/>
        <KpiCard icon={ShoppingCart} label="Venta neta"       value={cop(totalVenta)}  sub={`${num(totalGalones)} galones`} accent="var(--green)"/>
        <KpiCard icon={BarChart2}  label="Inversión / Venta"  value={`${roiPct.toFixed(1)}%`} sub="Peso inversión sobre venta" accent={roiPct>15?'var(--red)':roiPct>10?'var(--yellow)':'var(--green)'}/>
        {filtroDist&&totalGalones>0&&<KpiCard icon={DollarSign} label="Precio prom/galón" value={cop(totalVenta/totalGalones)} accent="var(--orange)"/>}
      </div>

      {/* Gráficas fila 1 */}
      <div style={{display:'grid',gridTemplateColumns:'1.3fr 1fr',gap:16}}>
        {/* Inversión vs Venta por mes */}
        <div style={{...S.card,padding:20}}>
          <h4 style={{fontSize:11,fontWeight:600,color:'var(--text2)',textTransform:'uppercase',letterSpacing:'0.05em',marginBottom:14}}>
            {filtroDist ? `Inversión por mes — ${filtroDist.split(' ')[0]}` : 'Inversión vs Venta Neta por mes'}
          </h4>
          <ResponsiveContainer width="100%" height={200}>
            <BarChart data={porMes} barGap={3}>
              <XAxis dataKey="name" tick={{fill:'var(--text3)',fontSize:11}} axisLine={false} tickLine={false}/>
              <YAxis tickFormatter={v=>`${(v/1000000).toFixed(0)}M`} tick={{fill:'var(--text3)',fontSize:10}} axisLine={false} tickLine={false}/>
              <Tooltip content={<CT/>}/>
              <Bar dataKey="venta"       name="Venta neta"   fill="var(--green)"  radius={[4,4,0,0]} opacity={0.7}/>
              <Bar dataKey="invertido"   name="Inversión"    fill="var(--accent)" radius={[4,4,0,0]}/>
              <Bar dataKey="presupuesto" name="Presupuesto"  fill="var(--bg4)"    radius={[4,4,0,0]}/>
            </BarChart>
          </ResponsiveContainer>
        </div>

        {/* PieChart TODOS los tipos reales */}
        <div style={{...S.card,padding:20}}>
          <h4 style={{fontSize:11,fontWeight:600,color:'var(--text2)',textTransform:'uppercase',letterSpacing:'0.05em',marginBottom:14}}>
            {filtroDist ? `Composición — ${filtroDist.split(' ')[0]}` : 'Inversión por tipo de plan'}
          </h4>
          {porTipo.length===0
            ? <div style={{height:200,display:'flex',alignItems:'center',justifyContent:'center',color:'var(--text3)',fontSize:13}}>Sin inversiones en el período</div>
            : <ResponsiveContainer width="100%" height={200}>
                <PieChart>
                  <Pie data={porTipo} cx="50%" cy="50%" innerRadius={45} outerRadius={75} dataKey="value" nameKey="name" paddingAngle={3}>
                    {porTipo.map((t,i)=><Cell key={i} fill={t.color}/>)}
                  </Pie>
                  <Tooltip formatter={v=>cop(v)} contentStyle={{background:'var(--bg3)',border:'1px solid var(--border2)',borderRadius:10,fontSize:11}}/>
                  <Legend iconSize={7} iconType="circle" wrapperStyle={{fontSize:10,color:'var(--text2)'}}/>
                </PieChart>
              </ResponsiveContainer>
          }
        </div>
      </div>

      {/* Tabla distribuidor vs venta */}
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
                <td style={S.td}>
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
        {!filtroDist&&<div style={{padding:'8px 18px',fontSize:11,color:'var(--text3)',borderTop:'1px solid var(--border)'}}>
          💡 Clic en una fila para ver el detalle del distribuidor · usa ‹ › para navegar cliente a cliente
        </div>}
      </div>

      {/* ── DESGLOSE: barra horizontal multicolor por tipo de inversión ── */}
      {invF.length>0 && (
        <div style={S.card}>
          {/* Header con navegación si hay distribuidor seleccionado */}
          <div style={{padding:'14px 18px',borderBottom:'1px solid var(--border)',display:'flex',alignItems:'center',justifyContent:'space-between',flexWrap:'wrap',gap:8}}>
            <div>
              <h4 style={{fontSize:11,fontWeight:600,color:'var(--text2)',textTransform:'uppercase',letterSpacing:'0.05em',marginBottom:filtroDist?3:0}}>
                {filtroDist
                  ? `Desglose por tipo de inversión — ${filtroDist}`
                  : `Desglose por distribuidor y tipo de inversión${filtroMes?' — '+filtroMes:''}`
                }
              </h4>
              {filtroDist&&(
                <div style={{display:'flex',alignItems:'center',gap:10}}>
                  <button onClick={irAnterior} style={{...S.btn('var(--bg4)','var(--text2)'),padding:'3px 10px',fontSize:12,fontWeight:700}}>‹ Anterior</button>
                  <span style={{fontSize:11,color:'var(--text3)'}}>{idxActual+1} de {distribuidoresOrdenados.length}</span>
                  <button onClick={irSiguiente} style={{...S.btn('var(--bg4)','var(--text2)'),padding:'3px 10px',fontSize:12,fontWeight:700}}>Siguiente ›</button>
                  <button onClick={()=>setFiltroDist('')} style={{...S.btn('var(--red-soft)','var(--red)'),padding:'3px 10px',fontSize:11}}>✕ Ver todos</button>
                </div>
              )}
            </div>
            {!filtroDist&&<span style={{fontSize:11,color:'var(--text3)'}}>Clic en fila para ver desglose individual</span>}
          </div>

          {/* Barra de participación por tipo (cuando hay un distribuidor seleccionado) */}
          {filtroDist&&invDistTotal>0&&(
            <div style={{padding:'16px 20px',borderBottom:'1px solid var(--border)'}}>
              {/* Barra multicolor proporcional */}
              <div style={{display:'flex',height:28,borderRadius:8,overflow:'hidden',marginBottom:12}}>
                {todosTipos.map((t,i)=>{
                  const v=invF.filter(inv=>inv.distribuidor===filtroDist&&inv.tipoPlan===t).reduce((s,inv)=>s+(inv.inversion||0),0)
                  if(v===0) return null
                  const pct=(v/invDistTotal)*100
                  return (
                    <div key={t} style={{width:pct+'%',background:COLORES[i%COLORES.length],position:'relative',display:'flex',alignItems:'center',justifyContent:'center',minWidth:pct>8?'auto':0}}
                      title={`${t}: ${cop(v)} (${pct.toFixed(1)}%)`}>
                      {pct>10&&<span style={{fontSize:10,fontWeight:600,color:'#fff',textShadow:'0 1px 2px rgba(0,0,0,0.5)',padding:'0 4px',whiteSpace:'nowrap',overflow:'hidden'}}>{pct.toFixed(0)}%</span>}
                    </div>
                  )
                })}
              </div>
              {/* Leyenda */}
              <div style={{display:'flex',flexWrap:'wrap',gap:10}}>
                {todosTipos.map((t,i)=>{
                  const v=invF.filter(inv=>inv.distribuidor===filtroDist&&inv.tipoPlan===t).reduce((s,inv)=>s+(inv.inversion||0),0)
                  if(v===0) return null
                  return (
                    <div key={t} style={{display:'flex',alignItems:'center',gap:5}}>
                      <div style={{width:10,height:10,borderRadius:3,background:COLORES[i%COLORES.length],flexShrink:0}}/>
                      <span style={{fontSize:11,color:'var(--text2)'}}>{t}</span>
                      <span style={{fontSize:11,fontFamily:'var(--mono)',color:COLORES[i%COLORES.length],fontWeight:600}}>{cop(v)}</span>
                    </div>
                  )
                })}
              </div>
            </div>
          )}

          {/* BarChart apilado horizontal (todos los distribuidores o el seleccionado) */}
          {datosDesglose.length>0&&(
            <div style={{padding:'16px 20px 0'}}>
              <ResponsiveContainer width="100%" height={Math.max(80, datosDesglose.length*46)}>
                <BarChart data={datosDesglose} layout="vertical" margin={{left:10,right:20,top:0,bottom:0}}>
                  <XAxis type="number" tickFormatter={v=>`${(v/1000000).toFixed(1)}M`} tick={{fill:'var(--text3)',fontSize:10}} axisLine={false} tickLine={false}/>
                  <YAxis type="category" dataKey="name" tick={{fill:'var(--text2)',fontSize:11}} axisLine={false} tickLine={false} width={150}/>
                  <Tooltip
                    formatter={(v,name)=>v>0?[cop(v),name]:null}
                    contentStyle={{background:'var(--bg3)',border:'1px solid var(--border2)',borderRadius:10,fontSize:11}}
                  />
                  <Legend iconSize={7} iconType="circle" wrapperStyle={{fontSize:10,color:'var(--text2)'}}/>
                  {todosTipos.map((t,i)=>(
                    <Bar key={t} dataKey={t} name={t} stackId="a" fill={COLORES[i%COLORES.length]}
                      radius={i===todosTipos.length-1?[0,4,4,0]:[0,0,0,0]}/>
                  ))}
                </BarChart>
              </ResponsiveContainer>
            </div>
          )}

          {/* Tabla detalle con TODOS los tipos reales */}
          <div style={{overflowX:'auto'}}>
            <table style={{width:'100%',borderCollapse:'collapse',minWidth:600}}>
              <thead>
                <tr style={{background:'var(--bg3)'}}>
                  <th style={S.th}>Distribuidor</th>
                  {todosTipos.map((t,i)=>(
                    <th key={t} style={{...S.th,fontSize:10,color:COLORES[i%COLORES.length]}}>
                      {t.length>18?t.slice(0,18)+'…':t}
                    </th>
                  ))}
                  <th style={S.th}>Total</th>
                </tr>
              </thead>
              <tbody>
                {distribuidoresDesglose.map((d,idx)=>{
                  const porTipoRow = todosTipos.map(t=>
                    invF.filter(i=>i.distribuidor===d&&i.tipoPlan===t).reduce((s,i)=>s+(i.inversion||0),0)
                  )
                  const totalRow = porTipoRow.reduce((s,v)=>s+v,0)
                  if(totalRow===0) return null
                  return (
                    <tr key={idx}
                      style={{cursor:filtroDist?'default':'pointer',background:filtroDist===d?'rgba(108,99,255,0.07)':'transparent'}}
                      onClick={()=>{if(!filtroDist) setFiltroDist(d)}}
                    >
                      <td style={{...S.td,fontWeight:500,color:filtroDist===d?'var(--accent2)':'var(--text)'}}>{d}</td>
                      {porTipoRow.map((v,i)=>(
                        <td key={i} style={{...S.td,fontFamily:'var(--mono)',fontSize:12,color:v>0?COLORES[i%COLORES.length]:'var(--text3)'}}>
                          {v>0?cop(v):'—'}
                        </td>
                      ))}
                      <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700,color:'var(--accent2)'}}>{cop(totalRow)}</td>
                    </tr>
                  )
                })}
                <tr style={{borderTop:'2px solid var(--border2)',background:'var(--bg3)'}}>
                  <td style={{...S.td,fontWeight:700}}>TOTAL</td>
                  {todosTipos.map((t,i)=>{
                    const v=invF.filter(i2=>i2.tipoPlan===t).reduce((s,i2)=>s+(i2.inversion||0),0)
                    return <td key={i} style={{...S.td,fontFamily:'var(--mono)',fontWeight:600,color:v>0?COLORES[i%COLORES.length]:'var(--text3)'}}>{v>0?cop(v):'—'}</td>
                  })}
                  <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700,color:'var(--accent2)'}}>{cop(totalInv)}</td>
                </tr>
              </tbody>
            </table>
          </div>
          <div style={{padding:'8px 18px',fontSize:11,color:'var(--text3)',borderTop:'1px solid var(--border)'}}>
            {filtroDist
              ? `💡 ${filtroDist} · ${idxActual+1}/${distribuidoresOrdenados.length} · Usa ‹ › para navegar entre clientes`
              : '💡 Clic en una fila para ver el desglose individual del distribuidor'
            }
          </div>
        </div>
      )}
    </div>
  )
}


// ═══════════════════════════════════════════════════════
// INVERSIONES — Tabla tipo Excel con fixes
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

  // FIX: usa normMes para comparar meses independiente de capitalización
  const lista = data.inversiones.filter(i=>
    (!filtroMes||normMes(i.mes)===normMes(filtroMes)) &&
    (!filtroAnio||i.anio===Number(filtroAnio)) &&
    (!filtroDist||i.distribuidor===filtroDist) &&
    (!busqueda||Object.values(i).some(v=>String(v||'').toLowerCase().includes(busqueda.toLowerCase())))
  ).sort((a,b)=>{
    const mi=MESES.indexOf(normMes(a.mes)), mj=MESES.indexOf(normMes(b.mes))
    return mi!==mj ? mi-mj : String(a.distribuidor||'').localeCompare(String(b.distribuidor||''))
  })

  const totalFiltrado = lista.reduce((s,i)=>s+(Number(i.inversion)||0),0)

  // FIX: presMes usa normMes para encontrar el presupuesto aunque venga en mayúsculas de Sheets
  const presMes = filtroMes
    ? data.presupuestos.filter(p=>normMes(p.mes)===normMes(filtroMes)&&(!filtroAnio||p.anio===Number(filtroAnio))).reduce((s,p)=>s+(Number(p.monto)||0),0)
    : 0

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

  // FIX: commitEdit ahora sincroniza la fila editada a Sheets
  const commitEdit = () => {
    if(!editCell) return
    const {id,field}=editCell
    const inversiones=data.inversiones.map(i=>{
      if(i.id!==id) return i
      const val=(field==='inversion'||field==='galonesPlan'||field==='anio')?(Number(editVal)||0):editVal
      return {...i,[field]:val}
    })
    const nd={...data,inversiones}
    setData(nd)
    save(nd)
    // Sincronizar fila editada a Sheets: eliminar la vieja e insertar la actualizada
    const filaActualizada = inversiones.find(i=>i.id===id)
    if(filaActualizada && window._sheetsSyncFn) {
      window._sheetsSyncFn(nd, {eliminar:id, tipo:'inversiones'})
      setTimeout(()=>{
        window._sheetsSyncFn(nd, {insertar:filaActualizada, tipo:'inversiones'})
      }, 600)
    }
    setEditCell(null)
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
      id:Date.now()+i, fecha:r[0]||'', anio:Number(r[1])||2026, mes:normMes(r[2]||''),
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
                <td style={{padding:'6px',textAlign:'center',borderTop:'1px solid var(--border)',borderRight:'1px solid var(--border)'}} onClick={e=>e.stopPropagation()}>
                  <input type="checkbox" checked={seleccionados.has(inv.id)} onChange={()=>{setSeleccionados(prev=>{const n=new Set(prev);n.has(inv.id)?n.delete(inv.id):n.add(inv.id);return n})}} style={{cursor:'pointer',width:14,height:14,accentColor:'var(--accent)'}}/>
                </td>
                <td style={{padding:'6px 4px',textAlign:'center',fontSize:10,color:'var(--text3)',borderTop:'1px solid var(--border)',borderRight:'1px solid var(--border)'}}>{idx+1}</td>
                {COLS.map(col=>(
                  <td key={col.key} style={celda(inv.id,col.key)} onClick={()=>startEdit(inv.id,col.key,inv[col.key])}>
                    {editCell?.id===inv.id&&editCell?.field===col.key ? (
                      col.key==='mes'?(
                        <select value={editVal} onChange={e=>setEditVal(e.target.value)} onBlur={()=>commitEdit()} ref={inputRef} autoFocus
                          style={{background:'var(--bg2)',color:'var(--text)',border:'1px solid var(--accent)',borderRadius:4,fontSize:12,width:'100%',fontFamily:'var(--font)',padding:'2px'}}>
                          <option value="">— Selecciona mes —</option>{MESES.map(m=><option key={m} value={m}>{m}</option>)}
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
// PRESUPUESTO — con filtro de mes corregido (normMes)
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
    {key:'mes',          label:'Mes',                    w:85},
    {key:'gasto',        label:'Gasto (Nom. Producto, Cliente)', w:260},
    {key:'valorFactura', label:'Valor Factura',           w:125},
    {key:'canal',        label:'Canal',                   w:80},
    {key:'observacion',  label:'Observación (ATJ)',       w:180},
    {key:'estado',       label:'Estado',                  w:100},
    {key:'centroCostos', label:'Centro de Costos',        w:120},
    {key:'ordenCompra',  label:'Orden de Compra',         w:110},
    {key:'nitProveedor', label:'NIT Proveedor',           w:110},
    {key:'nomProveedor', label:'Nom. Proveedor',          w:140},
    {key:'docCargado',   label:'Doc. Cargado',            w:90},
    {key:'obsKatherine', label:'Obs. Katherine',          w:160},
  ]
  const ESTADOS_G = ['Pendiente','Ingresado','Aprobado','Pagado','Rechazado']
  const anios = [...new Set([...(data.gastosPresupuesto||[]).map(g=>g.anio),...data.presupuestos.map(p=>p.anio)])].sort()

  // FIX: todos los filtros usan normMes para comparar independiente de capitalización
  const gastos = (data.gastosPresupuesto||[]).filter(g=>
    (!filtroMes||normMes(g.mes)===normMes(filtroMes)) &&
    (!filtroAnio||g.anio===Number(filtroAnio)) &&
    (!busqueda||[g.gasto,g.observacion,g.estado,g.centroCostos,g.canal].some(v=>String(v||'').toLowerCase().includes(busqueda.toLowerCase())))
  ).sort((a,b)=>{
    const mi=MESES.indexOf(normMes(a.mes)),mj=MESES.indexOf(normMes(b.mes))
    if(mi===-1&&mj===-1) return 0; if(mi===-1) return 1; if(mj===-1) return -1; return mi-mj
  })

  const totalGastado = gastos.reduce((s,g)=>s+(Number(g.valorFactura)||0),0)

  // FIX: presAsignado usa normMes
  const presAsignado = data.presupuestos
    .filter(p=>(!filtroMes||normMes(p.mes)===normMes(filtroMes))&&(!filtroAnio||p.anio===Number(filtroAnio)))
    .reduce((s,p)=>s+(Number(p.monto)||0),0)

  const ejec = presAsignado>0?(totalGastado/presAsignado)*100:0
  const disponible = presAsignado - totalGastado

  const toggleTodos = () => setSeleccionados(prev=>prev.size===gastos.length&&gastos.length>0?new Set():new Set(gastos.map(g=>g.id)))
  const eliminarSel = () => {
    if(!seleccionados.size||!confirm('¿Eliminar '+seleccionados.size+' gastos?')) return
    const nd={...data,gastosPresupuesto:(data.gastosPresupuesto||[]).filter(g=>!seleccionados.has(g.id))}
    setData(nd);save(nd);setSeleccionados(new Set())
  }

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

  const addFila = () => {
    const n={id:Date.now(),anio:Number(filtroAnio)||2026,mes:filtroMes||'',gasto:'',valorFactura:0,canal:'',observacion:'',estado:'Pendiente',centroCostos:'',notas:''}
    const nd={...data,gastosPresupuesto:[...(data.gastosPresupuesto||[]),n]};setData(nd);save(nd,{insertar:n,tipo:'gastosPresupuesto'})
    setTimeout(()=>startEdit(n.id,'gasto',''),60)
  }

  const handlePaste = e => {
    const text=e.clipboardData.getData('text')
    if(!text.includes('\t')&&!text.includes('\n')) return
    e.preventDefault()
    const rows=text.trim().split('\n').map(r=>r.split('\t'))
    const nuevas=rows.filter(r=>r[0]||r[1]).map((r,i)=>({
      id:Date.now()+i, anio:Number(filtroAnio)||2026,
      mes:normMes(r[0]||filtroMes||''), gasto:r[1]||'',
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

  // FIX: resumenMeses usa normMes
  const resumenMeses = MESES.map(m=>{
    const gastado=(data.gastosPresupuesto||[]).filter(g=>normMes(g.mes)===m&&(!filtroAnio||g.anio===Number(filtroAnio))).reduce((s,g)=>s+(Number(g.valorFactura)||0),0)
    const asignado=data.presupuestos.filter(p=>normMes(p.mes)===m&&(!filtroAnio||p.anio===Number(filtroAnio))).reduce((s,p)=>s+(Number(p.monto)||0),0)
    return {mes:m,gastado,asignado,ejec:asignado>0?(gastado/asignado)*100:0,disponible:asignado-gastado}
  }).filter(r=>r.gastado>0||r.asignado>0)

  const submitPres = () => {
    if(!formP.mes||!formP.monto) return
    const entry={id:Date.now(),mes:formP.mes,anio:Number(formP.anio),monto:Number(formP.monto)}
    const nd={...data,presupuestos:[...data.presupuestos,entry]};setData(nd);save(nd,{insertar:entry,tipo:'presupuestos'});setModalPres(false);setFormP({mes:'',anio:2026,monto:''})
  }
  const delPres = id => { const nd={...data,presupuestos:data.presupuestos.filter(p=>p.id!==id)};setData(nd);save(nd,{eliminar:id,tipo:'presupuestos'}) }

  return (
    <div style={{display:'flex',flexDirection:'column',gap:18}}>
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

      <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:12}}>
        <KpiCard icon={DollarSign} label="Presupuesto asignado" value={cop(presAsignado)} sub={filtroMes||(filtroAnio||'')+' Total período'}/>
        <KpiCard icon={TrendingUp} label="Total ejecutado" value={cop(totalGastado)} sub={gastos.length+' registros'} accent="var(--accent2)"/>
        <KpiCard icon={BarChart2} label="% Ejecutado" value={ejec.toFixed(1)+'%'} sub={ejec>100?'Excedido':ejec>80?'Casi al límite':'Dentro del presupuesto'} accent={ejec>100?'var(--red)':ejec>80?'var(--yellow)':'var(--green)'}/>
        <KpiCard icon={DollarSign} label={disponible>=0?'Disponible':'Excedido'} value={cop(Math.abs(disponible))} accent={disponible>=0?'var(--green)':'var(--red)'}/>
      </div>

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
                <th style={{...S.th,width:36}}></th>
                {COLS_G.map(c=><th key={c.key} style={{...S.th,width:c.w}}>{c.label}</th>)}
              </tr>
            </thead>
            <tbody>
              {gastos.length===0&&(
                <tr><td colSpan={COLS_G.length+3} style={{padding:48,textAlign:'center',color:'var(--text3)',fontSize:13}}>
                  No hay gastos. <strong>Añadir fila</strong>, importar Excel, o pegar con <strong>Ctrl+V</strong>.
                </td></tr>
              )}
              {gastos.map((g,idx)=>(
                <tr key={g.id} style={{background:seleccionados.has(g.id)?'rgba(108,99,255,0.06)':'transparent'}}>
                  <td style={{padding:'6px',textAlign:'center',borderTop:'1px solid var(--border)',borderRight:'1px solid var(--border)'}}>
                    <input type="checkbox" checked={seleccionados.has(g.id)} onChange={()=>{setSeleccionados(prev=>{const n=new Set(prev);n.has(g.id)?n.delete(g.id):n.add(g.id);return n})}} style={{cursor:'pointer',width:14,height:14,accentColor:'var(--accent)'}}/>
                  </td>
                  <td style={{padding:'6px 4px',textAlign:'center',fontSize:10,color:'var(--text3)',borderTop:'1px solid var(--border)',borderRight:'1px solid var(--border)'}}>{idx+1}</td>
                  <td style={{padding:'4px',borderTop:'1px solid var(--border)',textAlign:'center',width:36}}>
                    <button onClick={()=>{if(window.confirm('¿Eliminar?')){const nd={...data,gastosPresupuesto:(data.gastosPresupuesto||[]).filter(x=>x.id!==g.id)};setData(nd);save(nd)}}} style={{background:'var(--red-soft)',color:'var(--red)',border:'none',borderRadius:6,padding:'2px 8px',cursor:'pointer',fontSize:13,fontWeight:700,lineHeight:1}}>✕</button>
                  </td>
                  {COLS_G.map(col=>(
                    <td key={col.key} style={celda(g.id,col.key)} onClick={()=>startEdit(g.id,col.key,g[col.key])}>
                      {editCell?.id===g.id&&editCell?.field===col.key ? (
                        col.key==='mes'?(
                          <select defaultValue={editVal} onChange={e=>{
                            const v=e.target.value
                            if(!editCell) return
                            const {id,field}=editCell
                            const gastosPresupuesto=(data.gastosPresupuesto||[]).map(g=>g.id!==id?g:{...g,[field]:v})
                            const nd={...data,gastosPresupuesto};setData(nd);save(nd);setEditCell(null);setEditVal(v)
                          }} ref={inputRef} autoFocus
                            style={{background:'var(--bg2)',color:'var(--text)',border:'1px solid var(--accent)',borderRadius:4,fontSize:12,width:'100%',fontFamily:'var(--font)',padding:'2px'}}>
                            <option value="">— Selecciona mes —</option>{MESES.map(m=><option key={m} value={m}>{m}</option>)}
                          </select>
                        ):col.key==='estado'?(
                          <select defaultValue={editVal} onChange={e=>{
                            const v=e.target.value
                            if(!editCell) return
                            const {id,field}=editCell
                            const gastosPresupuesto=(data.gastosPresupuesto||[]).map(g=>g.id!==id?g:{...g,[field]:v})
                            const nd={...data,gastosPresupuesto};setData(nd);save(nd);setEditCell(null);setEditVal(v)
                          }} ref={inputRef} autoFocus
                            style={{background:'var(--bg2)',color:'var(--text)',border:'1px solid var(--accent)',borderRadius:4,fontSize:12,width:'100%',fontFamily:'var(--font)',padding:'2px'}}>
                            {ESTADOS_G.map(s=><option key={s} value={s}>{s}</option>)}
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
                  <td colSpan={3} style={{padding:'8px 10px'}}/>
                  <td colSpan={2} style={{padding:'8px 12px',fontWeight:700,fontSize:12,color:'var(--text2)'}}>TOTAL {filtroMes&&'— '+filtroMes}</td>
                  <td style={{padding:'8px 12px',fontFamily:'var(--mono)',fontWeight:700,fontSize:13,color:'var(--accent2)'}}>{cop(totalGastado)}</td>
                  <td colSpan={4} style={{padding:'8px 10px'}}/>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      <div style={S.card}>
        <div style={{padding:'12px 18px',borderBottom:'1px solid var(--border)'}}><h4 style={{fontSize:11,fontWeight:600,color:'var(--text2)',textTransform:'uppercase',letterSpacing:'0.05em'}}>Presupuestos asignados por mes</h4></div>
        <table style={{width:'100%',borderCollapse:'collapse'}}>
          <thead><tr>{['Año','Mes','Monto asignado','Ejecutado','Disponible',''].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
          <tbody>
            {data.presupuestos.length===0&&<tr><td colSpan={6} style={{...S.td,textAlign:'center',color:'var(--text3)',padding:28}}>No hay presupuestos. Usa "Asignar presupuesto".</td></tr>}
            {[...data.presupuestos].sort((a,b)=>a.anio!==b.anio?b.anio-a.anio:MESES.indexOf(normMes(a.mes))-MESES.indexOf(normMes(b.mes))).map(p=>{
              // FIX: compara mes normalizado
              const ejec2=(data.gastosPresupuesto||[]).filter(g=>normMes(g.mes)===normMes(p.mes)&&g.anio===p.anio).reduce((s,g)=>s+(Number(g.valorFactura)||0),0)
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

function Ventas({ data, setData }) {
  const [modal, setModal] = useState(false)
  const [filtroMes, setFiltroMes] = useState('')
  const [filtroAnio, setFiltroAnio] = useState('')
  const [filtroDist, setFiltroDist] = useState('')
  const [busqueda, setBusqueda] = useState('')
  const [editId, setEditId] = useState(null)
  const blank = { anio:2026, mes:'', distribuidor:'', galones:'', ventaNeta:'', notas:'' }
  const [form, setForm] = useState(blank)

  const anios = [...new Set([...data.ventas.map(v=>v.anio),...data.inversiones.map(i=>i.anio)])].sort()
  const distribuidores = [...new Set([...data.ventas.map(v=>v.distribuidor),...data.inversiones.map(i=>i.distribuidor)])].sort()

  const lista = data.ventas.filter(v=>
    (!filtroMes||normMes(v.mes)===normMes(filtroMes)) &&
    (!filtroAnio||v.anio===Number(filtroAnio)) &&
    (!filtroDist||v.distribuidor===filtroDist) &&
    (!busqueda||String(v.distribuidor||'').toLowerCase().includes(busqueda.toLowerCase()))
  ).sort((a,b)=>{
    const mi=MESES.indexOf(normMes(a.mes)), mj=MESES.indexOf(normMes(b.mes))
    return mi!==mj ? mi-mj : String(a.distribuidor||'').localeCompare(String(b.distribuidor||''))
  })

  const norm = s => String(s||'').toLowerCase().replace(/[.\s-]/g,'').replace(/sas$|sa$/,'').trim()
  const getInv = (dist,mes,anio) =>
    data.inversiones.filter(i=>(i.distribuidor===dist||norm(i.distribuidor)===norm(dist))&&normMes(i.mes)===normMes(mes)&&i.anio===anio)
      .reduce((s,i)=>s+(Number(i.inversion)||0),0)

  const promAcum = (dist) => {
    const rows = data.ventas.filter(v=>(v.distribuidor===dist||norm(v.distribuidor)===norm(dist))&&(!filtroAnio||v.anio===Number(filtroAnio)))
    const totalGal = rows.reduce((s,v)=>s+(Number(v.galones)||0),0)
    const totalVta = rows.reduce((s,v)=>s+(Number(v.ventaNeta)||0),0)
    return totalGal>0 ? totalVta/totalGal : 0
  }

  const totalGalones = lista.reduce((s,v)=>s+(Number(v.galones)||0),0)
  const totalVenta   = lista.reduce((s,v)=>s+(Number(v.ventaNeta)||0),0)
  const totalInv     = lista.reduce((s,v)=>s+getInv(v.distribuidor,v.mes,v.anio),0)
  const precioPromedio = totalGalones>0 ? totalVenta/totalGalones : 0

  const submit = () => {
    if(!form.distribuidor||!form.mes) return
    const entry={...form, id:editId||Date.now(), galones:Number(form.galones)||0, ventaNeta:Number(form.ventaNeta)||0, anio:Number(form.anio)}
    const ventas=editId?data.ventas.map(v=>v.id===editId?entry:v):[...data.ventas,entry]
    const nd={...data,ventas};setData(nd);save(nd,editId?{}:{insertar:entry,tipo:'ventas'});setModal(false);setEditId(null);setForm(blank)
  }
  const del = id => { const nd={...data,ventas:data.ventas.filter(v=>v.id!==id)};setData(nd);save(nd,{eliminar:id,tipo:'ventas'}) }
  const edit = v => { setForm({...v,galones:String(v.galones),ventaNeta:String(v.ventaNeta)});setEditId(v.id);setModal(true) }

  const resumenDist = [...new Set(lista.map(v=>v.distribuidor))].map(d=>{
    const rows = lista.filter(v=>v.distribuidor===d)
    const gal = rows.reduce((s,v)=>s+(Number(v.galones)||0),0)
    const vta = rows.reduce((s,v)=>s+(Number(v.ventaNeta)||0),0)
    const inv = rows.reduce((s,v)=>s+getInv(v.distribuidor,v.mes,v.anio),0)
    return { dist:d, gal, vta, inv, precioMes: gal>0?vta/gal:0, precioAcum: promAcum(d), pctInv: vta>0?(inv/vta)*100:0 }
  }).sort((a,b)=>b.vta-a.vta)

  return (
    <div style={{display:'flex',flexDirection:'column',gap:18}}>
      <div style={{display:'flex',gap:8,flexWrap:'wrap',alignItems:'center'}}>
        <div style={{position:'relative',display:'flex',alignItems:'center'}}>
          <Search size={13} style={{position:'absolute',left:9,color:'var(--text3)',pointerEvents:'none'}}/>
          <input value={busqueda} onChange={e=>setBusqueda(e.target.value)} placeholder="Buscar distribuidor..." style={{paddingLeft:30,width:200,fontSize:13}}/>
        </div>
        <select value={filtroAnio} onChange={e=>setFiltroAnio(e.target.value)} style={{width:100}}>
          <option value="">Año</option>{anios.map(a=><option key={a}>{a}</option>)}
        </select>
        <select value={filtroMes} onChange={e=>setFiltroMes(e.target.value)} style={{width:140}}>
          <option value="">Todos los meses</option>{MESES.map(m=><option key={m}>{m}</option>)}
        </select>
        <select value={filtroDist} onChange={e=>setFiltroDist(e.target.value)} style={{width:220}}>
          <option value="">Todos los distribuidores</option>{distribuidores.map(d=><option key={d}>{d}</option>)}
        </select>
        {(filtroMes||filtroAnio||filtroDist||busqueda)&&
          <button onClick={()=>{setFiltroMes('');setFiltroAnio('');setFiltroDist('');setBusqueda('')}} style={{...S.btn('var(--bg3)','var(--text2)'),padding:'5px 10px',fontSize:12}}>✕</button>}
        <button onClick={()=>{setEditId(null);setForm(blank);setModal(true)}} style={{...S.btn('var(--accent)','#fff'),marginLeft:'auto'}}>
          <PlusCircle size={15}/> Registrar venta
        </button>
      </div>

      <div style={{display:'grid',gridTemplateColumns:'repeat(5,1fr)',gap:12}}>
        <KpiCard icon={ShoppingCart} label="Venta neta total" value={cop(totalVenta)} sub={lista.length+' registros'} accent="var(--green)"/>
        <KpiCard icon={BarChart2} label="Total galones" value={num(Math.round(totalGalones))} accent="var(--accent2)"/>
        <KpiCard icon={DollarSign} label="Precio prom. período" value={cop(Math.round(precioPromedio))} sub="Venta Neta ÷ Galones" accent="var(--orange)"/>
        <KpiCard icon={TrendingUp} label="Inversión total" value={cop(totalInv)} sub="Desde pestaña Inversiones"/>
        <KpiCard icon={BarChart2} label="% Inv / Venta" value={totalVenta>0?((totalInv/totalVenta)*100).toFixed(1)+'%':'—'} accent={totalVenta>0&&totalInv/totalVenta>0.15?'var(--red)':'var(--green)'}/>
      </div>

      <div style={S.card}>
        <div style={{padding:'12px 18px',borderBottom:'1px solid var(--border)'}}>
          <h4 style={{fontSize:11,fontWeight:600,color:'var(--text2)',textTransform:'uppercase',letterSpacing:'0.05em'}}>Detalle por registro</h4>
        </div>
        <div style={{overflowX:'auto'}}>
          <table style={{width:'100%',borderCollapse:'collapse',minWidth:900}}>
            <thead>
              <tr style={{background:'var(--bg3)'}}>
                {['Año','Mes','Distribuidor','Galones','Venta Neta','Precio/Galón','Precio Acum./Galón','Inversión','% Inv/Venta',''].map(h=><th key={h} style={S.th}>{h}</th>)}
              </tr>
            </thead>
            <tbody>
              {lista.length===0&&<tr><td colSpan={10} style={{...S.td,textAlign:'center',color:'var(--text3)',padding:40}}>No hay ventas registradas.</td></tr>}
              {lista.map(v=>{
                const inv = getInv(v.distribuidor,v.mes,v.anio)
                const gal = Number(v.galones)||0
                const vta = Number(v.ventaNeta)||0
                const precio = gal>0 ? vta/gal : 0
                const acum  = promAcum(v.distribuidor)
                const pct   = vta>0 ? (inv/vta)*100 : 0
                return (
                  <tr key={v.id}>
                    <td style={{...S.td,color:'var(--text2)'}}>{v.anio||'—'}</td>
                    <td style={{...S.td,color:'var(--text2)'}}>{v.mes}</td>
                    <td style={{...S.td,fontWeight:500}}>{v.distribuidor}</td>
                    <td style={{...S.td,fontFamily:'var(--mono)'}}>{num(Math.round(gal))}</td>
                    <td style={{...S.td,fontFamily:'var(--mono)',color:'var(--green)',fontWeight:500}}>{cop(vta)}</td>
                    <td style={{...S.td,fontFamily:'var(--mono)',color:'var(--orange)'}}>{cop(Math.round(precio))}</td>
                    <td style={{...S.td,fontFamily:'var(--mono)',color:'var(--yellow)',fontWeight:500}}>{acum>0?cop(Math.round(acum)):'—'}</td>
                    <td style={{...S.td,fontFamily:'var(--mono)'}}>{inv>0?cop(inv):<span style={{color:'var(--text3)'}}>—</span>}</td>
                    <td style={{...S.td}}>
                      {inv>0?<span style={{fontFamily:'var(--mono)',fontWeight:600,color:pct>15?'var(--red)':pct>10?'var(--yellow)':'var(--green)'}}>{pct.toFixed(1)}%</span>:<span style={{color:'var(--text3)'}}>—</span>}
                    </td>
                    <td style={S.td}>
                      <div style={{display:'flex',gap:5}}>
                        <button onClick={()=>edit(v)} style={{...S.btn('var(--bg3)','var(--text2)'),padding:'4px 7px'}}><Edit2 size={13}/></button>
                        <button onClick={()=>del(v.id)} style={{...S.btn('var(--red-soft)','var(--red)'),padding:'4px 7px'}}><Trash2 size={13}/></button>
                      </div>
                    </td>
                  </tr>
                )
              })}
              {lista.length>0&&(
                <tr style={{background:'var(--bg3)',borderTop:'2px solid var(--border2)'}}>
                  <td colSpan={3} style={{...S.td,fontWeight:700,color:'var(--text2)'}}>TOTAL {filtroMes&&'— '+filtroMes}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700}}>{num(Math.round(totalGalones))}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700,color:'var(--green)'}}>{cop(totalVenta)}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:600,color:'var(--orange)'}}>{cop(Math.round(precioPromedio))}</td>
                  <td style={S.td}/>
                  <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700}}>{cop(totalInv)}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700,color:totalVenta>0&&totalInv/totalVenta>0.15?'var(--red)':'var(--green)'}}>{totalVenta>0?((totalInv/totalVenta)*100).toFixed(1)+'%':'—'}</td>
                  <td style={S.td}/>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {resumenDist.length>0&&(
        <div style={S.card}>
          <div style={{padding:'12px 18px',borderBottom:'1px solid var(--border)',display:'flex',alignItems:'center',justifyContent:'space-between'}}>
            <h4 style={{fontSize:11,fontWeight:600,color:'var(--text2)',textTransform:'uppercase',letterSpacing:'0.05em'}}>Resumen acumulado por distribuidor</h4>
            <span style={{fontSize:11,color:'var(--text3)'}}>Precio Acum. = Total Venta Neta ÷ Total Galones</span>
          </div>
          <div style={{overflowX:'auto'}}>
            <table style={{width:'100%',borderCollapse:'collapse'}}>
              <thead>
                <tr style={{background:'var(--bg3)'}}>
                  {['Distribuidor','Galones','Venta Neta','Precio/Galón período','Precio Acum./Galón','Inversión','% Inv/Venta'].map(h=><th key={h} style={S.th}>{h}</th>)}
                </tr>
              </thead>
              <tbody>
                {resumenDist.map((d,i)=>(
                  <tr key={i} onClick={()=>setFiltroDist(filtroDist===d.dist?'':d.dist)} style={{cursor:'pointer',background:filtroDist===d.dist?'rgba(108,99,255,0.06)':'transparent'}}>
                    <td style={{...S.td,fontWeight:600,color:filtroDist===d.dist?'var(--accent2)':'var(--text)'}}>{d.dist}</td>
                    <td style={{...S.td,fontFamily:'var(--mono)'}}>{num(Math.round(d.gal))}</td>
                    <td style={{...S.td,fontFamily:'var(--mono)',color:'var(--green)',fontWeight:500}}>{cop(d.vta)}</td>
                    <td style={{...S.td,fontFamily:'var(--mono)',color:'var(--orange)'}}>{cop(Math.round(d.precioMes))}</td>
                    <td style={{...S.td,fontFamily:'var(--mono)',color:'var(--yellow)',fontWeight:600}}>{cop(Math.round(d.precioAcum))}</td>
                    <td style={{...S.td,fontFamily:'var(--mono)'}}>{d.inv>0?cop(d.inv):<span style={{color:'var(--text3)'}}>—</span>}</td>
                    <td style={{...S.td}}>
                      {d.inv>0?<span style={{fontFamily:'var(--mono)',fontWeight:600,color:d.pctInv>15?'var(--red)':d.pctInv>10?'var(--yellow)':'var(--green)'}}>{d.pctInv.toFixed(1)}%</span>:<span style={{color:'var(--text3)'}}>—</span>}
                    </td>
                  </tr>
                ))}
                <tr style={{borderTop:'2px solid var(--border2)',background:'var(--bg3)'}}>
                  <td style={{...S.td,fontWeight:700}}>TOTAL</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700}}>{num(Math.round(totalGalones))}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700,color:'var(--green)'}}>{cop(totalVenta)}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:600,color:'var(--orange)'}}>{cop(Math.round(precioPromedio))}</td>
                  <td style={S.td}/>
                  <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700}}>{cop(totalInv)}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700}}>{totalVenta>0?((totalInv/totalVenta)*100).toFixed(1)+'%':'—'}</td>
                </tr>
              </tbody>
            </table>
          </div>
          <div style={{padding:'8px 18px',fontSize:11,color:'var(--text3)',borderTop:'1px solid var(--border)'}}>
            💡 Clic en una fila para filtrar el detalle
          </div>
        </div>
      )}

      {modal&&(
        <Modal title={editId?'Editar venta':'Registrar venta'} onClose={()=>{setModal(false);setEditId(null)}}>
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:14}}>
            <Field label="Año"><input type="number" value={form.anio} onChange={e=>setForm({...form,anio:e.target.value})} placeholder="2026"/></Field>
            <Field label="Mes *"><select value={form.mes} onChange={e=>setForm({...form,mes:e.target.value})}><option value="">Selecciona...</option>{MESES.map(m=><option key={m}>{m}</option>)}</select></Field>
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
// NOTA: Las siguientes funciones son IDÉNTICAS al original:
// Planes, Pendientes, ApoyoCierre, exportarExcel, parsearExcel,
// Asistente, LoginScreen, hexToRgb, DashboardLider, 
// PresupuestoConsolidado, App (función principal)
//
// Cópialas exactamente desde tu archivo App.jsx original
// desde la línea que empieza con:
//   function Planes({ data, setData }) {
// hasta el final del archivo (export default function App())
// ═══════════════════════════════════════════════════════


// ═══════════════════════════════════════════════════════
// PLANES
// ═══════════════════════════════════════════════════════
function Planes({ data, setData }) {
  const [modal, setModal] = useState(false)
  const [modalHoja, setModalHoja] = useState(null)
  const [filtroQ, setFiltroQ] = useState('')
  const [filtroAnio, setFiltroAnio] = useState('2026')
  const [editId, setEditId] = useState(null)
  const [nuevaNota, setNuevaNota] = useState('')

  const MESES_Q = {
    Q1:['Enero','Febrero','Marzo'],
    Q2:['Abril','Mayo','Junio'],
    Q3:['Julio','Agosto','Septiembre'],
    Q4:['Octubre','Noviembre','Diciembre'],
  }

  const blankPlanItem = () => ({ id: Date.now()+Math.random(), nombre:'', tienePago:'No', descripcionPago:'', condiciones:'' })

  const blank = {
    distribuidor:'', anio:2026, quarter:'Q1', estado:'Activo',
    planesItems: [blankPlanItem()],
    metaMes1:'', metaMes2:'', metaMes3:'',
    metaGalonesMes1:'', metaGalonesMes2:'', metaGalonesMes3:'',
    tieneCVC:'No', montoCVC:'', detalleCVC:'',
    tienePromotora:'No', nombrePromotora:'', montoPromotora:'',
    condicionesGenerales:'', acuerdos:'', notas:'', historial:[],
  }
  const [form, setForm] = useState(blank)

  const anios = [...new Set([...data.planes.map(p=>p.anio),...data.inversiones.map(i=>i.anio)])].sort()
  const distribuidores = [...new Set([...data.planes.map(p=>p.distribuidor),...data.inversiones.map(i=>i.distribuidor)])].sort()
  const lista = data.planes.filter(p=>(!filtroQ||p.quarter===filtroQ)&&(!filtroAnio||p.anio===Number(filtroAnio)))

  const totalMetaVenta = f => (Number(f.metaMes1)||0)+(Number(f.metaMes2)||0)+(Number(f.metaMes3)||0)
  const totalMetaGalones = f => (Number(f.metaGalonesMes1)||0)+(Number(f.metaGalonesMes2)||0)+(Number(f.metaGalonesMes3)||0)
  const mesesQ = q => MESES_Q[q]||['Mes 1','Mes 2','Mes 3']

  const updatePlanItem = (idx, field, val) => {
    const items = form.planesItems.map((p,i)=>i===idx?{...p,[field]:val}:p)
    setForm({...form, planesItems:items})
  }
  const addPlanItem = () => setForm({...form, planesItems:[...form.planesItems, blankPlanItem()]})
  const removePlanItem = idx => {
    if(form.planesItems.length===1) return
    setForm({...form, planesItems:form.planesItems.filter((_,i)=>i!==idx)})
  }

  const submit = () => {
    if(!form.distribuidor||!form.quarter) return
    const entry = {
      ...form,
      id: editId||Date.now(),
      anio: Number(form.anio),
      metaVenta: totalMetaVenta(form),
      metaGalones: totalMetaGalones(form),
      tiposPlan: form.planesItems.filter(p=>p.nombre).map(p=>p.nombre),
      historial: form.historial||[],
    }
    const planes = editId ? data.planes.map(p=>p.id===editId?entry:p) : [...data.planes,entry]
    const nd={...data,planes}
    setData(nd)
    if(editId) {
      save(nd, {eliminar:editId, tipo:'planes'})
      setTimeout(()=>save(nd, {insertar:entry, tipo:'planes'}), 500)
    } else {
      save(nd, {insertar:entry, tipo:'planes'})
    }
    setModal(false);setEditId(null);setForm(blank)
  }

  const del = id => {
    const nd={...data,planes:data.planes.filter(p=>p.id!==id)}
    setData(nd)
    save(nd, {eliminar:id, tipo:'planes'})
  }

  const edit = p => {
    const items = p.planesItems?.length>0 ? p.planesItems :
      (p.tiposPlan||[]).map(n=>({id:Date.now()+Math.random(),nombre:n,tienePago:'No',descripcionPago:'',condiciones:''}))
    if(items.length===0) items.push(blankPlanItem())
    setForm({...blank,...p, planesItems:items,
      metaMes1:p.metaMes1||'', metaMes2:p.metaMes2||'', metaMes3:p.metaMes3||'',
      metaGalonesMes1:p.metaGalonesMes1||'', metaGalonesMes2:p.metaGalonesMes2||'', metaGalonesMes3:p.metaGalonesMes3||'',
      tieneCVC:p.tieneCVC||'No', montoCVC:p.montoCVC||'', detalleCVC:p.detalleCVC||'',
      tienePromotora:p.tienePromotora||'No', nombrePromotora:p.nombrePromotora||'', montoPromotora:p.montoPromotora||'',
      historial:p.historial||[],
    })
    setEditId(p.id);setModal(true)
  }

  const agregarNota = planId => {
    if(!nuevaNota.trim()) return
    const nota={texto:nuevaNota,fecha:new Date().toLocaleDateString('es-CO')}
    const planes=data.planes.map(p=>p.id===planId?{...p,historial:[...(p.historial||[]),nota]}:p)
    const nd={...data,planes};setData(nd);save(nd)
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
          <option value="">Año</option>{anios.map(a=><option key={a}>{a}</option>)}
        </select>
        <select value={filtroQ} onChange={e=>setFiltroQ(e.target.value)} style={{width:120}}>
          <option value="">Todos los Q</option>{QUARTERS.map(q=><option key={q}>{q}</option>)}
        </select>
        <button onClick={()=>{setEditId(null);setForm(blank);setModal(true)}} style={{...S.btn('var(--accent)','#fff'),marginLeft:'auto'}}>
          <PlusCircle size={15}/> Nuevo plan
        </button>
      </div>

      {qResumen.length>0&&(
        <div style={{display:'flex',gap:12,flexWrap:'wrap'}}>
          {qResumen.map(q=>(
            <div key={q.q} onClick={()=>setFiltroQ(filtroQ===q.q?'':q.q)}
              style={{background:filtroQ===q.q?'var(--accent-soft)':'var(--bg2)',border:'1px solid '+(filtroQ===q.q?'var(--accent)':'var(--border)'),borderRadius:10,padding:'12px 20px',cursor:'pointer',display:'flex',flexDirection:'column',gap:4,minWidth:110}}>
              <span style={{fontSize:18,fontWeight:700,color:filtroQ===q.q?'var(--accent2)':'var(--text)'}}>{q.q}</span>
              <span style={{fontSize:11,color:'var(--text3)'}}>{q.activos} activo{q.activos!==1?'s':''} / {q.total} total</span>
            </div>
          ))}
        </div>
      )}

      <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fill,minmax(360px,1fr))',gap:14}}>
        {lista.length===0&&<div style={{...S.card,padding:40,textAlign:'center',color:'var(--text3)',gridColumn:'1/-1'}}>No hay planes. Crea el primero.</div>}
        {lista.map(p=>{
          const galonesReal=data.ventas.filter(v=>v.distribuidor===p.distribuidor&&(!filtroAnio||v.anio===Number(filtroAnio))).reduce((s,v)=>s+(Number(v.galones)||0),0)
          const ventaReal=data.ventas.filter(v=>v.distribuidor===p.distribuidor&&(!filtroAnio||v.anio===Number(filtroAnio))).reduce((s,v)=>s+(Number(v.ventaNeta)||0),0)
          const invTotal=data.inversiones.filter(i=>i.distribuidor===p.distribuidor&&(!filtroAnio||i.anio===Number(filtroAnio))).reduce((s,i)=>s+(Number(i.inversion)||0),0)
          const metaVta = p.metaVenta||totalMetaVenta(p)
          const metaGal = p.metaGalones||totalMetaGalones(p)
          const cumplVta = metaVta>0?(ventaReal/metaVta)*100:0
          const cumplGal = metaGal>0?(galonesReal/metaGal)*100:0
          const items = p.planesItems||[]
          return (
            <div key={p.id} style={{...S.card,display:'flex',flexDirection:'column'}}>
              <div style={{padding:'14px 18px',borderBottom:'1px solid var(--border)',display:'flex',alignItems:'flex-start',justifyContent:'space-between',gap:8}}>
                <div>
                  <div style={{fontWeight:600,fontSize:14,marginBottom:5}}>{p.distribuidor}</div>
                  <div style={{display:'flex',gap:6,flexWrap:'wrap'}}>
                    <span style={{background:'var(--bg4)',color:'var(--accent2)',fontSize:11,fontWeight:600,padding:'2px 8px',borderRadius:5}}>{p.quarter} · {p.anio}</span>
                    <Badge label={p.estado}/>
                    {p.tieneCVC==='Sí'&&<span style={{background:'rgba(6,182,212,0.15)',color:'#06b6d4',fontSize:10,fontWeight:600,padding:'2px 7px',borderRadius:5}}>CVC</span>}
                    {p.tienePromotora==='Sí'&&<span style={{background:'rgba(168,139,250,0.15)',color:'#a78bfa',fontSize:10,fontWeight:600,padding:'2px 7px',borderRadius:5}}>Promotora</span>}
                  </div>
                </div>
                <div style={{display:'flex',gap:5}}>
                  <button onClick={()=>setModalHoja(p)} style={{...S.btn('var(--bg3)','var(--text2)'),padding:'4px 8px',fontSize:11}}><BookOpen size={12}/> Ver</button>
                  <button onClick={()=>edit(p)} style={{...S.btn('var(--bg3)','var(--text2)'),padding:'4px 8px'}}><Edit2 size={12}/></button>
                  <button onClick={()=>del(p.id)} style={{...S.btn('var(--red-soft)','var(--red)'),padding:'4px 8px'}}><Trash2 size={12}/></button>
                </div>
              </div>
              <div style={{padding:'14px 18px',display:'flex',flexDirection:'column',gap:10,flex:1}}>
                {items.filter(i=>i.nombre).length>0&&(
                  <div style={{display:'flex',flexDirection:'column',gap:6}}>
                    {items.filter(i=>i.nombre).map((item,i)=>(
                      <div key={i} style={{background:'var(--bg3)',borderRadius:8,padding:'8px 12px'}}>
                        <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:item.descripcionPago||item.condiciones?4:0}}>
                          <span style={{fontSize:12,fontWeight:600,color:'var(--accent2)'}}>{item.nombre}</span>
                          {item.tienePago==='Sí'&&<span style={{fontSize:10,background:'var(--green-soft)',color:'var(--green)',padding:'1px 6px',borderRadius:4,fontWeight:600}}>Con pago</span>}
                        </div>
                        {item.descripcionPago&&<div style={{fontSize:11,color:'var(--text2)'}}>{item.descripcionPago}</div>}
                        {item.condiciones&&<div style={{fontSize:11,color:'var(--text3)',marginTop:2,borderLeft:'2px solid var(--border2)',paddingLeft:6}}>{item.condiciones.slice(0,60)}{item.condiciones.length>60?'…':''}</div>}
                      </div>
                    ))}
                  </div>
                )}
                <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:8}}>
                  <div style={{background:'var(--bg3)',borderRadius:8,padding:'8px 12px'}}>
                    <div style={{fontSize:10,color:'var(--text3)',marginBottom:2}}>META VENTA Q</div>
                    <div style={{fontFamily:'var(--mono)',fontWeight:600,fontSize:13}}>{cop(metaVta)}</div>
                    {metaVta>0&&<div style={{fontSize:10,color:'var(--text2)',marginTop:2}}>Real: {cop(ventaReal)} ({cumplVta.toFixed(0)}%)</div>}
                  </div>
                  <div style={{background:'var(--bg3)',borderRadius:8,padding:'8px 12px'}}>
                    <div style={{fontSize:10,color:'var(--text3)',marginBottom:2}}>META GALONES Q</div>
                    <div style={{fontFamily:'var(--mono)',fontWeight:600,fontSize:13}}>{num(metaGal)}</div>
                    {metaGal>0&&<div style={{fontSize:10,color:'var(--text2)',marginTop:2}}>Real: {num(galonesReal)} ({cumplGal.toFixed(0)}%)</div>}
                  </div>
                </div>
                {metaVta>0&&(
                  <div>
                    <div style={{display:'flex',justifyContent:'space-between',fontSize:10,color:'var(--text3)',marginBottom:3}}>
                      <span>Cumplimiento venta Q</span>
                      <span style={{color:cumplVta>=100?'var(--green)':cumplVta>=70?'var(--yellow)':'var(--red)'}}>{cumplVta.toFixed(1)}%</span>
                    </div>
                    <div style={{height:5,background:'var(--bg4)',borderRadius:3}}>
                      <div style={{width:Math.min(cumplVta,100)+'%',height:'100%',background:cumplVta>=100?'var(--green)':cumplVta>=70?'var(--yellow)':'var(--red)',borderRadius:3}}/>
                    </div>
                  </div>
                )}
                {(p.tieneCVC==='Sí'||p.tienePromotora==='Sí')&&(
                  <div style={{display:'flex',gap:7,flexWrap:'wrap'}}>
                    {p.tieneCVC==='Sí'&&<div style={{background:'rgba(6,182,212,0.08)',border:'1px solid rgba(6,182,212,0.2)',borderRadius:6,padding:'5px 10px',fontSize:11}}>
                      <span style={{color:'#06b6d4',fontWeight:600}}>CVC </span>
                      <span style={{color:'var(--text2)'}}>{p.montoCVC?cop(Number(p.montoCVC)):''} {p.detalleCVC}</span>
                    </div>}
                    {p.tienePromotora==='Sí'&&<div style={{background:'rgba(168,139,250,0.08)',border:'1px solid rgba(168,139,250,0.2)',borderRadius:6,padding:'5px 10px',fontSize:11}}>
                      <span style={{color:'#a78bfa',fontWeight:600}}>Promotora: </span>
                      <span style={{color:'var(--text2)'}}>{p.nombrePromotora}{p.montoPromotora?' — '+cop(Number(p.montoPromotora)):''}</span>
                    </div>}
                  </div>
                )}
                <div style={{display:'flex',justifyContent:'space-between',fontSize:12,color:'var(--text2)'}}>
                  <span>Inversión acumulada:</span>
                  <span style={{fontFamily:'var(--mono)',color:'var(--accent2)',fontWeight:600}}>{cop(invTotal)}</span>
                </div>
                {(p.historial?.length||0)>0&&<div style={{fontSize:11,color:'var(--text3)'}}>📝 {p.historial.length} nota{p.historial.length!==1?'s':''}</div>}
              </div>
            </div>
          )
        })}
      </div>

      {modal&&(
        <Modal title={editId?'Editar plan':'Nuevo plan'} onClose={()=>{setModal(false);setEditId(null)}} wide>
          <div style={{display:'flex',flexDirection:'column',gap:18}}>
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:14}}>
              <Field label="Distribuidor *" span>
                <input list="dist-plan" value={form.distribuidor} onChange={e=>setForm({...form,distribuidor:e.target.value})} placeholder="Nombre del distribuidor"/>
                <datalist id="dist-plan">{distribuidores.map(d=><option key={d} value={d}/>)}</datalist>
              </Field>
              <Field label="Año"><input type="number" value={form.anio} onChange={e=>setForm({...form,anio:e.target.value})}/></Field>
              <Field label="Quarter *"><select value={form.quarter} onChange={e=>setForm({...form,quarter:e.target.value})}>{QUARTERS.map(q=><option key={q}>{q}</option>)}</select></Field>
              <Field label="Estado"><select value={form.estado} onChange={e=>setForm({...form,estado:e.target.value})}>{ESTADOS_PLAN.map(s=><option key={s}>{s}</option>)}</select></Field>
            </div>
            <div>
              <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:10}}>
                <span style={{fontSize:12,fontWeight:600,color:'var(--text2)',textTransform:'uppercase',letterSpacing:'0.05em'}}>Planes del {form.quarter}</span>
                <button onClick={addPlanItem} style={{...S.btn('var(--accent-soft)','var(--accent2)'),fontSize:12,padding:'4px 12px',border:'1px solid rgba(108,99,255,0.25)'}}><PlusCircle size={13}/> Agregar plan</button>
              </div>
              <div style={{display:'flex',flexDirection:'column',gap:10}}>
                {form.planesItems.map((item,idx)=>(
                  <div key={item.id||idx} style={{background:'var(--bg3)',borderRadius:12,padding:'14px 16px',border:'1px solid var(--border2)',position:'relative'}}>
                    {form.planesItems.length>1&&(
                      <button onClick={()=>removePlanItem(idx)} style={{position:'absolute',top:10,right:10,...S.btn('var(--red-soft)','var(--red)'),padding:'3px 7px',fontSize:11}}><X size={12}/></button>
                    )}
                    <div style={{display:'flex',flexDirection:'column',gap:10}}>
                      <Field label={'Plan '+(idx+1)+' — Nombre *'}>
                        <input value={item.nombre} onChange={e=>updatePlanItem(idx,'nombre',e.target.value)} placeholder="Ej: Prolub acelera tu crecimiento..."/>
                      </Field>
                      <div style={{display:'flex',alignItems:'center',gap:12}}>
                        <span style={{fontSize:12,color:'var(--text2)',fontWeight:500}}>¿Tiene pago?</span>
                        <div style={{display:'flex',gap:7}}>
                          {['Sí','No'].map(v=>(
                            <button key={v} onClick={()=>updatePlanItem(idx,'tienePago',v)}
                              style={{...S.btn(item.tienePago===v?'var(--green-soft)':'var(--bg4)',item.tienePago===v?'var(--green)':'var(--text2)'),padding:'4px 14px',fontSize:12,border:item.tienePago===v?'1px solid rgba(61,214,140,0.3)':'1px solid var(--border2)'}}>{v}</button>
                          ))}
                        </div>
                      </div>
                      {item.tienePago==='Sí'&&(
                        <Field label="¿Qué incluye el pago?">
                          <textarea value={item.descripcionPago} onChange={e=>updatePlanItem(idx,'descripcionPago',e.target.value)} rows={2} style={{resize:'vertical'}} placeholder="Ej: Promotora $1.5M/mes..."/>
                        </Field>
                      )}
                      <Field label="Condiciones del plan">
                        <textarea value={item.condiciones} onChange={e=>updatePlanItem(idx,'condiciones',e.target.value)} rows={2} style={{resize:'vertical'}} placeholder="Condicionado a compra mínima..."/>
                      </Field>
                    </div>
                  </div>
                ))}
              </div>
            </div>
            <div>
              <div style={{fontSize:12,fontWeight:600,color:'var(--text2)',marginBottom:10,textTransform:'uppercase',letterSpacing:'0.05em'}}>Metas por mes — {form.quarter} ({mesesQ(form.quarter).join(' · ')})</div>
              <div style={{display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:12}}>
                {mesesQ(form.quarter).map((mes,i)=>(
                  <div key={mes} style={{background:'var(--bg3)',borderRadius:10,padding:'12px 14px',display:'flex',flexDirection:'column',gap:8}}>
                    <div style={{fontSize:12,fontWeight:600,color:'var(--accent2)'}}>{mes}</div>
                    <Field label="Meta Venta (COP)"><input type="number" value={form['metaMes'+(i+1)]} onChange={e=>setForm({...form,['metaMes'+(i+1)]:e.target.value})} placeholder="0"/></Field>
                    <Field label="Meta Galones"><input type="number" value={form['metaGalonesMes'+(i+1)]} onChange={e=>setForm({...form,['metaGalonesMes'+(i+1)]:e.target.value})} placeholder="0"/></Field>
                  </div>
                ))}
              </div>
              <div style={{display:'flex',gap:20,marginTop:10,padding:'10px 14px',background:'var(--accent-soft)',borderRadius:8,border:'1px solid rgba(108,99,255,0.2)'}}>
                <span style={{fontSize:13,color:'var(--text2)'}}>Total Q — Venta: <strong style={{color:'var(--accent2)',fontFamily:'var(--mono)'}}>{cop(totalMetaVenta(form))}</strong></span>
                <span style={{fontSize:13,color:'var(--text2)'}}>Galones: <strong style={{color:'var(--accent2)',fontFamily:'var(--mono)'}}>{num(totalMetaGalones(form))}</strong></span>
              </div>
            </div>
            <div style={{background:'var(--bg3)',borderRadius:10,padding:'14px 16px',display:'flex',flexDirection:'column',gap:10}}>
              <div style={{display:'flex',alignItems:'center',gap:12}}>
                <span style={{fontSize:13,fontWeight:600,color:'#06b6d4'}}>¿Tiene CVC?</span>
                <div style={{display:'flex',gap:8}}>
                  {['Sí','No'].map(v=>(
                    <button key={v} onClick={()=>setForm({...form,tieneCVC:v})}
                      style={{...S.btn(form.tieneCVC===v?'rgba(6,182,212,0.2)':'var(--bg4)',form.tieneCVC===v?'#06b6d4':'var(--text2)'),padding:'5px 16px',fontSize:13,border:form.tieneCVC===v?'1px solid rgba(6,182,212,0.4)':'1px solid var(--border2)'}}>{v}</button>
                  ))}
                </div>
              </div>
              {form.tieneCVC==='Sí'&&(
                <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:10}}>
                  <Field label="Monto CVC (COP)"><input type="number" value={form.montoCVC} onChange={e=>setForm({...form,montoCVC:e.target.value})} placeholder="0"/></Field>
                  <Field label="Detalle"><input value={form.detalleCVC} onChange={e=>setForm({...form,detalleCVC:e.target.value})} placeholder="Descripción del CVC..."/></Field>
                </div>
              )}
            </div>
            <div style={{background:'var(--bg3)',borderRadius:10,padding:'14px 16px',display:'flex',flexDirection:'column',gap:10}}>
              <div style={{display:'flex',alignItems:'center',gap:12}}>
                <span style={{fontSize:13,fontWeight:600,color:'#a78bfa'}}>¿Tiene Promotora?</span>
                <div style={{display:'flex',gap:8}}>
                  {['Sí','No'].map(v=>(
                    <button key={v} onClick={()=>setForm({...form,tienePromotora:v})}
                      style={{...S.btn(form.tienePromotora===v?'rgba(168,139,250,0.2)':'var(--bg4)',form.tienePromotora===v?'#a78bfa':'var(--text2)'),padding:'5px 16px',fontSize:13,border:form.tienePromotora===v?'1px solid rgba(168,139,250,0.4)':'1px solid var(--border2)'}}>{v}</button>
                  ))}
                </div>
              </div>
              {form.tienePromotora==='Sí'&&(
                <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:10}}>
                  <Field label="Nombre de la promotora"><input value={form.nombrePromotora} onChange={e=>setForm({...form,nombrePromotora:e.target.value})} placeholder="Nombre completo..."/></Field>
                  <Field label="Monto mensual (COP)"><input type="number" value={form.montoPromotora} onChange={e=>setForm({...form,montoPromotora:e.target.value})} placeholder="0"/></Field>
                </div>
              )}
            </div>
            <Field label="Condiciones generales del acuerdo">
              <textarea value={form.condicionesGenerales||form.condiciones||''} onChange={e=>setForm({...form,condicionesGenerales:e.target.value})} rows={2} style={{resize:'vertical'}} placeholder="Condiciones generales..."/>
            </Field>
            <Field label="Acuerdos especiales">
              <textarea value={form.acuerdos} onChange={e=>setForm({...form,acuerdos:e.target.value})} rows={2} style={{resize:'vertical'}} placeholder="Bonificaciones, descuentos adicionales..."/>
            </Field>
            <Field label="Notas internas">
              <textarea value={form.notas} onChange={e=>setForm({...form,notas:e.target.value})} rows={2} style={{resize:'vertical'}} placeholder="Alertas, oportunidades..."/>
            </Field>
          </div>
          <div style={{display:'flex',gap:10,justifyContent:'flex-end',marginTop:20}}>
            <button onClick={()=>{setModal(false);setEditId(null)}} style={S.btn('var(--bg3)','var(--text2)')}>Cancelar</button>
            <button onClick={submit} style={S.btn('var(--accent)','#fff')}><Check size={15}/> Guardar plan</button>
          </div>
        </Modal>
      )}

      {modalHoja&&(
        <Modal title={'Hoja de vida — '+modalHoja.distribuidor} onClose={()=>{setModalHoja(null);setNuevaNota('')}} wide>
          <div style={{display:'flex',flexDirection:'column',gap:16}}>
            <div style={{display:'flex',gap:8,flexWrap:'wrap'}}>
              <span style={{background:'var(--bg4)',color:'var(--accent2)',fontSize:12,fontWeight:600,padding:'3px 10px',borderRadius:6}}>{modalHoja.quarter} · {modalHoja.anio}</span>
              <Badge label={modalHoja.estado}/>
              {modalHoja.tieneCVC==='Sí'&&<span style={{background:'rgba(6,182,212,0.15)',color:'#06b6d4',fontSize:11,fontWeight:600,padding:'3px 10px',borderRadius:6}}>CVC</span>}
              {modalHoja.tienePromotora==='Sí'&&<span style={{background:'rgba(168,139,250,0.15)',color:'#a78bfa',fontSize:11,fontWeight:600,padding:'3px 10px',borderRadius:6}}>Promotora</span>}
            </div>
            {(modalHoja.planesItems||[]).filter(i=>i.nombre).length>0&&(
              <div>
                <div style={{fontSize:11,color:'var(--text3)',marginBottom:8,textTransform:'uppercase',letterSpacing:'0.05em'}}>Planes del {modalHoja.quarter}</div>
                <div style={{display:'flex',flexDirection:'column',gap:8}}>
                  {(modalHoja.planesItems||[]).filter(i=>i.nombre).map((item,i)=>(
                    <div key={i} style={{background:'var(--bg3)',borderRadius:10,padding:'12px 16px',border:'1px solid var(--border2)'}}>
                      <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:6}}>
                        <span style={{fontWeight:600,fontSize:13,color:'var(--accent2)'}}>{item.nombre}</span>
                        {item.tienePago==='Sí'&&<span style={{fontSize:11,background:'var(--green-soft)',color:'var(--green)',padding:'2px 8px',borderRadius:5,fontWeight:600}}>✓ Con pago</span>}
                      </div>
                      {item.descripcionPago&&<div style={{marginBottom:item.condiciones?8:0}}><span style={{fontSize:11,color:'var(--text3)',textTransform:'uppercase',letterSpacing:'0.04em'}}>Pago incluye: </span><span style={{fontSize:12,color:'var(--text)'}}>{item.descripcionPago}</span></div>}
                      {item.condiciones&&<div><span style={{fontSize:11,color:'var(--text3)',textTransform:'uppercase',letterSpacing:'0.04em'}}>Condiciones: </span><span style={{fontSize:12,color:'var(--text2)'}}>{item.condiciones}</span></div>}
                    </div>
                  ))}
                </div>
              </div>
            )}
            <div>
              <div style={{fontSize:11,color:'var(--text3)',marginBottom:8,textTransform:'uppercase',letterSpacing:'0.05em'}}>Metas por mes</div>
              <div style={{display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:8}}>
                {mesesQ(modalHoja.quarter).map((mes,i)=>(
                  <div key={mes} style={{background:'var(--bg3)',borderRadius:8,padding:'10px 12px'}}>
                    <div style={{fontSize:11,fontWeight:600,color:'var(--accent2)',marginBottom:4}}>{mes}</div>
                    <div style={{fontSize:12,color:'var(--text2)'}}>Venta: {cop(Number(modalHoja['metaMes'+(i+1)])||0)}</div>
                    <div style={{fontSize:12,color:'var(--text2)'}}>Galones: {num(Number(modalHoja['metaGalonesMes'+(i+1)])||0)}</div>
                  </div>
                ))}
              </div>
            </div>
            {modalHoja.tieneCVC==='Sí'&&<div style={{background:'rgba(6,182,212,0.08)',border:'1px solid rgba(6,182,212,0.2)',borderRadius:10,padding:'12px 16px'}}><div style={{fontSize:12,fontWeight:600,color:'#06b6d4',marginBottom:4}}>CVC</div><div style={{fontSize:13}}>{modalHoja.montoCVC?cop(Number(modalHoja.montoCVC)):''} {modalHoja.detalleCVC}</div></div>}
            {modalHoja.tienePromotora==='Sí'&&<div style={{background:'rgba(168,139,250,0.08)',border:'1px solid rgba(168,139,250,0.2)',borderRadius:10,padding:'12px 16px'}}><div style={{fontSize:12,fontWeight:600,color:'#a78bfa',marginBottom:4}}>Promotora</div><div style={{fontSize:13}}>{modalHoja.nombrePromotora}{modalHoja.montoPromotora&&' — '+cop(Number(modalHoja.montoPromotora))+'/mes'}</div></div>}
            {(modalHoja.condicionesGenerales||modalHoja.condiciones)&&<div><div style={{fontSize:11,color:'var(--text3)',marginBottom:5,textTransform:'uppercase',letterSpacing:'0.05em'}}>Condiciones generales</div><div style={{background:'var(--bg3)',borderRadius:10,padding:'12px 16px',fontSize:13,lineHeight:1.6}}>{modalHoja.condicionesGenerales||modalHoja.condiciones}</div></div>}
            {modalHoja.acuerdos&&<div><div style={{fontSize:11,color:'var(--text3)',marginBottom:5,textTransform:'uppercase',letterSpacing:'0.05em'}}>Acuerdos especiales</div><div style={{background:'var(--green-soft)',border:'1px solid rgba(61,214,140,0.2)',borderRadius:10,padding:'12px 16px',fontSize:13,lineHeight:1.6}}>{modalHoja.acuerdos}</div></div>}
            {modalHoja.notas&&<div><div style={{fontSize:11,color:'var(--text3)',marginBottom:5,textTransform:'uppercase',letterSpacing:'0.05em'}}>Notas internas</div><div style={{background:'var(--yellow-soft)',border:'1px solid rgba(255,209,102,0.2)',borderRadius:10,padding:'12px 16px',fontSize:13,lineHeight:1.6}}>{modalHoja.notas}</div></div>}
            <div>
              <div style={{fontSize:11,color:'var(--text3)',marginBottom:8,textTransform:'uppercase',letterSpacing:'0.05em'}}>Historial de seguimiento</div>
              <div style={{display:'flex',flexDirection:'column',gap:7,marginBottom:10}}>
                {(modalHoja.historial||[]).length===0&&<div style={{fontSize:12,color:'var(--text3)',fontStyle:'italic'}}>Sin notas aún</div>}
                {(modalHoja.historial||[]).map((n,i)=>(
                  <div key={i} style={{display:'flex',gap:10,padding:'9px 12px',background:'var(--bg3)',borderRadius:8}}>
                    <span style={{fontSize:11,color:'var(--text3)',whiteSpace:'nowrap'}}>{n.fecha}</span>
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
    const nd={...data,pendientes}; setData(nd); save(nd,editId?{}:{insertar:entry,tipo:'pendientes'}); setModal(false); setEditId(null); setForm(blank)
  }
  const del = id => { const nd={...data,pendientes:data.pendientes.filter(p=>p.id!==id)}; setData(nd); save(nd,{eliminar:id,tipo:'pendientes'}) }
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
                {p.fechaLimite&&<span>📅 {String(p.fechaLimite).slice(0,10)}</span>}
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
            <Field label="Categoría"><input value={form.categoria} onChange={e=>setForm({...form,categoria:e.target.value})} placeholder="Ej: Seguimiento..."/></Field>
            <Field label="Fecha límite"><input type="date" value={form.fechaLimite} onChange={e=>setForm({...form,fechaLimite:e.target.value})}/></Field>
            <Field label="Prioridad"><select value={form.prioridad} onChange={e=>setForm({...form,prioridad:e.target.value})}>{PRIORIDADES.map(p=><option key={p}>{p}</option>)}</select></Field>
            <Field label="Estado"><select value={form.estado} onChange={e=>setForm({...form,estado:e.target.value})}>{ESTADOS_PEND.map(s=><option key={s}>{s}</option>)}</select></Field>
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
// APOYO CIERRE
// ═══════════════════════════════════════════════════════
function ApoyoCierre({ data, setData, usuario }) {
  const [mesActivo, setMesActivo] = useState('')
  const [comercialDetalle, setComercialDetalle] = useState(null)
  const [comercialActivo, setComercialActivo] = useState(null)
  const [modalAsignacion, setModalAsignacion] = useState(false)
  const [modalRedencion, setModalRedencion] = useState(false)
  const [enviandoEmail, setEnviandoEmail] = useState(false)
  const [formAsig, setFormAsig] = useState({ comercial:'', distribuidor:'', mes:'', anio:2026, monto:'' })
  const [formRed, setFormRed] = useState({ cliente:'', producto:'', valor:'', notas:'', copiarJefe:false })

  const SHEETS_URL = 'https://script.google.com/macros/s/AKfycbxzqhKZr7ue545BDwpgqQrNsognsVOtxxZNPd6P4Jm_wtdj1_Sqkd9FdswrdtfM2d8S/exec'

  const _canalActual = usuario?.id || 'distribucion'
  const EMAILS = EMAILS_POR_CANAL[_canalActual] || EMAILS_POR_CANAL.distribucion
  const JEFES_CANAL = JEFES_POR_CANAL[_canalActual] || {}
  const EMAIL_JEFE = 'cgil@prolub.com.co'

  // Emails por unidad para notificaciones de tareas
  const EMAILS_UNIDAD = {
    industria:    'epabon@gulfcolombia.com',
    distribucion: 'msilva@prolub.com.co',
    zonas:        'msilva@prolub.com.co',
    diseno:       'aalvarado@gulfcolombia.com',
    lider:        'cgil@prolub.com.co',
  }

  const apoyos = data.apoyoCierre || []
  const redenciones = data.redenciones || []
  const distribuidores = [...new Set(data.inversiones.map(i=>i.distribuidor))].sort()
  const comerciales = [...new Set(apoyos.map(a=>a.comercial))].sort()
  const apoyosFiltrados = mesActivo ? apoyos.filter(a=>a.mes===mesActivo) : apoyos

  const getSaldo = (comercial, mes) => {
    const asig = apoyos.filter(a=>a.comercial===comercial&&(!mes||a.mes===mes)).reduce((s,a)=>s+(Number(a.monto)||0),0)
    const redim = redenciones.filter(r=>r.comercial===comercial&&(!mes||r.mes===mes)).reduce((s,r)=>s+(Number(r.valor)||0),0)
    return { asignado:asig, redimido:redim, saldo:asig-redim }
  }

  const guardarAsignacion = () => {
    if(!formAsig.comercial||!formAsig.mes||!formAsig.monto) return
    const nueva = { id:Date.now(), ...formAsig, monto:Number(formAsig.monto), anio:Number(formAsig.anio) }
    const nd = {...data, apoyoCierre:[...apoyos, nueva]}
    setData(nd); save(nd)
    setModalAsignacion(false)
    setFormAsig({ comercial:'', distribuidor:'', mes:'', anio:2026, monto:'' })
  }

  const guardarRedencion = async () => {
    if(!comercialActivo||!formRed.producto||!formRed.valor) return
    const saldoActual = getSaldo(comercialActivo.comercial, mesActivo||comercialActivo.mes)
    const nuevaSaldo = saldoActual.saldo - Number(formRed.valor)
    const nueva = {
      id: Date.now(),
      comercial: comercialActivo.comercial,
      distribuidor: comercialActivo.distribuidor,
      mes: mesActivo||comercialActivo.mes||'',
      anio: comercialActivo.anio||2026,
      cliente: formRed.cliente,
      producto: formRed.producto,
      valor: Number(formRed.valor),
      notas: formRed.notas,
      fecha: new Date().toLocaleDateString('es-CO'),
    }
    const nd = {...data, redenciones:[...redenciones, nueva]}
    setData(nd); save(nd)
    setModalRedencion(false)
    setFormRed({ cliente:'', producto:'', valor:'', notas:'', copiarJefe:false })

    const emailDest = EMAILS[comercialActivo.comercial]
    if(emailDest) {
      setEnviandoEmail(true)
      try {
        enviarViaJsonp(SHEETS_URL, {
            accion: 'enviar_correo',
            destinatario: emailDest,
            copia: EMAIL_JEFE,
            copiaJefe: formRed.copiarJefe ? Object.values(JEFES_CANAL).join(',') : '',
            comercial: comercialActivo.comercial,
            cliente: formRed.cliente,
            producto: formRed.producto,
            valorRedimido: Number(formRed.valor),
            nuevoSaldo: nuevaSaldo,
            mes: mesActivo||comercialActivo.mes||'',
            fecha: nueva.fecha,
            notas: formRed.notas,
          })
      } catch(e) { console.error('Error enviando email:', e) }
      setEnviandoEmail(false)
    }
  }

  const delAsig = id => { const nd={...data,apoyoCierre:apoyos.filter(a=>a.id!==id)};setData(nd);save(nd) }
  const delRed = id => { const nd={...data,redenciones:redenciones.filter(r=>r.id!==id)};setData(nd);save(nd) }

  const resumenComerciales = comerciales.map(c=>{
    const asigRows = mesActivo ? apoyos.filter(a=>a.comercial===c&&a.mes===mesActivo) : apoyos.filter(a=>a.comercial===c)
    const redsRows = mesActivo ? redenciones.filter(r=>r.comercial===c&&r.mes===mesActivo) : redenciones.filter(r=>r.comercial===c)
    const meses_c = [...new Set(asigRows.map(a=>a.mes))]
    const totalAsig = asigRows.reduce((s,a)=>s+(Number(a.monto)||0),0)
    const totalRedim = redsRows.reduce((s,r)=>s+(Number(r.valor)||0),0)
    return { comercial:c, meses:meses_c, totalAsig, totalRedim, saldo:totalAsig-totalRedim, redsRows, asigRows }
  }).filter(c=>!mesActivo||c.asigRows.length>0)

  const totalAsig = apoyosFiltrados.reduce((s,a)=>s+(Number(a.monto)||0),0)
  const totalRedim = redenciones.filter(r=>!mesActivo||r.mes===mesActivo).reduce((s,r)=>s+(Number(r.valor)||0),0)

  return (
    <div style={{display:'flex',flexDirection:'column',gap:18}}>
      {enviandoEmail&&<div style={{position:'fixed',top:70,right:24,background:'var(--accent)',color:'#fff',padding:'10px 18px',borderRadius:10,fontSize:13,zIndex:400}}>✉️ Enviando correo...</div>}
      <div style={{display:'flex',gap:10,alignItems:'center',flexWrap:'wrap'}}>
        <select value={mesActivo} onChange={e=>setMesActivo(e.target.value)} style={{width:160}}>
          <option value="">Todos los meses</option>
          {MESES.map(m=><option key={m}>{m}</option>)}
        </select>
        {mesActivo&&<button onClick={()=>setMesActivo('')} style={{...S.btn('var(--bg3)','var(--text2)'),padding:'5px 10px',fontSize:12}}>✕</button>}
        <button onClick={()=>setModalAsignacion(true)} style={{...S.btn('var(--accent)','#fff'),marginLeft:'auto'}}>
          <PlusCircle size={15}/> Asignar apoyo
        </button>
      </div>
      <div style={{display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:12}}>
        <KpiCard icon={DollarSign} label="Total asignado" value={cop(totalAsig)} sub={mesActivo||'Todos los meses'}/>
        <KpiCard icon={TrendingUp} label="Total redimido" value={cop(totalRedim)} accent="var(--orange)"/>
        <KpiCard icon={BarChart2} label="Saldo disponible" value={cop(totalAsig-totalRedim)} accent={totalAsig-totalRedim<0?'var(--red)':'var(--green)'}/>
      </div>
      <div style={{display:'flex',flexDirection:'column',gap:12}}>
        {resumenComerciales.length===0&&(
          <div style={{...S.card,padding:40,textAlign:'center',color:'var(--text3)'}}>No hay apoyos registrados. Usa <strong>Asignar apoyo</strong>.</div>
        )}
        {resumenComerciales.map((c,i)=>{
          const pct = c.totalAsig>0?(c.totalRedim/c.totalAsig)*100:0
          const expandido = comercialDetalle===c.comercial
          return (
            <div key={i} style={{...S.card}}>
              <div style={{padding:'14px 20px',display:'flex',alignItems:'center',gap:14,cursor:'pointer'}} onClick={()=>setComercialDetalle(expandido?null:c.comercial)}>
                <div style={{width:40,height:40,borderRadius:10,background:'var(--accent-soft)',display:'flex',alignItems:'center',justifyContent:'center',fontSize:16,fontWeight:700,color:'var(--accent2)',flexShrink:0}}>
                  {c.comercial[0]}
                </div>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{display:'flex',alignItems:'center',gap:10,marginBottom:4}}>
                    <span style={{fontWeight:700,fontSize:15}}>{c.comercial}</span>
                    {EMAILS[c.comercial]&&<span style={{fontSize:11,color:'var(--text3)'}}>{EMAILS[c.comercial]}</span>}
                    <span style={{fontSize:11,background:'var(--bg4)',color:'var(--text2)',padding:'2px 8px',borderRadius:5}}>{c.meses.join(' · ')}</span>
                  </div>
                  <div style={{display:'flex',alignItems:'center',gap:10}}>
                    <div style={{flex:1,height:6,background:'var(--bg4)',borderRadius:3}}>
                      <div style={{width:Math.min(pct,100)+'%',height:'100%',background:pct>100?'var(--red)':pct>80?'var(--yellow)':'var(--orange)',borderRadius:3}}/>
                    </div>
                    <span style={{fontSize:12,color:'var(--text2)',whiteSpace:'nowrap'}}>
                      <strong style={{color:'var(--orange)',fontFamily:'var(--mono)'}}>{cop(c.totalRedim)}</strong>
                      <span style={{color:'var(--text3)'}}> / {cop(c.totalAsig)}</span>
                      <span style={{marginLeft:8,color:c.saldo<0?'var(--red)':'var(--green)',fontFamily:'var(--mono)',fontWeight:600}}> Saldo: {cop(c.saldo)}</span>
                    </span>
                  </div>
                </div>
                <div style={{display:'flex',gap:8,alignItems:'center'}}>
                  <button onClick={e=>{e.stopPropagation();
                    setComercialActivo({comercial:c.comercial,distribuidor:'',mes:mesActivo||c.meses[0]||'',anio:2026})
                    setModalRedencion(true)
                  }} style={{...S.btn('var(--orange)','#fff'),fontSize:12,padding:'6px 14px'}}>
                    + Redención
                  </button>
                  <span style={{fontSize:14,color:'var(--text3)',transform:expandido?'rotate(180deg)':'rotate(0deg)'}}>▼</span>
                </div>
              </div>
              {expandido&&(
                <div style={{borderTop:'1px solid var(--border)',padding:'16px 20px',display:'flex',flexDirection:'column',gap:14}}>
                  <div>
                    <div style={{fontSize:10,color:'var(--text3)',textTransform:'uppercase',letterSpacing:'0.05em',marginBottom:8}}>Asignaciones</div>
                    <div style={{display:'flex',flexDirection:'column',gap:5}}>
                      {c.asigRows.map((a,j)=>(
                        <div key={j} style={{display:'flex',justifyContent:'space-between',alignItems:'center',padding:'7px 12px',background:'var(--bg3)',borderRadius:8}}>
                          <div>
                            <span style={{fontSize:12,fontWeight:500}}>{a.mes} {a.anio}</span>
                            {a.distribuidor&&<span style={{fontSize:11,color:'var(--text3)',marginLeft:8}}>→ {a.distribuidor}</span>}
                          </div>
                          <div style={{display:'flex',alignItems:'center',gap:8}}>
                            <span style={{fontFamily:'var(--mono)',fontWeight:600,color:'var(--accent2)'}}>{cop(a.monto)}</span>
                            <button onClick={()=>{if(window.confirm('¿Eliminar esta asignación de $'+a.monto.toLocaleString('es-CO')+'?'))delAsig(a.id)}} style={{...S.btn('var(--red)','#fff'),padding:'5px 10px',fontSize:12,fontWeight:600}}>🗑️ Eliminar</button>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                  {c.redsRows.length>0&&(
                    <div>
                      <div style={{fontSize:10,color:'var(--text3)',textTransform:'uppercase',letterSpacing:'0.05em',marginBottom:8}}>Redenciones por cliente</div>
                      {[...new Set(c.redsRows.map(r=>r.cliente||r.distribuidor||'Sin cliente'))].map(cliente=>{
                        const redsCliente = c.redsRows.filter(r=>(r.cliente||r.distribuidor||'Sin cliente')===cliente)
                        const totalCliente = redsCliente.reduce((s,r)=>s+(Number(r.valor)||0),0)
                        return (
                          <div key={cliente} style={{marginBottom:10,background:'var(--bg3)',borderRadius:10,overflow:'hidden'}}>
                            <div style={{padding:'8px 14px',display:'flex',justifyContent:'space-between',alignItems:'center',borderBottom:'1px solid var(--border)'}}>
                              <span style={{fontSize:12,fontWeight:600,color:'var(--accent2)'}}>👤 {cliente}</span>
                              <span style={{fontFamily:'var(--mono)',fontSize:12,fontWeight:600,color:'var(--orange)'}}>{cop(totalCliente)}</span>
                            </div>
                            {redsCliente.map((r,j)=>(
                              <div key={j} style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',padding:'7px 14px',borderBottom:j<redsCliente.length-1?'1px solid var(--border)':'none'}}>
                                <div style={{flex:1,minWidth:0}}>
                                  <div style={{fontSize:12,fontWeight:500}}>{r.producto}</div>
                                  <div style={{fontSize:11,color:'var(--text3)'}}>{r.fecha} {r.mes&&'· '+r.mes}</div>
                                  {r.notas&&<div style={{fontSize:11,color:'var(--text3)'}}>{r.notas}</div>}
                                </div>
                                <div style={{display:'flex',alignItems:'center',gap:7,flexShrink:0,marginLeft:10}}>
                                  <span style={{fontFamily:'var(--mono)',fontWeight:600,color:'var(--orange)'}}>{cop(r.valor)}</span>
                                  <button onClick={()=>delRed(r.id)} style={{...S.btn('var(--red-soft)','var(--red)'),padding:'3px 6px'}}><Trash2 size={11}/></button>
                                </div>
                              </div>
                            ))}
                          </div>
                        )
                      })}
                    </div>
                  )}
                  {c.redsRows.length===0&&<div style={{fontSize:12,color:'var(--text3)',fontStyle:'italic'}}>Sin redenciones aún</div>}
                </div>
              )}
            </div>
          )
        })}
      </div>

      {modalAsignacion&&(
        <Modal title="Asignar apoyo de cierre" onClose={()=>setModalAsignacion(false)}>
          <div style={{display:'flex',flexDirection:'column',gap:14}}>
            <Field label="Comercial *">
              <input list="comerciales-list" value={formAsig.comercial} onChange={e=>setFormAsig({...formAsig,comercial:e.target.value})} placeholder="Nombre del comercial"/>
              <datalist id="comerciales-list">{comerciales.map(c=><option key={c} value={c}/>)}</datalist>
            </Field>
            <Field label="Distribuidor">
              <input list="dist-apoyo" value={formAsig.distribuidor} onChange={e=>setFormAsig({...formAsig,distribuidor:e.target.value})} placeholder="Ej: LUBRICAFE S.A.S."/>
              <datalist id="dist-apoyo">{distribuidores.map(d=><option key={d} value={d}/>)}</datalist>
            </Field>
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:12}}>
              <Field label="Mes *"><select value={formAsig.mes} onChange={e=>setFormAsig({...formAsig,mes:e.target.value})}><option value="">Selecciona...</option>{MESES.map(m=><option key={m}>{m}</option>)}</select></Field>
              <Field label="Año"><input type="number" value={formAsig.anio} onChange={e=>setFormAsig({...formAsig,anio:e.target.value})} placeholder="2026"/></Field>
            </div>
            <Field label="Monto asignado (COP) *"><input type="number" value={formAsig.monto} onChange={e=>setFormAsig({...formAsig,monto:e.target.value})} placeholder="Ej: 1200000"/></Field>
          </div>
          <div style={{display:'flex',gap:10,justifyContent:'flex-end',marginTop:20}}>
            <button onClick={()=>setModalAsignacion(false)} style={S.btn('var(--bg3)','var(--text2)')}>Cancelar</button>
            <button onClick={guardarAsignacion} style={S.btn('var(--accent)','#fff')}><Check size={15}/> Guardar</button>
          </div>
        </Modal>
      )}

      {modalRedencion&&comercialActivo&&(
        <Modal title={'Registrar redención — '+comercialActivo.comercial} onClose={()=>setModalRedencion(false)}>
          <div style={{padding:'10px 14px',background:'var(--bg3)',borderRadius:8,marginBottom:14,fontSize:12,color:'var(--text2)',display:'flex',alignItems:'center',gap:10,flexWrap:'wrap'}}>
            <span>Comercial: <strong>{comercialActivo.comercial}</strong></span>
            {(mesActivo||comercialActivo.mes)&&<span>· Mes: <strong>{mesActivo||comercialActivo.mes}</strong></span>}
            {(()=>{ const s = getSaldo(comercialActivo.comercial, mesActivo||comercialActivo.mes); return s.asignado>0 ? <span>· Saldo: <strong style={{color:s.saldo<0?'var(--red)':'var(--green)'}}>{cop(s.saldo)}</strong></span> : null })()}
            {EMAILS[comercialActivo.comercial]&&<span style={{color:'var(--text3)'}}>· ✉️ {EMAILS[comercialActivo.comercial]}</span>}
          </div>
          <div style={{display:'flex',flexDirection:'column',gap:14}}>
            <Field label="Cliente que redime *">
              <input list="clientes-red" value={formRed.cliente} onChange={e=>setFormRed({...formRed,cliente:e.target.value})} placeholder="Ej: LUBRICAFE S.A.S...."/>
              <datalist id="clientes-red">{distribuidores.map(d=><option key={d} value={d}/>)}</datalist>
            </Field>
            <Field label="Producto / Concepto redimido *"><input value={formRed.producto} onChange={e=>setFormRed({...formRed,producto:e.target.value})} placeholder="Ej: Producto Plan Expreso..."/></Field>
            <Field label="Valor redimido (COP) *"><input type="number" value={formRed.valor} onChange={e=>setFormRed({...formRed,valor:e.target.value})} placeholder="0"/></Field>
            <Field label="Notas"><textarea value={formRed.notas} onChange={e=>setFormRed({...formRed,notas:e.target.value})} rows={2} style={{resize:'vertical'}} placeholder="Observaciones..."/></Field>
          </div>
          <div style={{background:'rgba(255,159,67,0.08)',border:'1px solid rgba(255,159,67,0.2)',borderRadius:8,padding:'8px 12px',fontSize:11,color:'var(--orange)',marginTop:14}}>
            ✉️ Al registrar se enviará correo automático a <strong>{EMAILS[comercialActivo.comercial]||'comercial'}</strong> con copia a Paola
          </div>
          <div style={{display:'flex',gap:10,justifyContent:'flex-end',marginTop:14}}>
            <button onClick={()=>setModalRedencion(false)} style={S.btn('var(--bg3)','var(--text2)')}>Cancelar</button>
            <button onClick={guardarRedencion} style={S.btn('var(--orange)','#fff')}><Check size={15}/> Registrar y enviar correo</button>
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
  const ws3 = XLSX.utils.json_to_sheet(data.presupuestos.map(p=>{ const inv=data.inversiones.filter(i=>normMes(i.mes)===normMes(p.mes)&&i.anio===p.anio).reduce((s,i)=>s+(i.inversion||0),0); return {'Año':p.anio,Mes:p.mes,'Presupuesto COP':p.monto,'Invertido COP':inv,'Diferencia':inv-p.monto,'% Ejec':p.monto>0?((inv/p.monto)*100).toFixed(1)+'%':'0%'} }))
  ws3['!cols']=[8,12,16,14,14,10].map(w=>({wch:w}))
  XLSX.utils.book_append_sheet(wb,ws3,'Presupuesto')
  const ws4 = XLSX.utils.json_to_sheet(data.planes.map(p=>({ Distribuidor:p.distribuidor, 'Año':p.anio, Quarter:p.quarter, Estado:p.estado, 'Tipos de Plan':(p.tiposPlan||[]).join(', '), 'Meta Galones':p.metaGalones, 'Meta Venta COP':p.metaVenta, Condiciones:p.condiciones, Acuerdos:p.acuerdos, Notas:p.notas })))
  ws4['!cols']=[25,8,8,15,30,14,16,40,40,30].map(w=>({wch:w}))
  XLSX.utils.book_append_sheet(wb,ws4,'Planes')
  const ws5 = XLSX.utils.json_to_sheet(data.pendientes.map(p=>({ Distribuidor:p.distribuidor, Tarea:p.tarea, Categoria:p.categoria, 'Fecha Limite':p.fechaLimite, Prioridad:p.prioridad, Estado:p.estado, Responsable:p.responsable, Notas:p.notas })))
  ws5['!cols']=[22,35,16,13,10,12,16,28].map(w=>({wch:w}))
  XLSX.utils.book_append_sheet(wb,ws5,'Pendientes')
  const ws6 = XLSX.utils.json_to_sheet((data.gastosPresupuesto||[]).filter(g=>g.mes||g.gasto).map((g,i)=>({'id':g.id||Date.now()+i,'año':g.anio||2026,'Mes':g.mes||'','Gasto (Nom. Producto, Cliente)':g.gasto||'','Valor Factura':Number(g.valorFactura)||0,'Canal':g.canal||'','Observación (ATJ)':g.observacion||'','Estado':g.estado||'Pendiente','Centro de Costos':g.centroCostos||'','Orden de Compra':g.ordenCompra||'','NIT Proveedor':g.nitProveedor||'','Nom. Proveedor':g.nomProveedor||'','Doc. Cargado':g.docCargado||'','Obs. Katherine':g.obsKatherine||'','notas':g.notas||''})))
  ws6['!cols']=[14,8,12,35,16,10,25,12,16,14,14,18,10,18,20].map(w=>({wch:w}))
  XLSX.utils.book_append_sheet(wb,ws6,'GASTOS_PRESUPUESTO')
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
      const esInv   = cols0.some(c=>c.toLowerCase().includes('distribuidor'))
      const esGasto = cols0.some(c=>c.toLowerCase().includes('gasto')||c.toLowerCase().includes('nom. producto')||c.toLowerCase().includes('valor factura'))

      const invRows = leer('Inversiones') || (esInv && !esGasto ? primeraRows : null)
      if(invRows?.length) {
        const nuevas=invRows.filter(r=>r['Distribuidor']||r['distribuidor']).map((r,i)=>({
          id:Date.now()+i, fecha:String(r['Fecha']||''),
          anio:Number(r['Año']||r['Ano']||r['año']||2026), mes:normMes(String(r['Mes']||'').trim()),
          distribuidor:String(r['Distribuidor']||r['distribuidor']||'').trim(),
          tipoPlan:String(r['Tipo Plan']||'').trim(), concepto:String(r['Concepto']||'').trim(),
          inversion:parseN(r['Inversión COP']||r['Inversion COP']||r['Inversión']||r['Inversion']||0),
          galonesPlan:r['Galones Plan']?parseN(r['Galones Plan']):'',
          notas:String(r['Notas']||'').trim(),
        }))
        if(nuevas.length){data={...data,inversiones:[...data.inversiones,...nuevas]};importados.Inversiones=nuevas.length}
      }
      const ventRows = leer('Ventas') || leer('Hoja2')
      if(ventRows?.length) {
        const nuevas=ventRows.filter(r=>r['Distribuidor']||r['distribuidor']).map((r,i)=>({
          id:Date.now()+10000+i,
          anio:Number(r['Año']||r['Ano']||r['año']||2026),
          mes:normMes(String(r['Mes']||r['mes']||'').trim()),
          distribuidor:String(r['Distribuidor']||r['distribuidor']||'').trim(),
          galones:parseN(r['Galones']||r['galones']||0),
          ventaNeta:parseN(r['Venta Neta']||r['Venta Neta COP']||r['ventaNeta']||0),
          notas:'',
        }))
        if(nuevas.length){data={...data,ventas:[...data.ventas,...nuevas]};importados.Ventas=nuevas.length}
      }
      const presRows = leer('Presupuesto')
      if(presRows?.length) {
        const nuevas=presRows.filter(r=>r['Mes']||r['mes']).map((r,i)=>({
          id:Date.now()+20000+i, anio:Number(r['Año']||r['Ano']||2026),
          mes:normMes(String(r['Mes']||r['mes']||'').trim()), monto:parseN(r['Presupuesto COP']||r['Presupuesto']||0),
        }))
        if(nuevas.length){data={...data,presupuestos:[...data.presupuestos,...nuevas]};importados.Presupuesto=nuevas.length}
      }
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
      const gastosRows = leer('Presupuesto Mercadeo') || leer('Presupuesto Gastos') || leer('Control Presupuesto') || (esGasto ? primeraRows : null)
      if(gastosRows?.length) {
        const nuevas=gastosRows
          .filter(r=>(r['Gasto (Nom. Producto, Cliente)']||r['Gasto'])&&r['Gasto (Nom. Producto, Cliente)']!=='Gasto (Nom. Producto, Cliente)')
          .map((r,i)=>{
            const docCargado = r['\u00bfDocumento Cargado a la Carpeta?']||r['Documento Cargado']
            const estadoRaw = String(r['Estado']||'')
            let estado = 'Pendiente'
            if(estadoRaw && !estadoRaw.startsWith('=IF') && !estadoRaw.startsWith('=if')) {
              estado = estadoRaw.trim()
            } else {
              estado = (docCargado===true||docCargado==='TRUE'||docCargado==='true'||docCargado==='Sí') ? 'Ingresado' : 'Pendiente'
            }
            const rawVal = r['Valor Factura']
            let valorFactura = 0
            if(typeof rawVal==='number') valorFactura = Math.abs(rawVal)
            else if(rawVal && !String(rawVal).startsWith('=')) valorFactura = parseN(rawVal)
            return {
              id: Date.now()+50000+i,
              anio: Number(r['Año']||r['Ano']||2026),
              mes: normMes(String(r['Mes']||'Sin mes').trim()),
              gasto: String(r['Gasto (Nom. Producto, Cliente)']||r['Gasto']||'').trim(),
              valorFactura,
              canal: String(r['Canal']||'').trim(),
              observacion: String(r['Observación (ATJ, otros)']||r['Observacion']||'').trim(),
              estado,
              centroCostos: String(r['Centro de Costos']||'').trim(),
              ordenCompra: String(r['Orden de Compra']||'').trim(),
              nitProveedor: String(r['Nit Proveedor']||'').trim(),
              nomProveedor: String(r['Nom. Proveedor']||'').trim(),
              docCargado: (docCargado===true||docCargado==='TRUE'||docCargado==='Sí') ? 'Sí' : 'No',
              obsKatherine: String(r['Observación Katherine']||r['Observacion Katherine']||'').trim(),
              notas: '',
            }
          }).filter(r=>r.gasto)
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
        const nueva={id:Date.now(),fecha:obj.fecha||new Date().toISOString().slice(0,10),anio:Number(obj.anio||new Date().getFullYear()),mes:normMes(obj.mes||''),distribuidor:obj.distribuidor||'',tipoPlan:obj.tipoPlan||'',concepto:obj.concepto||'',inversion:Number(obj.inversion||0),galonesPlan:obj.galonesPlan||'',notas:obj.notas||''}
        const nd={...currentData,inversiones:[...currentData.inversiones,nueva]}; setData(nd); save(nd)
        return '✅ Inversión registrada: '+nueva.distribuidor+' — '+cop(nueva.inversion)
      }
      if(obj.tipo==='crear_presupuesto'){
        const nuevo={id:Date.now(),anio:Number(obj.anio||new Date().getFullYear()),mes:normMes(obj.mes||''),monto:Number(obj.monto||0)}
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
// TAREAS ASIGNADAS — vista del canal
// ═══════════════════════════════════════════════════════
function TareasAsignadas({ data, setData, usuario }) {
  const tareas = data.tareasAsignadas || []
  const [cerrando,setCerrando] = useState(null)
  const [guardandoCierre,setGuardandoCierre] = useState(false)
  const [editando,setEditando] = useState(null)
  const [guardandoEdicion,setGuardandoEdicion] = useState(false)

  const abiertas = tareas.filter(t=>t.estado!=='Finalizada'&&t.estado!=='Cancelada')
  const cerradas = tareas.filter(t=>t.estado==='Finalizada'||t.estado==='Cancelada')

  const persistir = async (actualizada, anterior=null) => {
    const lista = tareas.map(t=>String(t.id)===String(actualizada.id)?actualizada:t)
    setData({...data,tareasAsignadas:lista})

    const uidAnterior = anterior ? obtenerUidResponsable(anterior.responsable) : uidPersonaPorUsuario(usuario)
    const uidNuevo = obtenerUidResponsable(actualizada.responsable) || uidPersonaPorUsuario(usuario)

    await eliminarFilaSheets('distribucion','tareasAsignadas',actualizada.id)
    await insertarFilaSheets('distribucion','tareasAsignadas',actualizada)

    if(uidAnterior && uidAnterior!=='distribucion' && uidAnterior!==uidNuevo) {
      await eliminarFilaSheets(uidAnterior,'tareasAsignadas',actualizada.id)
    }

    if(uidNuevo && uidNuevo!=='distribucion') {
      await eliminarFilaSheets(uidNuevo,'tareasAsignadas',actualizada.id)
      await insertarFilaSheets(uidNuevo,'tareasAsignadas',actualizada)
    }
  }

  const cambiarEstado = async (t,estado) => {
    if(estado==='Finalizada'&&t.estado!=='Finalizada') {
      setCerrando(t)
      return
    }
    try { await persistir({...t,estado},t) }
    catch(e){ console.error(e); alert('No se pudo actualizar la tarea.') }
  }

  const confirmarFinalizacion = async tiempo => {
    if(!cerrando||guardandoCierre) return
    setGuardandoCierre(true)
    const actualizada = {...cerrando,estado:'Finalizada',fechaFinalizacion:new Date().toLocaleString('es-CO'),...tiempo}
    try {
      await persistir(actualizada,cerrando)
      await enviarCorreoFinalizacion({
        tipo:'tarea_finalizada',
        proyecto:{nombre:'Tarea asignada'},
        subtarea:{nombre:actualizada.nombre,responsable:actualizada.responsable,descripcion:actualizada.descripcion},
        destinatarios:[obtenerCorreoResponsable(actualizada.responsable)],
      })
      setCerrando(null)
    } catch(e) {
      console.error(e)
      alert('No se pudo finalizar la tarea.')
    } finally {
      setGuardandoCierre(false)
    }
  }

  const guardarEdicion = async form => {
    if(!form.nombre) return alert('La tarea debe tener nombre.')
    setGuardandoEdicion(true)
    try {
      const anterior = editando
      const actualizada = {...anterior,...form}
      await persistir(actualizada,anterior)

      if(normalizarNombreResponsable(anterior.responsable)!==normalizarNombreResponsable(actualizada.responsable)) {
        const email = obtenerCorreoResponsable(actualizada.responsable)
        if(email) {
          await enviarViaJsonp(SHEETS_CONFIG.distribucion.url,{
            accion:'enviar_correo_tarea',
            destinatario:email,
            copia:EMAIL_LIDER_PROY,
            comercial:actualizada.responsable,
            responsable:actualizada.responsable,
            tarea:actualizada.nombre,
            producto:actualizada.nombre,
            notas:actualizada.descripcion||'',
            descripcion:actualizada.descripcion||'',
            fechaLimite:actualizada.fechaLimite||'',
            mes:actualizada.fechaLimite||'',
            fecha:new Date().toLocaleDateString('es-CO'),
          })
        }
      }
      setEditando(null)
    } catch(e) {
      console.error(e)
      alert('No se pudo guardar la edición.')
    } finally {
      setGuardandoEdicion(false)
    }
  }

  const Tarjeta = ({t}) => (
    <div style={{...S.card,padding:'15px 18px',display:'flex',gap:14,alignItems:'flex-start'}}>
      <div style={{flex:1,minWidth:0}}>
        <div style={{display:'flex',gap:8,alignItems:'center',flexWrap:'wrap',marginBottom:6}}>
          <span style={{fontWeight:700,fontSize:14}}>{t.nombre}</span>
          <span style={{fontSize:10,padding:'2px 8px',borderRadius:6,background:'var(--accent-soft)',color:'var(--accent2)'}}>Asignada por Paola</span>
        </div>
        {t.descripcion&&<p style={{fontSize:12,color:'var(--text2)',margin:'0 0 7px'}}>{t.descripcion}</p>}
        <div style={{display:'flex',gap:12,flexWrap:'wrap',fontSize:11,color:'var(--text3)'}}>
          {t.fechaLimite&&<span>📅 Límite: {String(t.fechaLimite).slice(0,10)}</span>}
          {t.areas&&<span>👥 {Array.isArray(t.areas)?t.areas.join(', '):t.areas}</span>}
          {t.fechaFinalizacion&&<span>✅ {t.fechaFinalizacion}</span>}
          {Number(t.tiempoMinutos)>0&&<span>⏱️ {formatoTiempoTarea(t)}</span>}
        </div>
      </div>
      <div style={{display:'flex',gap:6,alignItems:'center'}}>
        <select value={t.estado||'Pendiente'} onChange={e=>cambiarEstado(t,e.target.value)} style={{fontSize:11,padding:'5px 8px'}}>
          {['Pendiente','En proceso','Finalizada','Cancelada'].map(s=><option key={s}>{s}</option>)}
        </select>
        <button onClick={()=>setEditando(t)} style={{...S.btn('var(--accent-soft)','var(--accent2)'),padding:'5px 7px'}}><Edit2 size={12}/></button>
      </div>
    </div>
  )

  return (
    <div style={{display:'flex',flexDirection:'column',gap:16}}>
      <div style={{display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:12}}>
        <KpiCard icon={ListTodo} label="Asignadas" value={tareas.length} accent="var(--accent2)"/>
        <KpiCard icon={ListTodo} label="Abiertas" value={abiertas.length} accent="var(--yellow)"/>
        <KpiCard icon={Check} label="Finalizadas" value={cerradas.filter(t=>t.estado==='Finalizada').length} accent="var(--green)"/>
      </div>
      {tareas.length===0&&<div style={{...S.card,padding:44,textAlign:'center',color:'var(--text3)'}}>Sin tareas asignadas</div>}
      {abiertas.map(t=><Tarjeta key={t.id} t={t}/>)}
      {cerradas.map(t=><Tarjeta key={t.id} t={t}/>)}
      {cerrando&&<ModalTiempoTarea tarea={cerrando} onClose={()=>setCerrando(null)} onConfirm={confirmarFinalizacion} guardando={guardandoCierre}/>}
      {editando&&<ModalEditarTarea tarea={editando} onClose={()=>setEditando(null)} onSave={guardarEdicion} guardando={guardandoEdicion} titulo="Editar tarea asignada"/>}
    </div>
  )
}


// ═══════════════════════════════════════════════════════
// PROYECTOS
// ═══════════════════════════════════════════════════════
const EMAILS_EQUIPO = {
  'Emma':    'epabon@gulfcolombia.com',
  'María P': 'msilva@prolub.com.co',
  'Maria P': 'msilva@prolub.com.co',
  'Juan':    'acomercial@prolub.com.co',
  'Camilo':  'aalvarado@gulfcolombia.com',
  'Julian':  'jacifuentes@prolub.com.co',
  'Julián':  'jacifuentes@prolub.com.co',
}

const EMAIL_LIDER_PROY = 'cgil@prolub.com.co'
const AREAS = ['Compras','Facturación','Logística','Comercial']
const RESPONSABLES = ['Emma','María P','Juan','Camilo','Julian']

// Archivos personales de seguimiento.
// El Apps Script puede usar estos IDs para copiar tareas y proyectos
// al archivo individual del responsable.
const SHEETS_RESPONSABLES = {
  'Juan':   '1s6fkyyCihzOki9Fan9EKIIvxwlra-QizwI7fy210GUI',
  'Camilo': '1jVKNc8krqeClgfzaXr-i7YIHwuYRmjkoU35NEjz4oDk',
  'Julian': '1nEdGAcy45hMckYxOC_SRfrYEqW6ifgRv9fFoJA75YlY',
  'Julián': '1nEdGAcy45hMckYxOC_SRfrYEqW6ifgRv9fFoJA75YlY',
}

const SHEETS_URL_PROY = SHEETS_CONFIG.distribucion.url

function normalizarNombreResponsable(nombre='') {
  return String(nombre)
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g,'')
    .toLowerCase()
}

function obtenerCorreoResponsable(nombre='') {
  const normalizado = normalizarNombreResponsable(nombre)
  const mapa = {
    'emma': 'epabon@gulfcolombia.com',
    'maria p': 'msilva@prolub.com.co',
    'juan': 'acomercial@prolub.com.co',
    'camilo': 'aalvarado@gulfcolombia.com',
    'julian': 'jacifuentes@prolub.com.co',
  }
  return mapa[normalizado] || EMAILS_EQUIPO[nombre] || ''
}

function obtenerUidResponsable(nombre='') {
  const normalizado = normalizarNombreResponsable(nombre)
  const mapa = {
    'maria p': 'distribucion',
    'emma': 'industria',
    'juan': 'juan',
    'camilo': 'camilo',
    'julian': 'julian',
  }
  return mapa[normalizado] || 'distribucion'
}

const ESTADO_PROY = {
  'Pendiente': { color:'var(--yellow)',  bg:'rgba(180,83,9,0.1)' },
  'En proceso':{ color:'var(--accent2)', bg:'var(--accent-soft)' },
  'Finalizada':{ color:'var(--green)',   bg:'var(--green-soft)'  },
}


// ── Envío confiable hacia Apps Script ───────────────────────────────
// Apps Script devuelve JSON, pero el navegador puede bloquear la lectura
// por CORS. Usamos GET con mode:'no-cors': la solicitud sí llega y ejecuta
// MailApp, aunque el frontend no pueda leer el cuerpo de la respuesta.
async function enviarViaJsonp(url, payload) {
  if(!url) return { ok:false, error:'No hay URL de Apps Script configurada' }

  const params = new URLSearchParams({
    data: JSON.stringify(payload),
    _ts: String(Date.now()),
  })

  const requestUrl = `${url}?${params.toString()}`

  try {
    await fetch(requestUrl, {
      method: 'GET',
      mode: 'no-cors',
      cache: 'no-store',
      redirect: 'follow',
    })

    // En no-cors la respuesta es opaca; si fetch no lanzó error,
    // la solicitud fue despachada al Apps Script.
    return { ok:true, mensaje:'Solicitud de correo enviada al Apps Script' }
  } catch(fetchError) {
    console.warn('Fetch hacia Apps Script falló; usando respaldo por imagen.', fetchError)

    // Respaldo: fuerza una petición GET aunque la respuesta no sea una imagen.
    return await new Promise(resolve => {
      const img = new Image()
      let terminado = false

      const finalizar = () => {
        if(terminado) return
        terminado = true
        resolve({ ok:true, mensaje:'Solicitud enviada mediante respaldo' })
      }

      img.onload = finalizar
      img.onerror = finalizar
      img.src = requestUrl

      setTimeout(finalizar, 5000)
    })
  }
}

async function enviarCorreoProyecto({ proyecto, subtarea, uid='distribucion' }) {
  const resp = subtarea?.responsable || proyecto.responsable
  const email = obtenerCorreoResponsable(resp)
  if(!email) {
    console.error('No se encontró correo para el responsable:', resp)
    return {ok:false, error:'No se encontró correo para '+resp}
  }
  const fechaEntrega = subtarea?.fechaEntrega || proyecto.fechaEntrega || ''
  // Generar enlace Google Calendar
  const titulo = encodeURIComponent(`Entrega: ${subtarea?.nombre || proyecto.nombre}`)
  const detalle = encodeURIComponent(`Proyecto: ${proyecto.nombre} | Responsable: ${resp} | Areas: ${(subtarea?.areas||proyecto.areas||[]).join(', ')}`)
  let calLink = ''
  if(fechaEntrega) {
    const d = new Date(fechaEntrega)
    const fmt = t => t.toISOString().replace(/[-:]/g,'').split('.')[0]+'Z'
    const start = fmt(d)
    const end = fmt(new Date(d.getTime()+3600000))
    calLink = `https://calendar.google.com/calendar/r/eventedit?text=${titulo}&dates=${start}/${end}&details=${detalle}&add=${encodeURIComponent(email)}`
  }
  // Generar enlace calendario también para Paola
  let calLinkLider = ''
  if(fechaEntrega) {
    const d2 = new Date(fechaEntrega)
    const fmt2 = t => t.toISOString().replace(/[-:]/g,'').split('.')[0]+'Z'
    const start2 = fmt2(d2)
    const end2 = fmt2(new Date(d2.getTime()+3600000))
    const titulo2 = encodeURIComponent(`Entrega: ${subtarea?.nombre || proyecto.nombre}`)
    const detalle2 = encodeURIComponent(`Proyecto: ${proyecto.nombre} | Responsable: ${resp} | Areas: ${(subtarea?.areas||proyecto.areas||[]).join(', ')}`)
    calLinkLider = `https://calendar.google.com/calendar/r/eventedit?text=${titulo2}&dates=${start2}/${end2}&details=${detalle2}&add=${encodeURIComponent(EMAIL_LIDER_PROY)}`
  }
  const urlProyecto = SHEETS_CONFIG[uid]?.url || SHEETS_URL_PROY
  try {
    const respuesta = await enviarViaJsonp(urlProyecto, {
      accion: 'enviar_correo_proyecto',
      destinatario: email,
      copia: EMAIL_LIDER_PROY,
      proyecto: proyecto.nombre,
      subtarea: subtarea?.nombre || '',
      responsable: resp,
      areas: (subtarea?.areas||proyecto.areas||[]).join(', '),
      fechaEntrega,
      calLink,
      calLinkLider,
      descripcion: subtarea?.descripcion || proyecto.descripcion || '',
    })

    console.log('Correo proyecto:', respuesta)

    if(!respuesta?.ok) {
      console.error('Apps Script no confirmó el correo del proyecto:', respuesta)
    }

    return respuesta
  } catch(error) {
    console.error('Error enviando correo del proyecto:', error)
    return {ok:false, error:error?.message || 'Error enviando correo'}
  }
}


async function enviarCorreoFinalizacion({
  tipo,
  proyecto,
  subtarea=null,
  destinatarios=[],
}) {
  const correos = [...new Set(
    [EMAIL_LIDER_PROY, ...destinatarios]
      .map(v => String(v || '').trim())
      .filter(Boolean)
  )]

  if(!correos.length) {
    return {ok:false, error:'No hay destinatarios para la notificación.'}
  }

  try {
    return await enviarViaJsonp(SHEETS_URL_PROY, {
      accion: 'enviar_correo_finalizacion',
      tipo,
      destinatarios: correos.join(','),
      proyecto: proyecto?.nombre || '',
      subtarea: subtarea?.nombre || '',
      responsable: subtarea?.responsable || proyecto?.responsable || '',
      fechaFinalizacion: new Date().toLocaleString('es-CO'),
      descripcion: subtarea?.descripcion || proyecto?.descripcion || '',
    })
  } catch(error) {
    console.error('Error enviando notificación de finalización:', error)
    return {ok:false, error:error?.message || 'Error enviando correo de finalización'}
  }
}



// ═══════════════════════════════════════════════════════
// FORM PROYECTO — componente independiente (evita bug de re-render)
// ═══════════════════════════════════════════════════════
function FormProyecto({ form, setForm, onSave, titulo, onClose, guardando=false, textoGuardando='Guardando...' }) {
  return (
    <Modal title={titulo} onClose={onClose}>
      <div style={{display:'flex',flexDirection:'column',gap:14}}>
        <Field label="Nombre *"><input autoFocus value={form.nombre} onChange={e=>setForm({...form,nombre:e.target.value})} placeholder="Nombre del proyecto o tarea"/></Field>
        <Field label="Descripción"><textarea value={form.descripcion} onChange={e=>setForm({...form,descripcion:e.target.value})} rows={2} placeholder="Descripción breve" style={{resize:'vertical'}}/></Field>
        <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:12}}>
          <Field label="Responsable">
            <select value={form.responsable} onChange={e=>setForm({...form,responsable:e.target.value})}>
              <option value="">Sin asignar</option>
              {RESPONSABLES.map(r=><option key={r}>{r}</option>)}
            </select>
          </Field>
          <Field label="Fecha de entrega"><input type="date" value={form.fechaEntrega} onChange={e=>setForm({...form,fechaEntrega:e.target.value})}/></Field>
        </div>
        <Field label="Estado">
          <select value={form.estado} onChange={e=>setForm({...form,estado:e.target.value})}>
            {Object.keys(ESTADO_PROY).map(e=><option key={e}>{e}</option>)}
          </select>
        </Field>
        <Field label="Áreas que intervienen">
          <div style={{display:'flex',gap:8,flexWrap:'wrap',marginTop:4}}>
            {AREAS.map(a=>(
              <label key={a} style={{display:'flex',alignItems:'center',gap:5,cursor:'pointer',fontSize:13,padding:'5px 10px',borderRadius:7,border:'1px solid var(--border2)',background:form.areas.includes(a)?'var(--accent-soft)':'var(--bg3)',color:form.areas.includes(a)?'var(--accent2)':'var(--text2)'}}>
                <input type="checkbox" checked={form.areas.includes(a)} onChange={e=>setForm({...form,areas:e.target.checked?[...form.areas,a]:form.areas.filter(x=>x!==a)})} style={{display:'none'}}/>
                {a}
              </label>
            ))}
          </div>
        </Field>
        <button
          onClick={onSave}
          disabled={guardando}
          style={{
            ...S.btn(guardando ? 'var(--bg3)' : 'var(--accent)', guardando ? 'var(--text3)' : '#fff'),
            marginTop:4,
            cursor:guardando ? 'not-allowed' : 'pointer',
            opacity:guardando ? 0.8 : 1,
            minHeight:38,
          }}
        >
          {guardando ? `⏳ ${textoGuardando}` : 'Guardar'}
        </button>
      </div>
    </Modal>
  )
}

const OPCIONES_TIEMPO = [
  {label:'15 min', minutos:15, valor:15, unidad:'minutos'},
  {label:'30 min', minutos:30, valor:30, unidad:'minutos'},
  {label:'45 min', minutos:45, valor:45, unidad:'minutos'},
  {label:'1 hora', minutos:60, valor:1, unidad:'horas'},
  {label:'1 h 15 min', minutos:75, valor:1.25, unidad:'horas'},
  {label:'1 h 30 min', minutos:90, valor:1.5, unidad:'horas'},
  {label:'1 h 45 min', minutos:105, valor:1.75, unidad:'horas'},
  {label:'2 horas', minutos:120, valor:2, unidad:'horas'},
  {label:'2 h 15 min', minutos:135, valor:2.25, unidad:'horas'},
  {label:'2 h 30 min', minutos:150, valor:2.5, unidad:'horas'},
  {label:'2 h 45 min', minutos:165, valor:2.75, unidad:'horas'},
  {label:'3 horas', minutos:180, valor:3, unidad:'horas'},
  {label:'4 horas', minutos:240, valor:4, unidad:'horas'},
  {label:'5 horas', minutos:300, valor:5, unidad:'horas'},
  {label:'6 horas', minutos:360, valor:6, unidad:'horas'},
  {label:'7 horas', minutos:420, valor:7, unidad:'horas'},
  {label:'8 horas', minutos:480, valor:8, unidad:'horas'},
]

function nombrePersonaPorUsuario(usuario) {
  const mapa = {
    industria:'Emma',
    distribucion:'María P',
    presupuesto:'Juan',
    diseno:'Camilo',
    julian:'Julian',
    lider:'Paola',
  }
  return mapa[usuario?.id] || usuario?.nombre || ''
}

function uidPersonaPorUsuario(usuario) {
  const mapa = {
    industria:'industria',
    distribucion:'distribucion',
    presupuesto:'juan',
    diseno:'camilo',
    julian:'julian',
  }
  return mapa[usuario?.id] || usuario?.id || 'distribucion'
}

function formatoTiempoTarea(t) {
  const minutos = Number(t?.tiempoMinutos || 0)
  if(!minutos) return '—'
  if(minutos < 60) return minutos + ' min'
  if(minutos < 480) {
    const h = Math.floor(minutos/60)
    const m = Math.round(minutos%60)
    return m ? `${h} h ${m} min` : `${h} h`
  }
  const dias = minutos/480
  return Number.isInteger(dias) ? `${dias} día${dias===1?'':'s'}` : `${dias.toFixed(1)} días`
}

function ModalTiempoTarea({ tarea, onClose, onConfirm, guardando=false }) {
  const [tipo, setTipo] = useState('lista')
  const [opcion, setOpcion] = useState('60')
  const [dias, setDias] = useState('1')

  const confirmar = () => {
    if(tipo==='dias') {
      const n = Math.max(1, Number(dias)||1)
      onConfirm({tiempoValor:n,tiempoUnidad:'dias',tiempoMinutos:n*480})
      return
    }
    const opt = OPCIONES_TIEMPO.find(o=>String(o.minutos)===String(opcion)) || OPCIONES_TIEMPO[3]
    onConfirm({tiempoValor:opt.valor,tiempoUnidad:opt.unidad,tiempoMinutos:opt.minutos})
  }

  return (
    <Modal title="Finalizar tarea" onClose={()=>!guardando&&onClose()}>
      <div style={{display:'flex',flexDirection:'column',gap:16}}>
        <div style={{background:'var(--bg3)',padding:14,borderRadius:10}}>
          <div style={{fontSize:11,color:'var(--text3)',marginBottom:4}}>Tarea</div>
          <div style={{fontWeight:700,fontSize:14}}>{tarea?.nombre}</div>
        </div>
        <div style={{fontSize:13,color:'var(--text2)'}}>¿Cuánto tiempo te tomó completar esta tarea?</div>
        <div style={{display:'flex',gap:8}}>
          <button onClick={()=>setTipo('lista')} disabled={guardando}
            style={{...S.btn(tipo==='lista'?'var(--accent)':'var(--bg3)',tipo==='lista'?'#fff':'var(--text2)'),flex:1,justifyContent:'center'}}>
            Minutos / horas
          </button>
          <button onClick={()=>setTipo('dias')} disabled={guardando}
            style={{...S.btn(tipo==='dias'?'var(--accent)':'var(--bg3)',tipo==='dias'?'#fff':'var(--text2)'),flex:1,justifyContent:'center'}}>
            Días
          </button>
        </div>
        {tipo==='lista' ? (
          <Field label="Tiempo empleado">
            <select value={opcion} onChange={e=>setOpcion(e.target.value)}>
              {OPCIONES_TIEMPO.map(o=><option key={o.minutos} value={o.minutos}>{o.label}</option>)}
            </select>
          </Field>
        ) : (
          <Field label="Cantidad de días">
            <input type="number" min="1" step="1" value={dias} onChange={e=>setDias(e.target.value)}/>
          </Field>
        )}
        <button onClick={confirmar} disabled={guardando}
          style={{...S.btn(guardando?'var(--bg3)':'var(--green)',guardando?'var(--text3)':'#fff'),justifyContent:'center',minHeight:40}}>
          {guardando ? '⏳ Guardando cierre...' : '✅ Finalizar tarea'}
        </button>
      </div>
    </Modal>
  )
}



function ModalEditarTarea({ tarea, onClose, onSave, guardando=false, permiteResponsable=true, titulo='Editar tarea' }) {
  const [form,setForm] = useState({
    ...tarea,
    areas:Array.isArray(tarea?.areas)
      ? tarea.areas
      : String(tarea?.areas||'').split(',').map(x=>x.trim()).filter(Boolean)
  })

  return (
    <Modal title={titulo} onClose={()=>!guardando&&onClose()} wide>
      <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:14}}>
        <Field label="Nombre *" span>
          <input value={form.nombre||''} onChange={e=>setForm({...form,nombre:e.target.value})}/>
        </Field>
        <Field label="Descripción" span>
          <textarea rows={3} value={form.descripcion||''} onChange={e=>setForm({...form,descripcion:e.target.value})}/>
        </Field>

        {permiteResponsable&&(
          <Field label="Responsable">
            <select value={form.responsable||''} onChange={e=>setForm({...form,responsable:e.target.value})}>
              <option value="">Sin asignar</option>
              {RESPONSABLES.map(r=><option key={r}>{r}</option>)}
            </select>
          </Field>
        )}

        <Field label="Fecha límite">
          <input type="date" value={String(form.fechaLimite||'').slice(0,10)} onChange={e=>setForm({...form,fechaLimite:e.target.value})}/>
        </Field>

        <Field label="Estado">
          <select value={form.estado||'Pendiente'} onChange={e=>setForm({...form,estado:e.target.value})}>
            {['Pendiente','En proceso','Finalizada','Cancelada'].map(s=><option key={s}>{s}</option>)}
          </select>
        </Field>

        <Field label="Áreas que intervienen" span>
          <div style={{display:'flex',gap:8,flexWrap:'wrap'}}>
            {AREAS.map(a=>(
              <label key={a} style={{
                fontSize:12,padding:'6px 10px',border:'1px solid var(--border2)',
                borderRadius:7,cursor:'pointer',
                background:form.areas.includes(a)?'var(--accent-soft)':'var(--bg3)',
                color:form.areas.includes(a)?'var(--accent2)':'var(--text2)'
              }}>
                <input
                  type="checkbox"
                  style={{display:'none'}}
                  checked={form.areas.includes(a)}
                  onChange={e=>setForm({
                    ...form,
                    areas:e.target.checked?[...form.areas,a]:form.areas.filter(x=>x!==a)
                  })}
                />
                {a}
              </label>
            ))}
          </div>
        </Field>
      </div>

      <div style={{display:'flex',justifyContent:'flex-end',gap:8,marginTop:20}}>
        <button disabled={guardando} onClick={onClose} style={S.btn('var(--bg3)','var(--text2)')}>Cancelar</button>
        <button disabled={guardando} onClick={()=>onSave(form)}
          style={S.btn(guardando?'var(--bg3)':'var(--accent)',guardando?'var(--text3)':'#fff')}>
          {guardando?'⏳ Guardando cambios...':'Guardar cambios'}
        </button>
      </div>
    </Modal>
  )
}

// ═══════════════════════════════════════════════════════
// MIS TAREAS — tareas personales del usuario
// ═══════════════════════════════════════════════════════
function MisTareas({ data, setData, usuario }) {
  const tareas = data.misTareas || []
  const [modal,setModal] = useState(false)
  const [guardando,setGuardando] = useState(false)
  const [cerrando,setCerrando] = useState(null)
  const [guardandoCierre,setGuardandoCierre] = useState(false)
  const [editando,setEditando] = useState(null)
  const [guardandoEdicion,setGuardandoEdicion] = useState(false)
  const [eliminandoId,setEliminandoId] = useState(null)
  const [form,setForm] = useState({nombre:'',descripcion:'',areas:[],fechaLimite:'',estado:'Pendiente'})

  const persona = nombrePersonaPorUsuario(usuario)
  const uid = uidPersonaPorUsuario(usuario)

  const crear = async () => {
    if(guardando) return
    if(!form.nombre) return alert('Escribe el nombre de la tarea.')
    setGuardando(true)
    const nueva = {
      id:Date.now(),...form,responsable:persona,asignadoPor:persona,
      fechaCreacion:new Date().toLocaleString('es-CO'),
      fechaFinalizacion:'',tiempoValor:'',tiempoUnidad:'',tiempoMinutos:0
    }
    try {
      setData({...data,misTareas:[...tareas,nueva]})
      await insertarFilaSheets(uid,'misTareas',nueva)
      if(uid!=='distribucion') await insertarFilaSheets('distribucion','misTareas',nueva)
      setModal(false)
      setForm({nombre:'',descripcion:'',areas:[],fechaLimite:'',estado:'Pendiente'})
    } catch(e) {
      console.error(e)
      alert('No se pudo guardar la tarea.')
    } finally { setGuardando(false) }
  }

  const persistir = async actualizada => {
    const lista = tareas.map(t=>String(t.id)===String(actualizada.id)?actualizada:t)
    setData({...data,misTareas:lista})
    await eliminarFilaSheets(uid,'misTareas',actualizada.id)
    await insertarFilaSheets(uid,'misTareas',actualizada)
    if(uid!=='distribucion') {
      await eliminarFilaSheets('distribucion','misTareas',actualizada.id)
      await insertarFilaSheets('distribucion','misTareas',actualizada)
    }
  }

  const cambiarEstado = async (t,estado) => {
    if(estado==='Finalizada'&&t.estado!=='Finalizada') return setCerrando(t)
    try { await persistir({...t,estado}) } catch(e){console.error(e)}
  }

  const confirmarFinalizacion = async tiempo => {
    if(!cerrando||guardandoCierre) return
    setGuardandoCierre(true)
    const actualizada = {...cerrando,estado:'Finalizada',fechaFinalizacion:new Date().toLocaleString('es-CO'),...tiempo}
    try {
      await persistir(actualizada)
      await enviarCorreoFinalizacion({
        tipo:'tarea_finalizada',
        proyecto:{nombre:'Tarea personal'},
        subtarea:{nombre:actualizada.nombre,responsable:persona,descripcion:actualizada.descripcion},
        destinatarios:[obtenerCorreoResponsable(persona)],
      })
      setCerrando(null)
    } finally { setGuardandoCierre(false) }
  }

  const guardarEdicion = async formEdit => {
    if(!formEdit.nombre) return alert('La tarea debe tener nombre.')
    setGuardandoEdicion(true)
    try {
      await persistir({...editando,...formEdit,responsable:persona,asignadoPor:persona})
      setEditando(null)
    } catch(e) {
      console.error(e)
      alert('No se pudo guardar la edición.')
    } finally {
      setGuardandoEdicion(false)
    }
  }

  const eliminar = async id => {
    if(!confirm('¿Eliminar esta tarea personal?')) return
    setEliminandoId(id)
    try {
      setData({...data,misTareas:tareas.filter(t=>String(t.id)!==String(id))})
      await eliminarFilaSheets(uid,'misTareas',id)
      if(uid!=='distribucion') await eliminarFilaSheets('distribucion','misTareas',id)
    } catch(e) {
      console.error(e)
    } finally {
      setEliminandoId(null)
    }
  }

  return (
    <div style={{display:'flex',flexDirection:'column',gap:14}}>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center'}}>
        <div style={{fontSize:12,color:'var(--text2)'}}>Tareas personales: {tareas.length}</div>
        <button onClick={()=>setModal(true)} style={S.btn('var(--accent)','#fff')}><PlusCircle size={15}/> Nueva tarea</button>
      </div>

      {tareas.map(t=>(
        <div key={t.id} style={{...S.card,padding:'14px 17px',display:'flex',gap:12,alignItems:'flex-start'}}>
          <div style={{flex:1}}>
            <div style={{fontWeight:700,fontSize:14,marginBottom:5}}>{t.nombre}</div>
            {t.descripcion&&<div style={{fontSize:12,color:'var(--text2)',marginBottom:6}}>{t.descripcion}</div>}
            <div style={{display:'flex',gap:12,fontSize:11,color:'var(--text3)',flexWrap:'wrap'}}>
              {t.fechaLimite&&<span>📅 {String(t.fechaLimite).slice(0,10)}</span>}
              {t.areas&&<span>👥 {Array.isArray(t.areas)?t.areas.join(', '):t.areas}</span>}
              {Number(t.tiempoMinutos)>0&&<span>⏱️ {formatoTiempoTarea(t)}</span>}
            </div>
          </div>
          <select value={t.estado||'Pendiente'} onChange={e=>cambiarEstado(t,e.target.value)} style={{fontSize:11}}>
            {['Pendiente','En proceso','Finalizada','Cancelada'].map(s=><option key={s}>{s}</option>)}
          </select>
          <button onClick={()=>setEditando(t)} style={{...S.btn('var(--accent-soft)','var(--accent2)'),padding:'5px 7px'}}><Edit2 size={12}/></button>
          <button disabled={eliminandoId===t.id} onClick={()=>eliminar(t.id)} style={{...S.btn('var(--red-soft)','var(--red)'),padding:'5px 7px'}}>
            {eliminandoId===t.id?'…':<Trash2 size={12}/>}
          </button>
        </div>
      ))}

      {modal&&(
        <Modal title="Nueva tarea personal" onClose={()=>!guardando&&setModal(false)}>
          <div style={{display:'flex',flexDirection:'column',gap:14}}>
            <Field label="Nombre *"><input value={form.nombre} onChange={e=>setForm({...form,nombre:e.target.value})}/></Field>
            <Field label="Descripción"><textarea rows={2} value={form.descripcion} onChange={e=>setForm({...form,descripcion:e.target.value})}/></Field>
            <Field label="Fecha límite"><input type="date" value={form.fechaLimite} onChange={e=>setForm({...form,fechaLimite:e.target.value})}/></Field>
            <Field label="Áreas que intervienen">
              <div style={{display:'flex',gap:7,flexWrap:'wrap'}}>
                {AREAS.map(a=>(
                  <label key={a} style={{fontSize:12,padding:'5px 9px',border:'1px solid var(--border2)',borderRadius:6,cursor:'pointer',background:form.areas.includes(a)?'var(--accent-soft)':'var(--bg3)'}}>
                    <input type="checkbox" style={{display:'none'}} checked={form.areas.includes(a)}
                      onChange={e=>setForm({...form,areas:e.target.checked?[...form.areas,a]:form.areas.filter(x=>x!==a)})}/>
                    {a}
                  </label>
                ))}
              </div>
            </Field>
            <button onClick={crear} disabled={guardando}
              style={{...S.btn(guardando?'var(--bg3)':'var(--accent)',guardando?'var(--text3)':'#fff'),justifyContent:'center'}}>
              {guardando?'⏳ Guardando tarea...':'Crear tarea'}
            </button>
          </div>
        </Modal>
      )}

      {cerrando&&<ModalTiempoTarea tarea={cerrando} onClose={()=>setCerrando(null)} onConfirm={confirmarFinalizacion} guardando={guardandoCierre}/>}
      {editando&&<ModalEditarTarea tarea={editando} onClose={()=>setEditando(null)} onSave={guardarEdicion} guardando={guardandoEdicion} permiteResponsable={false} titulo="Editar mi tarea"/>}
    </div>
  )
}

function TareasYProyectos({ data, setData, usuario }) {
  const [seccion, setSeccion] = useState('proyectos') // proyectos | tareas
  const proyectos = data.proyectos || []
  const [vista, setVista] = useState('lista') // lista | detalle
  const [proyectoActivo, setProyectoActivo] = useState(null)
  const [modalProyecto, setModalProyecto] = useState(false)
  const [modalSubtarea, setModalSubtarea] = useState(false)
  const [guardandoProyecto, setGuardandoProyecto] = useState(false)
  const [guardandoSubtarea, setGuardandoSubtarea] = useState(false)
  const [editProyecto, setEditProyecto] = useState(null)
  const [editSubtarea, setEditSubtarea] = useState(null)
  const [guardandoEdicionProyecto, setGuardandoEdicionProyecto] = useState(false)
  const [guardandoEdicionSubtarea, setGuardandoEdicionSubtarea] = useState(false)
  const [eliminandoProyId, setEliminandoProyId] = useState(null)
  const [eliminandoSubId, setEliminandoSubId] = useState(null)
  const [proyPadre, setProyPadre] = useState(null)
  const [filtroEstado, setFiltroEstado] = useState('')
  const [formP, setFormP] = useState({ nombre:'', descripcion:'', responsable:'', areas:[], fechaEntrega:'', estado:'Pendiente' })
  const [formS, setFormS] = useState({ nombre:'', descripcion:'', responsable:'', areas:[], fechaEntrega:'', estado:'Pendiente' })

  const sheetUid = SHEETS_CONFIG[usuario.id]?.enabled ? usuario.id : 'distribucion'

  const crearProyecto = async () => {
    if(guardandoProyecto) return
    if(!formP.nombre) {
      alert('Escribe el nombre del proyecto.')
      return
    }

    setGuardandoProyecto(true)

    const nuevo = { id:Date.now(), ...formP, creadoPor:usuario.nombre, subtareas:[], fechaCreacion:new Date().toLocaleDateString('es-CO'), correoFinalEnviado:'NO' }
    const nd = {...data, proyectos:[...proyectos, nuevo]}
    setData(nd)
    // Todos los proyectos se guardan en el archivo central de Distribución.
    // Así Paola siempre ve el historial completo, sin importar el responsable.
    const uidCentral = 'distribucion'
    const uidResponsable = obtenerUidResponsable(nuevo.responsable)

    try {
      // 1. Siempre conservar una copia maestra en el archivo de María P / Distribución.
      await insertarFilaSheets(uidCentral, 'proyectos', nuevo)

      // 2. Guardar también el proyecto en el archivo personal del responsable.
      // María P usa el mismo archivo central, por eso evitamos duplicarlo.
      if(uidResponsable && uidResponsable !== uidCentral) {
        await insertarFilaSheets(uidResponsable, 'proyectos', nuevo)
      }
    } catch(e) {
      console.error('Error guardando proyecto central o en el archivo del responsable:', e)
      alert('El proyecto se creó en pantalla, pero hubo un error al copiarlo a Google Sheets.')
    }

    // El correo se dirige al responsable seleccionado.
    const resultadoCorreo = await enviarCorreoProyecto({
      proyecto:nuevo,
      subtarea:null,
      uid:uidResponsable,
    })
    console.log('Resultado correo nuevo proyecto:', resultadoCorreo)
    setModalProyecto(false)
    setFormP({ nombre:'', descripcion:'', responsable:'', areas:[], fechaEntrega:'', estado:'Pendiente' })
    setGuardandoProyecto(false)
  }

  const crearSubtarea = async () => {
    if(guardandoSubtarea) return
    if(!formS.nombre || !proyPadre) {
      alert('Escribe el nombre de la subtarea.')
      return
    }

    if(!formS.responsable) {
      alert('Selecciona un responsable para la subtarea.')
      return
    }

    setGuardandoSubtarea(true)

    const ahora = Date.now()
    const subtarea = {
      id: ahora,
      ...formS,
      creadoPor: usuario.nombre,
      fechaCreacion: new Date().toLocaleDateString('es-CO'),
    }

    // 1. Mantener la subtarea dentro del proyecto padre.
    const proyectosActualizados = proyectos.map(p =>
      p.id === proyPadre.id
        ? {...p, subtareas:[...(p.subtareas || []), subtarea]}
        : p
    )
    const proyectoActualizado = proyectosActualizados.find(p => p.id === proyPadre.id)

    // 2. La subtarea se guarda en una hoja separada.
    const filaSubtarea = {
      id: subtarea.id,
      proyecto: proyPadre.nombre,
      nombre: subtarea.nombre,
      descripcion: subtarea.descripcion || '',
      responsable: subtarea.responsable,
      fechaEntrega: subtarea.fechaEntrega || '',
      estado: subtarea.estado || 'Pendiente',
      areas: (subtarea.areas || []).join(', '),
      creadoPor: subtarea.creadoPor,
      fechaCreacion: subtarea.fechaCreacion,
    }

    const nd = {
      ...data,
      proyectos: proyectosActualizados,
    }
    setData(nd)

    // El proyecto y todas sus subtareas se conservan en el archivo central.
    const uidCentral = 'distribucion'
    const uidResponsable = obtenerUidResponsable(subtarea.responsable)

    try {
      await eliminarFilaSheets(uidCentral, 'proyectos', proyPadre.id)
      if(proyectoActualizado) {
        await insertarFilaSheets(uidCentral, 'proyectos', proyectoActualizado)
      }

      // Registro central para que Paola siempre vea todas las subtareas.
      await insertarFilaSheets(uidCentral, 'subtareasProyectos', filaSubtarea)

      // Copia adicional en el archivo personal/área del responsable.
      if(uidResponsable !== uidCentral) {
        await insertarFilaSheets(uidResponsable, 'subtareasProyectos', filaSubtarea)
      }

      console.log(
        'Subtarea guardada centralmente y asignada a',
        subtarea.responsable
      )
    } catch(e) {
      console.error('Error guardando la subtarea:', e)
      alert('La subtarea se creó en pantalla, pero hubo un error guardándola en Sheets.')
    }

    const resultadoCorreo = await enviarCorreoProyecto({
      proyecto: proyPadre,
      subtarea,
      uid: uidResponsable,
    })

    console.log('Resultado correo nueva subtarea:', resultadoCorreo)

    setModalSubtarea(false)
    setFormS({
      nombre:'',
      descripcion:'',
      responsable:'',
      areas:[],
      fechaEntrega:'',
      estado:'Pendiente'
    })
    setGuardandoSubtarea(false)
  }

  const guardarEdicionProyecto = async formEdit => {
    if(!formEdit.nombre) return alert('El proyecto debe tener nombre.')
    setGuardandoEdicionProyecto(true)
    const anterior = proyectos.find(p=>String(p.id)===String(editProyecto.id))
    const actualizado = {...anterior,...formEdit}
    const uidAnterior = obtenerUidResponsable(anterior?.responsable)
    const uidNuevo = obtenerUidResponsable(actualizado.responsable)

    try {
      const lista = proyectos.map(p=>String(p.id)===String(actualizado.id)?actualizado:p)
      setData({...data,proyectos:lista})
      if(proyectoActivo?.id===actualizado.id) setProyectoActivo(actualizado)

      await eliminarFilaSheets('distribucion','proyectos',actualizado.id)
      await insertarFilaSheets('distribucion','proyectos',actualizado)

      if(uidAnterior && uidAnterior!=='distribucion' && uidAnterior!==uidNuevo) {
        await eliminarFilaSheets(uidAnterior,'proyectos',actualizado.id)
      }
      if(uidNuevo && uidNuevo!=='distribucion') {
        await eliminarFilaSheets(uidNuevo,'proyectos',actualizado.id)
        await insertarFilaSheets(uidNuevo,'proyectos',actualizado)
      }

      if(normalizarNombreResponsable(anterior?.responsable)!==normalizarNombreResponsable(actualizado.responsable)) {
        await enviarCorreoProyecto({proyecto:actualizado,subtarea:null,uid:uidNuevo})
      }
      setEditProyecto(null)
    } catch(e) {
      console.error(e)
      alert('No se pudo editar el proyecto.')
    } finally {
      setGuardandoEdicionProyecto(false)
    }
  }

  const guardarEdicionSubtarea = async formEdit => {
    if(!formEdit.nombre) return alert('La subtarea debe tener nombre.')
    setGuardandoEdicionSubtarea(true)
    const proyecto = proyectos.find(p=>String(p.id)===String(editSubtarea.pid))
    const anterior = proyecto?.subtareas?.find(s=>String(s.id)===String(editSubtarea.subtarea.id))
    if(!proyecto||!anterior) {
      setGuardandoEdicionSubtarea(false)
      return
    }

    const actualizada = {...anterior,...formEdit}
    const proyectoActualizado = {
      ...proyecto,
      subtareas:(proyecto.subtareas||[]).map(s=>String(s.id)===String(actualizada.id)?actualizada:s)
    }
    const uidAnterior = obtenerUidResponsable(anterior.responsable)
    const uidNuevo = obtenerUidResponsable(actualizada.responsable)
    const fila = {
      id:actualizada.id,proyecto:proyecto.nombre,nombre:actualizada.nombre,
      descripcion:actualizada.descripcion||'',responsable:actualizada.responsable||'',
      fechaEntrega:actualizada.fechaEntrega||'',estado:actualizada.estado||'Pendiente',
      areas:Array.isArray(actualizada.areas)?actualizada.areas.join(', '):(actualizada.areas||''),
      creadoPor:actualizada.creadoPor||'',fechaCreacion:actualizada.fechaCreacion||''
    }

    try {
      const lista = proyectos.map(p=>String(p.id)===String(proyecto.id)?proyectoActualizado:p)
      setData({...data,proyectos:lista})
      if(proyectoActivo?.id===proyecto.id) setProyectoActivo(proyectoActualizado)

      await eliminarFilaSheets('distribucion','proyectos',proyecto.id)
      await insertarFilaSheets('distribucion','proyectos',proyectoActualizado)
      await eliminarFilaSheets('distribucion','subtareasProyectos',actualizada.id)
      await insertarFilaSheets('distribucion','subtareasProyectos',fila)

      if(uidAnterior && uidAnterior!=='distribucion' && uidAnterior!==uidNuevo) {
        await eliminarFilaSheets(uidAnterior,'subtareasProyectos',actualizada.id)
      }
      if(uidNuevo && uidNuevo!=='distribucion') {
        await eliminarFilaSheets(uidNuevo,'subtareasProyectos',actualizada.id)
        await insertarFilaSheets(uidNuevo,'subtareasProyectos',fila)
      }

      if(normalizarNombreResponsable(anterior.responsable)!==normalizarNombreResponsable(actualizada.responsable)) {
        await enviarCorreoProyecto({proyecto,subtarea:actualizada,uid:uidNuevo})
      }
      setEditSubtarea(null)
    } catch(e) {
      console.error(e)
      alert('No se pudo editar la subtarea.')
    } finally {
      setGuardandoEdicionSubtarea(false)
    }
  }

  const cambiarEstadoProy = (pid, estado) => {
    const proyectosActualizados = proyectos.map(p=> p.id===pid ? {...p,estado} : p)
    const nd = {...data, proyectos: proyectosActualizados}
    const proyectoActualizado = proyectosActualizados.find(p=>p.id===pid)
    setData(nd)
    const destinoUid = obtenerUidResponsable(proyectoActualizado?.responsable)
    eliminarFilaSheets(destinoUid,'proyectos',pid)
      .then(()=>proyectoActualizado&&insertarFilaSheets(destinoUid,'proyectos',proyectoActualizado))
      .catch(e=>console.error('Error actualizando proyecto:',e))
    if(proyectoActivo?.id===pid) setProyectoActivo({...proyectoActivo,estado})
  }

  const cambiarEstadoSub = async (pid, sid, estado) => {
    const proyectoOriginal = proyectos.find(p => p.id === pid)
    const subtareaOriginal = proyectoOriginal?.subtareas?.find(s => String(s.id) === String(sid))

    if(!proyectoOriginal || !subtareaOriginal) return

    const tareaSeAcabaDeFinalizar =
      estado === 'Finalizada' && subtareaOriginal.estado !== 'Finalizada'

    let proyectoActualizado = {
      ...proyectoOriginal,
      subtareas:(proyectoOriginal.subtareas || []).map(s =>
        String(s.id) === String(sid) ? {...s, estado} : s
      ),
    }

    const subtareasProyecto = proyectoActualizado.subtareas || []
    const proyectoLlegoAl100 =
      subtareasProyecto.length > 0 &&
      subtareasProyecto.every(s => s.estado === 'Finalizada')

    const debeNotificarProyecto =
      proyectoLlegoAl100 &&
      String(proyectoOriginal.correoFinalEnviado || 'NO').toUpperCase() !== 'SI'

    if(proyectoLlegoAl100) {
      proyectoActualizado = {
        ...proyectoActualizado,
        estado:'Finalizada',
        correoFinalEnviado: debeNotificarProyecto ? 'SI' : (proyectoOriginal.correoFinalEnviado || 'SI'),
      }
    }

    const proyectosActualizados = proyectos.map(p =>
      p.id === pid ? proyectoActualizado : p
    )

    const nd = {...data, proyectos: proyectosActualizados}
    const subtareaActualizada = proyectoActualizado.subtareas.find(
      s => String(s.id) === String(sid)
    )

    setData(nd)
    if(proyectoActivo?.id === pid) setProyectoActivo(proyectoActualizado)

    const uidCentral = 'distribucion'
    const uidResponsable = obtenerUidResponsable(subtareaActualizada?.responsable)

    try {
      // Guardar primero; después notificar.
      await eliminarFilaSheets(uidCentral, 'proyectos', pid)
      await insertarFilaSheets(uidCentral, 'proyectos', proyectoActualizado)

      if(subtareaActualizada) {
        const filaSubtarea = {
          id: subtareaActualizada.id,
          proyecto: proyectoOriginal.nombre || '',
          nombre: subtareaActualizada.nombre || '',
          descripcion: subtareaActualizada.descripcion || '',
          responsable: subtareaActualizada.responsable || '',
          fechaEntrega: subtareaActualizada.fechaEntrega || '',
          estado: subtareaActualizada.estado || 'Pendiente',
          areas: (subtareaActualizada.areas || []).join(', '),
          creadoPor: subtareaActualizada.creadoPor || '',
          fechaCreacion: subtareaActualizada.fechaCreacion || '',
        }

        await eliminarFilaSheets(uidCentral, 'subtareasProyectos', sid)
        await insertarFilaSheets(uidCentral, 'subtareasProyectos', filaSubtarea)

        if(uidResponsable !== uidCentral) {
          await eliminarFilaSheets(uidResponsable, 'subtareasProyectos', sid)
          await insertarFilaSheets(uidResponsable, 'subtareasProyectos', filaSubtarea)
        }
      }

      // 1. Correo cuando una tarea pasa por primera vez a Finalizada.
      if(tareaSeAcabaDeFinalizar) {
        await enviarCorreoFinalizacion({
          tipo:'tarea_finalizada',
          proyecto:proyectoActualizado,
          subtarea:subtareaActualizada,
          destinatarios:[
            obtenerCorreoResponsable(subtareaActualizada?.responsable),
          ],
        })
      }

      // 2. Correo general cuando el proyecto llega al 100 %.
      if(debeNotificarProyecto) {
        const correosInvolucrados = [
          obtenerCorreoResponsable(proyectoActualizado.responsable),
          ...subtareasProyecto.map(s => obtenerCorreoResponsable(s.responsable)),
        ]

        await enviarCorreoFinalizacion({
          tipo:'proyecto_finalizado',
          proyecto:proyectoActualizado,
          destinatarios:correosInvolucrados,
        })
      }
    } catch(e) {
      console.error('Error actualizando o notificando la subtarea:', e)
      alert('El estado cambió en pantalla, pero hubo un error al guardar o enviar la notificación.')
    }
  }

  const eliminarProy = (pid) => {
    if(!window.confirm('¿Eliminar este proyecto y todas sus subtareas?')) return
    const nd = {...data, proyectos: proyectos.filter(p=>p.id!==pid)}
    setData(nd)
    eliminarFilaSheets('distribucion','proyectos',pid)
      .catch(e=>console.error('Error eliminando proyecto:',e))
    setVista('lista'); setProyectoActivo(null)
  }

  const eliminarSub = (pid, sid) => {
    const proyectosActualizados = proyectos.map(p=> p.id===pid ? {...p, subtareas:(p.subtareas||[]).filter(s=>s.id!==sid)} : p)
    const nd = {...data, proyectos: proyectosActualizados}
    const proyectoActualizado = proyectosActualizados.find(p=>p.id===pid)
    setData(nd)
    const destinoUid = obtenerUidResponsable(proyectoActualizado?.responsable)
    eliminarFilaSheets(destinoUid,'proyectos',pid)
      .then(()=>proyectoActualizado&&insertarFilaSheets(destinoUid,'proyectos',proyectoActualizado))
      .catch(e=>console.error('Error eliminando subtarea:',e))
    if(proyectoActivo?.id===pid) setProyectoActivo(nd.proyectos.find(p=>p.id===pid))
  }

  const proyFiltrados = filtroEstado ? proyectos.filter(p=>p.estado===filtroEstado) : proyectos

  // FormProyecto movido afuera como componente global

  // ── DETALLE DE PROYECTO ──
  if(vista==='detalle'&&proyectoActivo) {
    const proy = proyectos.find(p=>p.id===proyectoActivo.id)||proyectoActivo
    const subtareas = proy.subtareas||[]
    const total = subtareas.length
    const finalizadas = subtareas.filter(s=>s.estado==='Finalizada').length
    const pct = total>0?Math.round(finalizadas/total*100):0
    const est = ESTADO_PROY[proy.estado]||ESTADO_PROY['Pendiente']
    return (
      <div style={{display:'flex',flexDirection:'column',gap:16}}>
        <div style={{display:'flex',alignItems:'center',gap:10}}>
          <button onClick={()=>setVista('lista')} style={{...S.btn('var(--bg3)','var(--text2)'),padding:'6px 12px',fontSize:12}}>← Volver</button>
          <span style={{fontWeight:700,fontSize:17,color:'var(--text)',flex:1}}>{proy.nombre}</span>
          <button onClick={()=>{setProyPadre(proy);setModalSubtarea(true)}} style={{...S.btn('var(--accent)','#fff'),fontSize:12}}>+ Subtarea</button>
          <button onClick={()=>setEditProyecto(proy)} style={{...S.btn('var(--accent-soft)','var(--accent2)'),padding:'6px 10px'}}><Edit2 size={14}/> Editar</button>
          <button onClick={()=>eliminarProy(proy.id)} style={{...S.btn('var(--red-soft)','var(--red)'),padding:'6px 10px'}}><Trash2 size={14}/></button>
        </div>
        <div style={{...S.card,padding:'18px 20px',display:'flex',flexDirection:'column',gap:12}}>
          <div style={{display:'flex',gap:10,flexWrap:'wrap',alignItems:'center'}}>
            <select value={proy.estado} onChange={e=>cambiarEstadoProy(proy.id,e.target.value)} style={{fontSize:12,padding:'4px 10px',borderRadius:7,border:'1px solid '+est.color,color:est.color,background:est.bg,fontWeight:600,cursor:'pointer'}}>
              {Object.keys(ESTADO_PROY).map(e=><option key={e}>{e}</option>)}
            </select>
            {proy.responsable&&<span style={{fontSize:12,background:'var(--bg3)',padding:'3px 10px',borderRadius:6,color:'var(--text2)'}}>👤 {proy.responsable}</span>}
            {proy.fechaEntrega&&<span style={{fontSize:12,color:'var(--text3)'}}>📅 Entrega: {proy.fechaEntrega}</span>}
            {(proy.areas||[]).map(a=><span key={a} style={{fontSize:11,background:'var(--accent-soft)',color:'var(--accent2)',padding:'2px 8px',borderRadius:5}}>{a}</span>)}
          </div>
          {proy.descripcion&&<p style={{fontSize:13,color:'var(--text2)',margin:0}}>{proy.descripcion}</p>}
          {total>0&&(
            <div>
              <div style={{display:'flex',justifyContent:'space-between',fontSize:11,color:'var(--text3)',marginBottom:4}}>
                <span>Progreso</span><span>{finalizadas}/{total} subtareas · {pct}%</span>
              </div>
              <div style={{height:6,background:'var(--bg4)',borderRadius:3}}>
                <div style={{width:pct+'%',height:'100%',background:pct===100?'var(--green)':'var(--accent)',borderRadius:3,transition:'width 0.3s'}}/>
              </div>
            </div>
          )}
        </div>
        {subtareas.length===0&&(
          <div style={{...S.card,padding:36,textAlign:'center',color:'var(--text3)'}}>
            <div style={{fontSize:28,marginBottom:8}}>📋</div>
            <div>Sin subtareas aún — usa <strong>+ Subtarea</strong></div>
          </div>
        )}
        {subtareas.map((s,i)=>{
          const se = ESTADO_PROY[s.estado]||ESTADO_PROY['Pendiente']
          return (
            <div key={i} style={{...S.card,padding:'14px 18px',borderLeft:'3px solid '+se.color,display:'flex',gap:12,alignItems:'flex-start'}}>
              <div style={{flex:1,minWidth:0}}>
                <div style={{display:'flex',gap:8,alignItems:'center',flexWrap:'wrap',marginBottom:4}}>
                  <span style={{fontWeight:600,fontSize:14}}>{s.nombre}</span>
                  {s.responsable&&<span style={{fontSize:11,background:'var(--bg3)',padding:'2px 8px',borderRadius:5,color:'var(--text2)'}}>👤 {s.responsable}</span>}
                  {s.fechaEntrega&&<span style={{fontSize:11,color:'var(--text3)'}}>📅 {s.fechaEntrega}</span>}
                </div>
                {s.descripcion&&<p style={{margin:'0 0 6px',fontSize:12,color:'var(--text2)'}}>{s.descripcion}</p>}
                <div style={{display:'flex',gap:6,flexWrap:'wrap'}}>
                  {(s.areas||[]).map(a=><span key={a} style={{fontSize:10,background:'var(--accent-soft)',color:'var(--accent2)',padding:'1px 7px',borderRadius:4}}>{a}</span>)}
                </div>
              </div>
              <div style={{display:'flex',flexDirection:'column',gap:6,alignItems:'flex-end',flexShrink:0}}>
                <select value={s.estado} onChange={e=>cambiarEstadoSub(proy.id,s.id,e.target.value)} style={{fontSize:11,padding:'3px 8px',borderRadius:6,border:'1px solid '+se.color,color:se.color,background:se.bg,fontWeight:600,cursor:'pointer'}}>
                  {Object.keys(ESTADO_PROY).map(e=><option key={e}>{e}</option>)}
                </select>
                <div style={{display:'flex',gap:5}}>
                  <button onClick={()=>setEditSubtarea({pid:proy.id,subtarea:s})} style={{...S.btn('var(--accent-soft)','var(--accent2)'),padding:'3px 7px'}}><Edit2 size={11}/></button>
                  <button onClick={()=>eliminarSub(proy.id,s.id)} style={{...S.btn('var(--red-soft)','var(--red)'),padding:'3px 7px'}}><Trash2 size={11}/></button>
                </div>
              </div>
            </div>
          )
        })}
        {editProyecto&&<FormProyecto
          form={editProyecto}
          setForm={setEditProyecto}
          onSave={()=>guardarEdicionProyecto(editProyecto)}
          titulo="Editar proyecto"
          onClose={()=>{if(!guardandoEdicionProyecto)setEditProyecto(null)}}
          guardando={guardandoEdicionProyecto}
          textoGuardando="Guardando cambios..."
        />}
        {editSubtarea&&<FormProyecto
          form={editSubtarea.subtarea}
          setForm={f=>setEditSubtarea(prev=>({...prev,subtarea:typeof f==='function'?f(prev.subtarea):f}))}
          onSave={()=>guardarEdicionSubtarea(editSubtarea.subtarea)}
          titulo="Editar subtarea"
          onClose={()=>{if(!guardandoEdicionSubtarea)setEditSubtarea(null)}}
          guardando={guardandoEdicionSubtarea}
          textoGuardando="Guardando cambios..."
        />}
        {modalSubtarea&&<FormProyecto
          form={formS}
          setForm={setFormS}
          onSave={crearSubtarea}
          titulo="Nueva subtarea"
          onClose={()=>{ if(!guardandoSubtarea) setModalSubtarea(false) }}
          guardando={guardandoSubtarea}
          textoGuardando="Guardando subtarea..."
        />}
      </div>
    )
  }

  // ── LISTA ──
  return (
    <div style={{display:'flex',flexDirection:'column',gap:16}}>
      {/* Tabs de sección */}
      <div style={{display:'flex',gap:6,borderBottom:'2px solid var(--border)',paddingBottom:0}}>
        {[['proyectos','🗂️ Proyectos'],['tareas','✅ Mis Tareas']].map(([id,label])=>(
          <button key={id} onClick={()=>setSeccion(id)} style={{padding:'8px 18px',fontSize:13,fontWeight:seccion===id?700:400,border:'none',background:'none',cursor:'pointer',color:seccion===id?'var(--accent2)':'var(--text2)',borderBottom:seccion===id?'2px solid var(--accent2)':'2px solid transparent',marginBottom:-2,fontFamily:'var(--font)'}}>
            {label}
          </button>
        ))}
      </div>

      {/* TAREAS PERSONALES */}
      {seccion==='tareas'&&<MisTareas data={data} setData={setData} usuario={usuario}/>}

      {/* PROYECTOS */}
      {seccion==='proyectos'&&<>
      <div style={{display:'flex',gap:10,alignItems:'center',flexWrap:'wrap'}}>
        <div style={{display:'flex',gap:6}}>
          {['','Pendiente','En proceso','Finalizada'].map(e=>(
            <button key={e} onClick={()=>setFiltroEstado(e)} style={{...S.btn(filtroEstado===e?'var(--accent)':'var(--bg3)',filtroEstado===e?'#fff':'var(--text2)'),fontSize:12,padding:'5px 12px'}}>{e||'Todos'}</button>
          ))}
        </div>
        <button onClick={()=>setModalProyecto(true)} style={{...S.btn('var(--accent)','#fff'),marginLeft:'auto'}}><PlusCircle size={15}/> Nuevo proyecto</button>
      </div>
      <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fill,minmax(300px,1fr))',gap:14}}>
        {proyFiltrados.length===0&&(
          <div style={{...S.card,padding:48,textAlign:'center',color:'var(--text3)',gridColumn:'1/-1'}}>
            <div style={{fontSize:32,marginBottom:10}}>🗂️</div>
            <div style={{fontWeight:600,fontSize:15,marginBottom:6}}>Sin proyectos aún</div>
            <div style={{fontSize:13}}>Crea tu primer proyecto con <strong>Nuevo proyecto</strong></div>
          </div>
        )}
        {proyFiltrados.map((p,i)=>{
          const est = ESTADO_PROY[p.estado]||ESTADO_PROY['Pendiente']
          const subs = p.subtareas||[]
          const fin = subs.filter(s=>s.estado==='Finalizada').length
          const pct = subs.length>0?Math.round(fin/subs.length*100):0
          return (
            <div key={i} style={{...S.card,padding:0,overflow:'hidden',cursor:'pointer',borderTop:'3px solid '+est.color}} onClick={()=>{setProyectoActivo(p);setVista('detalle')}}>
              <div style={{padding:'16px 18px'}}>
                <div style={{display:'flex',justifyContent:'space-between',alignItems:'flex-start',marginBottom:8}}>
                  <span style={{fontWeight:700,fontSize:14,color:'var(--text)',flex:1,marginRight:8}}>{p.nombre}</span>
                  <button onClick={e=>{e.stopPropagation();setEditProyecto(p)}} style={{...S.btn('var(--accent-soft)','var(--accent2)'),padding:'3px 6px',marginRight:5}}><Edit2 size={11}/></button>
                  <span style={{fontSize:10,fontWeight:700,padding:'2px 8px',borderRadius:5,background:est.bg,color:est.color,flexShrink:0,whiteSpace:'nowrap'}}>{p.estado}</span>
                </div>
                {p.descripcion&&<p style={{margin:'0 0 8px',fontSize:12,color:'var(--text2)',display:'-webkit-box',WebkitLineClamp:2,WebkitBoxOrient:'vertical',overflow:'hidden'}}>{p.descripcion}</p>}
                <div style={{display:'flex',gap:8,flexWrap:'wrap',marginBottom:10}}>
                  {p.responsable&&<span style={{fontSize:11,color:'var(--text2)'}}>👤 {p.responsable}</span>}
                  {p.fechaEntrega&&<span style={{fontSize:11,color:'var(--text3)'}}>📅 {p.fechaEntrega}</span>}
                </div>
                {(p.areas||[]).length>0&&(
                  <div style={{display:'flex',gap:5,flexWrap:'wrap',marginBottom:10}}>
                    {p.areas.map(a=><span key={a} style={{fontSize:10,background:'var(--accent-soft)',color:'var(--accent2)',padding:'1px 7px',borderRadius:4}}>{a}</span>)}
                  </div>
                )}
                {subs.length>0&&(
                  <div>
                    <div style={{display:'flex',justifyContent:'space-between',fontSize:10,color:'var(--text3)',marginBottom:3}}>
                      <span>{fin}/{subs.length} subtareas</span><span>{pct}%</span>
                    </div>
                    <div style={{height:4,background:'var(--bg4)',borderRadius:2}}>
                      <div style={{width:pct+'%',height:'100%',background:pct===100?'var(--green)':'var(--accent)',borderRadius:2}}/>
                    </div>
                  </div>
                )}
                {subs.length===0&&<span style={{fontSize:11,color:'var(--text3)'}}>Sin subtareas</span>}
              </div>
              <div style={{padding:'8px 18px',borderTop:'1px solid var(--border)',fontSize:11,color:'var(--text3)',display:'flex',justifyContent:'space-between'}}>
                <span>Por {p.creadoPor}</span><span>{p.fechaCreacion}</span>
              </div>
            </div>
          )
        })}
      </div>
      {editProyecto&&<FormProyecto
        form={editProyecto}
        setForm={setEditProyecto}
        onSave={()=>guardarEdicionProyecto(editProyecto)}
        titulo="Editar proyecto"
        onClose={()=>{if(!guardandoEdicionProyecto)setEditProyecto(null)}}
        guardando={guardandoEdicionProyecto}
        textoGuardando="Guardando cambios..."
      />}
      {modalProyecto&&<FormProyecto
        form={formP}
        setForm={setFormP}
        onSave={crearProyecto}
        titulo="Nuevo proyecto"
        onClose={()=>{ if(!guardandoProyecto) setModalProyecto(false) }}
        guardando={guardandoProyecto}
        textoGuardando="Guardando proyecto..."
      />}
      </>}
    </div>
  )
}

// ═══════════════════════════════════════════════════════
// SISTEMA DE USUARIOS
// ═══════════════════════════════════════════════════════
const USUARIOS = [
  { id:'industria',    nombre:'Industria',          pass:'ind2026',   color:'#06b6d4', rol:'normal' },
  { id:'distribucion', nombre:'Distribución',        pass:'0000',      color:'#3dd68c', rol:'normal' },
  { id:'zonas',        nombre:'Zonas Directas',      pass:'zon2026',   color:'#ff9f43', rol:'normal' },
  { id:'presupuesto',  nombre:'Presupuesto (Juan)',  pass:'pres2026',  color:'#a78bfa', rol:'presupuesto' },
  { id:'lider',        nombre:'Líder de Mercadeo',   pass:'lider2026', color:'#f59e0b', rol:'lider' },
  { id:'diseno',       nombre:'Diseño',              pass:'dis2026x',  color:'#ec4899', rol:'normal' },
  { id:'julian',       nombre:'Julian',              pass:'jul2026',   color:'#14b8a6', rol:'colaborador' },
]
const AUTH_KEY = 'prolub_auth'
const getStorageKey = uid => 'tracker_v3_'+uid

function getSessionUser() {
  try { const s=localStorage.getItem(AUTH_KEY); return s?JSON.parse(s):null } catch { return null }
}

function hexToRgb(hex) {
  const r=parseInt(hex.slice(1,3),16), g=parseInt(hex.slice(3,5),16), b=parseInt(hex.slice(5,7),16)
  return r+','+g+','+b
}

function LoginScreen({ onLogin }) {
  const [selUser, setSelUser] = useState(null)
  const [pass, setPass] = useState('')
  const [error, setError] = useState('')

  const intentar = () => {
    if(!selUser) return
    const u = USUARIOS.find(u=>u.id===selUser)
    if(u.pass===pass) { localStorage.setItem(AUTH_KEY,JSON.stringify({id:u.id,nombre:u.nombre,rol:u.rol,color:u.color})); onLogin({id:u.id,nombre:u.nombre,rol:u.rol,color:u.color}) }
    else { setError('Contraseña incorrecta'); setTimeout(()=>setError(''),2000) }
  }

  return (
    <div style={{minHeight:'100vh',display:'flex',alignItems:'center',justifyContent:'center',background:'linear-gradient(135deg,#1e1040 0%,#2d1a6e 50%,#1e1040 100%)',padding:20}}>
      <div style={{width:'100%',maxWidth:420}}>
        <div style={{textAlign:'center',marginBottom:32}}>
          <div style={{width:52,height:52,borderRadius:14,background:'var(--accent)',display:'flex',alignItems:'center',justifyContent:'center',margin:'0 auto 16px'}}>
            <BarChart2 size={26} color="#fff"/>
          </div>
          <h1 style={{fontSize:22,fontWeight:700,letterSpacing:'-0.02em',marginBottom:6}}>Prolub Trade Marketing</h1>
          <p style={{fontSize:13,color:'var(--text3)'}}>Selecciona tu unidad y accede</p>
        </div>
        <div style={{background:'rgba(255,255,255,0.97)',border:'1px solid rgba(255,255,255,0.2)',borderRadius:20,padding:32,display:'flex',flexDirection:'column',gap:16,boxShadow:'0 24px 64px rgba(0,0,0,0.4)'}}>
          <div>
            <label style={{fontSize:11,color:'var(--text2)',fontWeight:500,textTransform:'uppercase',letterSpacing:'0.05em',display:'block',marginBottom:10}}>Unidad de negocio</label>
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:8}}>
              {USUARIOS.map(u=>(
                <button key={u.id} onClick={()=>{setSelUser(u.id);setPass('');setError('')}}
                  style={{padding:'10px 14px',borderRadius:10,border:'2px solid '+(selUser===u.id?u.color:'var(--border2)'),background:selUser===u.id?'rgba('+hexToRgb(u.color)+',0.1)':'var(--bg3)',color:selUser===u.id?u.color:'var(--text2)',cursor:'pointer',fontSize:13,fontWeight:selUser===u.id?600:400,fontFamily:'var(--font)',textAlign:'left',transition:'all 0.15s'}}>
                  <div style={{display:'flex',alignItems:'center',gap:7}}>
                    <div style={{width:8,height:8,borderRadius:'50%',background:u.color,flexShrink:0}}/>
                    {u.nombre}
                  </div>
                </button>
              ))}
            </div>
          </div>
          {selUser&&(
            <div style={{display:'flex',flexDirection:'column',gap:10}}>
              <div>
                <label style={{fontSize:11,color:'var(--text2)',fontWeight:500,textTransform:'uppercase',letterSpacing:'0.05em',display:'block',marginBottom:6}}>Contraseña</label>
                <input type="password" value={pass} onChange={e=>setPass(e.target.value)} onKeyDown={e=>e.key==='Enter'&&intentar()} placeholder="Ingresa tu contraseña" autoFocus style={{width:'100%',padding:'10px 14px',fontSize:14,borderRadius:10}}/>
              </div>
              {error&&<div style={{fontSize:12,color:'var(--red)',textAlign:'center'}}>{error}</div>}
              <button onClick={intentar} style={{background:'var(--accent)',color:'#fff',border:'none',borderRadius:10,padding:'11px',fontSize:14,fontWeight:600,cursor:'pointer',fontFamily:'var(--font)'}}>Entrar →</button>
            </div>
          )}
        </div>
      </div>
    </div>
  )
}

// ═══════════════════════════════════════════════════════
// DASHBOARD CONSOLIDADO — LÍDER
// ═══════════════════════════════════════════════════════

function TareasEquipoPaola({ asignadas, personales, onRefresh }) {
  const [tipo,setTipo] = useState('asignadas')
  const [editando,setEditando] = useState(null)
  const [guardando,setGuardando] = useState(false)
  const lista = tipo==='asignadas'?asignadas:personales

  const guardar = async form => {
    if(!editando) return
    setGuardando(true)
    const actualizada = {...editando,...form}
    const tipoSheet = tipo==='asignadas'?'tareasAsignadas':'misTareas'
    try {
      await eliminarFilaSheets('distribucion',tipoSheet,actualizada.id)
      await insertarFilaSheets('distribucion',tipoSheet,actualizada)

      const uidAnterior = obtenerUidResponsable(editando.responsable)
      const uidNuevo = obtenerUidResponsable(actualizada.responsable)

      if(uidAnterior!=='distribucion' && uidAnterior!==uidNuevo) {
        await eliminarFilaSheets(uidAnterior,tipoSheet,actualizada.id)
      }
      if(uidNuevo!=='distribucion') {
        await eliminarFilaSheets(uidNuevo,tipoSheet,actualizada.id)
        await insertarFilaSheets(uidNuevo,tipoSheet,actualizada)
      }

      if(tipo==='asignadas' && normalizarNombreResponsable(editando.responsable)!==normalizarNombreResponsable(actualizada.responsable)) {
        const email = obtenerCorreoResponsable(actualizada.responsable)
        if(email) {
          await enviarViaJsonp(SHEETS_CONFIG.distribucion.url,{
            accion:'enviar_correo_tarea',
            destinatario:email,copia:EMAIL_LIDER_PROY,
            responsable:actualizada.responsable,comercial:actualizada.responsable,
            tarea:actualizada.nombre,producto:actualizada.nombre,
            descripcion:actualizada.descripcion||'',notas:actualizada.descripcion||'',
            fechaLimite:actualizada.fechaLimite||'',mes:actualizada.fechaLimite||'',
            fecha:new Date().toLocaleDateString('es-CO')
          })
        }
      }
      setEditando(null)
      await onRefresh()
    } catch(e) {
      console.error(e)
      alert('No se pudo guardar la edición.')
    } finally {
      setGuardando(false)
    }
  }

  return (
    <div style={{display:'flex',flexDirection:'column',gap:14}}>
      <div style={{display:'flex',gap:8}}>
        <button onClick={()=>setTipo('asignadas')} style={S.btn(tipo==='asignadas'?'var(--accent)':'var(--bg3)',tipo==='asignadas'?'#fff':'var(--text2)')}>
          Asignadas por Paola · {asignadas.length}
        </button>
        <button onClick={()=>setTipo('personales')} style={S.btn(tipo==='personales'?'var(--accent)':'var(--bg3)',tipo==='personales'?'#fff':'var(--text2)')}>
          Personales del equipo · {personales.length}
        </button>
      </div>

      <div style={S.card}>
        <div style={{padding:12,display:'flex',flexDirection:'column',gap:8}}>
          {lista.length===0&&<div style={{padding:30,textAlign:'center',color:'var(--text3)'}}>No hay tareas registradas.</div>}
          {lista.map(t=>(
            <div key={t.id} style={{background:'var(--bg3)',border:'1px solid var(--border2)',borderRadius:10,padding:'11px 14px',display:'flex',gap:12,alignItems:'center'}}>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontWeight:700,fontSize:13}}>{t.nombre}</div>
                <div style={{fontSize:11,color:'var(--text3)',marginTop:4,display:'flex',gap:10,flexWrap:'wrap'}}>
                  <span>👤 {t.responsable||'Sin responsable'}</span>
                  {t.fechaLimite&&<span>📅 {String(t.fechaLimite).slice(0,10)}</span>}
                  <span>{t.estado||'Pendiente'}</span>
                  {Number(t.tiempoMinutos)>0&&<span>⏱️ {formatoTiempoTarea(t)}</span>}
                </div>
              </div>
              <button onClick={()=>setEditando(t)} style={{...S.btn('var(--accent-soft)','var(--accent2)'),padding:'5px 8px'}}><Edit2 size={12}/> Editar</button>
            </div>
          ))}
        </div>
      </div>

      {editando&&(
        <ModalEditarTarea
          tarea={editando}
          onClose={()=>setEditando(null)}
          onSave={guardar}
          guardando={guardando}
          permiteResponsable={tipo==='asignadas'}
          titulo={tipo==='asignadas'?'Editar tarea asignada':'Editar tarea personal'}
        />
      )}
    </div>
  )
}


function AsignarTareasEquipoPaola() {
  const [tareasCentral,setTareasCentral] = useState({asignadas:[],personales:[]})
  const [modal,setModal] = useState(false)
  const [guardando,setGuardando] = useState(false)
  const [form,setForm] = useState({
    nombre:'',
    descripcion:'',
    responsable:'',
    areas:[],
    fechaLimite:'',
    estado:'Pendiente'
  })

  const refrescar = async () => {
    const d = await cargarDesdeSheets('distribucion')
    if(d) {
      setTareasCentral({
        asignadas:d.tareasAsignadas||[],
        personales:d.misTareas||[],
      })
    }
  }

  useEffect(()=>{ refrescar() },[])

  const crear = async () => {
    if(guardando) return
    if(!form.nombre || !form.responsable) {
      alert('Completa el nombre y el responsable.')
      return
    }

    setGuardando(true)

    const nueva = {
      id:Date.now(),
      ...form,
      asignadoPor:'Paola',
      fechaCreacion:new Date().toLocaleString('es-CO'),
      fechaFinalizacion:'',
      tiempoValor:'',
      tiempoUnidad:'',
      tiempoMinutos:0,
    }

    const destinoUid = obtenerUidResponsable(nueva.responsable)

    try {
      // Copia central de Paola.
      await insertarFilaSheets('distribucion','tareasAsignadas',nueva)

      // Copia en el archivo del responsable.
      if(destinoUid!=='distribucion') {
        await insertarFilaSheets(destinoUid,'tareasAsignadas',nueva)
      }

      // Correo al responsable.
      const email = obtenerCorreoResponsable(nueva.responsable)
      if(email) {
        await enviarViaJsonp(SHEETS_CONFIG.distribucion.url,{
          accion:'enviar_correo_tarea',
          destinatario:email,
          copia:EMAIL_LIDER_PROY,
          comercial:nueva.responsable,
          responsable:nueva.responsable,
          tarea:nueva.nombre,
          producto:nueva.nombre,
          notas:nueva.descripcion||'',
          descripcion:nueva.descripcion||'',
          fechaLimite:nueva.fechaLimite||'',
          mes:nueva.fechaLimite||'',
          fecha:new Date().toLocaleDateString('es-CO'),
          areas:(nueva.areas||[]).join(', '),
        })
      }

      await refrescar()
      setModal(false)
      setForm({nombre:'',descripcion:'',responsable:'',areas:[],fechaLimite:'',estado:'Pendiente'})
    } catch(e) {
      console.error(e)
      alert('No se pudo guardar la tarea.')
    } finally {
      setGuardando(false)
    }
  }

  return (
    <div style={{display:'flex',flexDirection:'column',gap:16}}>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',gap:12,flexWrap:'wrap'}}>
        <div>
          <h2 style={{fontSize:18,margin:0}}>Asignar tareas al equipo</h2>
          <p style={{fontSize:12,color:'var(--text3)',margin:'4px 0 0'}}>
            Paola asigna tareas independientes y consulta el consolidado del equipo.
          </p>
        </div>
        <div style={{display:'flex',gap:8}}>
          <button onClick={refrescar} style={S.btn('var(--bg3)','var(--text2)')}>↻ Actualizar</button>
          <button onClick={()=>setModal(true)} style={S.btn('var(--accent)','#fff')}><PlusCircle size={14}/> Nueva tarea</button>
        </div>
      </div>

      <TareasEquipoPaola
        asignadas={tareasCentral.asignadas}
        personales={tareasCentral.personales}
        onRefresh={refrescar}
      />

      {modal&&(
        <Modal title="Asignar nueva tarea" onClose={()=>!guardando&&setModal(false)} wide>
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:14}}>
            <Field label="Nombre *" span>
              <input value={form.nombre} onChange={e=>setForm({...form,nombre:e.target.value})}/>
            </Field>
            <Field label="Descripción" span>
              <textarea rows={3} value={form.descripcion} onChange={e=>setForm({...form,descripcion:e.target.value})}/>
            </Field>
            <Field label="Responsable *">
              <select value={form.responsable} onChange={e=>setForm({...form,responsable:e.target.value})}>
                <option value="">Selecciona responsable...</option>
                {RESPONSABLES.map(r=><option key={r}>{r}</option>)}
              </select>
            </Field>
            <Field label="Fecha límite">
              <input type="date" value={form.fechaLimite} onChange={e=>setForm({...form,fechaLimite:e.target.value})}/>
            </Field>
            <Field label="Estado">
              <select value={form.estado} onChange={e=>setForm({...form,estado:e.target.value})}>
                {['Pendiente','En proceso'].map(s=><option key={s}>{s}</option>)}
              </select>
            </Field>
            <Field label="Áreas que intervienen" span>
              <div style={{display:'flex',gap:8,flexWrap:'wrap'}}>
                {AREAS.map(a=>(
                  <label key={a} style={{
                    fontSize:12,padding:'6px 10px',border:'1px solid var(--border2)',
                    borderRadius:7,cursor:'pointer',
                    background:form.areas.includes(a)?'var(--accent-soft)':'var(--bg3)'
                  }}>
                    <input
                      type="checkbox"
                      style={{display:'none'}}
                      checked={form.areas.includes(a)}
                      onChange={e=>setForm({
                        ...form,
                        areas:e.target.checked?[...form.areas,a]:form.areas.filter(x=>x!==a)
                      })}
                    />
                    {a}
                  </label>
                ))}
              </div>
            </Field>
          </div>

          <div style={{display:'flex',justifyContent:'flex-end',gap:8,marginTop:20}}>
            <button disabled={guardando} onClick={()=>setModal(false)} style={S.btn('var(--bg3)','var(--text2)')}>Cancelar</button>
            <button disabled={guardando} onClick={crear}
              style={S.btn(guardando?'var(--bg3)':'var(--accent)',guardando?'var(--text3)':'#fff')}>
              {guardando?'⏳ Guardando tarea...':'Asignar tarea'}
            </button>
          </div>
        </Modal>
      )}
    </div>
  )
}

function NotificacionesEquipoPaola() {
  const [datos,setDatos] = useState({asignadas:[],personales:[]})
  const [cargando,setCargando] = useState(true)

  const refrescar = async () => {
    setCargando(true)
    const d = await cargarDesdeSheets('distribucion')
    if(d) setDatos({asignadas:d.tareasAsignadas||[],personales:d.misTareas||[]})
    setCargando(false)
  }

  useEffect(()=>{ refrescar() },[])

  const movimientos = [
    ...datos.asignadas.map(t=>({
      id:'a-'+t.id,
      titulo:t.estado==='Finalizada'?'Tarea finalizada':'Tarea asignada por Paola',
      detalle:`${t.responsable||'Sin responsable'} · ${t.nombre||'Tarea'}`,
      fecha:t.fechaFinalizacion||t.fechaCreacion||'',
      estado:t.estado||'Pendiente',
      finalizada:t.estado==='Finalizada'
    })),
    ...datos.personales.map(t=>({
      id:'p-'+t.id,
      titulo:t.estado==='Finalizada'?'Tarea personal finalizada':'Tarea personal creada',
      detalle:`${t.responsable||'Equipo'} · ${t.nombre||'Tarea'}`,
      fecha:t.fechaFinalizacion||t.fechaCreacion||'',
      estado:t.estado||'Pendiente',
      finalizada:t.estado==='Finalizada'
    })),
  ]

  return (
    <div style={{display:'flex',flexDirection:'column',gap:16}}>
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center'}}>
        <div>
          <h2 style={{fontSize:18,margin:0}}>Notificaciones</h2>
          <p style={{fontSize:12,color:'var(--text3)',margin:'4px 0 0'}}>Movimientos registrados en las tareas del equipo.</p>
        </div>
        <button disabled={cargando} onClick={refrescar} style={S.btn('var(--bg3)','var(--text2)')}>
          {cargando?'⏳ Cargando...':'↻ Actualizar'}
        </button>
      </div>

      <div style={S.card}>
        <div style={{padding:12,display:'flex',flexDirection:'column',gap:8}}>
          {!cargando&&movimientos.length===0&&(
            <div style={{padding:35,textAlign:'center',color:'var(--text3)'}}>Aún no hay movimientos registrados.</div>
          )}
          {movimientos.map(n=>(
            <div key={n.id} style={{
              display:'flex',gap:12,alignItems:'center',padding:'11px 12px',
              border:'1px solid var(--border2)',borderRadius:10,background:'var(--bg3)'
            }}>
              <div style={{
                width:34,height:34,borderRadius:10,display:'grid',placeItems:'center',
                background:n.finalizada?'var(--green-soft)':'var(--accent-soft)',
                color:n.finalizada?'var(--green)':'var(--accent2)',fontWeight:800
              }}>
                {n.finalizada?'✓':'!'}
              </div>
              <div style={{flex:1}}>
                <div style={{fontSize:13,fontWeight:700}}>{n.titulo}</div>
                <div style={{fontSize:12,color:'var(--text2)',marginTop:2}}>{n.detalle}</div>
              </div>
              <div style={{textAlign:'right'}}>
                <div style={{fontSize:11,fontWeight:700,color:n.finalizada?'var(--green)':'var(--orange)'}}>{n.estado}</div>
                {n.fecha&&<div style={{fontSize:10,color:'var(--text3)',marginTop:3}}>{String(n.fecha)}</div>}
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  )
}

function DashboardLider() {
  const [tabLider, setTabLider] = useState('dashboard')
  const [filtroUnidad, setFiltroUnidad] = useState('')
  const [filtroMes, setFiltroMes] = useState('')
  const [modalPendiente, setModalPendiente] = useState(false)
  const [guardandoPendiente, setGuardandoPendiente] = useState(false)
  const [tareasCentral, setTareasCentral] = useState({asignadas:[],personales:[]})
  const [formPend, setFormPend] = useState({nombre:'',descripcion:'',responsable:'',areas:[],fechaLimite:'',estado:'Pendiente'})

  const refrescarTareasCentral = async () => {
    const d = await cargarDesdeSheets('distribucion')
    if(d) setTareasCentral({asignadas:d.tareasAsignadas||[],personales:d.misTareas||[]})
  }

  useEffect(()=>{ refrescarTareasCentral() },[])

  const unidades = USUARIOS.filter(u=>u.rol==='normal')
  const allData = unidades.map(u=>{
    try { const d=localStorage.getItem(getStorageKey(u.id)); return d?{...JSON.parse(d),_id:u.id,_unidad:u.nombre,_color:u.color}:{inversiones:[],ventas:[],planes:[],presupuestos:[],pendientes:[],gastosPresupuesto:[],_id:u.id,_unidad:u.nombre,_color:u.color} } catch { return {inversiones:[],ventas:[],planes:[],presupuestos:[],pendientes:[],gastosPresupuesto:[],_id:u.id,_unidad:u.nombre,_color:u.color} }
  })

  const filtered = filtroUnidad ? allData.filter(d=>d._id===filtroUnidad) : allData

  const totalInv = filtered.reduce((s,d)=>s+d.inversiones.filter(i=>!filtroMes||normMes(i.mes)===normMes(filtroMes)).reduce((ss,i)=>ss+(Number(i.inversion)||0),0),0)
  const totalVenta = filtered.reduce((s,d)=>s+d.ventas.filter(v=>!filtroMes||normMes(v.mes)===normMes(filtroMes)).reduce((ss,v)=>ss+(Number(v.ventaNeta)||0),0),0)
  const totalPres = filtered.reduce((s,d)=>s+d.presupuestos.filter(p=>!filtroMes||normMes(p.mes)===normMes(filtroMes)).reduce((ss,p)=>ss+(Number(p.monto)||0),0),0)
  const totalGastos = filtered.reduce((s,d)=>s+(d.gastosPresupuesto||[]).filter(g=>!filtroMes||normMes(g.mes)===normMes(filtroMes)).reduce((ss,g)=>ss+(Number(g.valorFactura)||0),0),0)
  const roiPct = totalVenta>0?(totalInv/totalVenta)*100:0

  const porUnidad = filtered.map(d=>{
    const inv=d.inversiones.filter(i=>!filtroMes||normMes(i.mes)===normMes(filtroMes)).reduce((s,i)=>s+(Number(i.inversion)||0),0)
    const venta=d.ventas.filter(v=>!filtroMes||normMes(v.mes)===normMes(filtroMes)).reduce((s,v)=>s+(Number(v.ventaNeta)||0),0)
    const pres=d.presupuestos.filter(p=>!filtroMes||normMes(p.mes)===normMes(filtroMes)).reduce((s,p)=>s+(Number(p.monto)||0),0)
    const gastos=(d.gastosPresupuesto||[]).filter(g=>!filtroMes||normMes(g.mes)===normMes(filtroMes)).reduce((s,g)=>s+(Number(g.valorFactura)||0),0)
    return {unidad:d._unidad,color:d._color,id:d._id,inv,venta,pres,gastos,pctInv:venta>0?(inv/venta)*100:0,pctEjec:pres>0?(gastos/pres)*100:0,pendientes:d.pendientes.filter(p=>p.estado!=='Listo'&&p.estado!=='Cancelado').length}
  })

  const porMesChart = MESES.map(m=>({
    name:m.slice(0,3),
    inv: filtered.reduce((s,d)=>s+d.inversiones.filter(i=>normMes(i.mes)===m).reduce((ss,i)=>ss+(Number(i.inversion)||0),0),0),
    venta: filtered.reduce((s,d)=>s+d.ventas.filter(v=>normMes(v.mes)===m).reduce((ss,v)=>ss+(Number(v.ventaNeta)||0),0),0),
    gastos: filtered.reduce((s,d)=>s+(d.gastosPresupuesto||[]).filter(g=>normMes(g.mes)===m).reduce((ss,g)=>ss+(Number(g.valorFactura)||0),0),0),
  })).filter(m=>m.inv>0||m.venta>0||m.gastos>0)

  const clientesMap = {}
  filtered.forEach(d=>{
    d.inversiones.filter(i=>!filtroMes||normMes(i.mes)===normMes(filtroMes)).forEach(i=>{
      const k = i.distribuidor; if(!k) return
      if(!clientesMap[k]) clientesMap[k]={nombre:k,inv:0,venta:0,unidad:d._unidad,color:d._color}
      clientesMap[k].inv += Number(i.inversion)||0
    })
    d.ventas.filter(v=>!filtroMes||normMes(v.mes)===normMes(filtroMes)).forEach(v=>{
      const k = v.distribuidor; if(!k) return
      if(!clientesMap[k]) clientesMap[k]={nombre:k,inv:0,venta:0,unidad:d._unidad,color:d._color}
      clientesMap[k].venta += Number(v.ventaNeta)||0
    })
  })
  const topClientes = Object.values(clientesMap).sort((a,b)=>b.inv-a.inv).slice(0,15)

  const todasActividades = filtered.flatMap(d=>d.pendientes.map(p=>({...p,_unidad:d._unidad,_color:d._color,_uid:d._id}))).sort((a,b)=>['Alta','Media','Baja'].indexOf(a.prioridad)-['Alta','Media','Baja'].indexOf(b.prioridad))
  const actFiltradas = todasActividades.filter(a=>!filtroUnidad||a._uid===filtroUnidad)

  const EMAILS_UNIDAD_LIDER = {
    industria:    'epabon@gulfcolombia.com',
    distribucion: 'msilva@prolub.com.co',
    zonas:        'msilva@prolub.com.co',
    diseno:       'aalvarado@gulfcolombia.com',
  }
  const SHEETS_URL_LIDER = 'https://script.google.com/macros/s/AKfycbxzqhKZr7ue545BDwpgqQrNsognsVOtxxZNPd6P4Jm_wtdj1_Sqkd9FdswrdtfM2d8S/exec'

  const crearPendiente = async () => {
    if(guardandoPendiente) return
    if(!formPend.nombre||!formPend.responsable) return alert('Completa nombre y responsable.')
    setGuardandoPendiente(true)

    const nuevo = {
      id:Date.now(),...formPend,asignadoPor:'Paola',
      fechaCreacion:new Date().toLocaleString('es-CO'),
      fechaFinalizacion:'',tiempoValor:'',tiempoUnidad:'',tiempoMinutos:0
    }
    const destinoUid = obtenerUidResponsable(nuevo.responsable)

    try {
      await insertarFilaSheets('distribucion','tareasAsignadas',nuevo)
      if(destinoUid!=='distribucion') await insertarFilaSheets(destinoUid,'tareasAsignadas',nuevo)

      const email = obtenerCorreoResponsable(nuevo.responsable)
      if(email) {
        await enviarViaJsonp(SHEETS_CONFIG.distribucion.url,{
          accion:'enviar_correo_tarea',
          destinatario:email,copia:EMAIL_LIDER_PROY,
          comercial:nuevo.responsable,responsable:nuevo.responsable,
          tarea:nuevo.nombre,producto:nuevo.nombre,
          notas:nuevo.descripcion||'',descripcion:nuevo.descripcion||'',
          fechaLimite:nuevo.fechaLimite||'',mes:nuevo.fechaLimite||'',
          fecha:new Date().toLocaleDateString('es-CO'),
          areas:(nuevo.areas||[]).join(', '),
        })
      }

      await refrescarTareasCentral()
      setModalPendiente(false)
      setFormPend({nombre:'',descripcion:'',responsable:'',areas:[],fechaLimite:'',estado:'Pendiente'})
    } catch(e) {
      console.error(e)
      alert('Error guardando la tarea: '+e.message)
    } finally { setGuardandoPendiente(false) }
  }

  const CT2 = ({active,payload,label}) => {
    if(!active||!payload?.length) return null
    return <div style={{background:'var(--bg3)',border:'1px solid var(--border2)',borderRadius:10,padding:'10px 14px',fontSize:12}}>
      <p style={{fontWeight:600,marginBottom:4}}>{label}</p>
      {payload.map((p,i)=><p key={i} style={{color:p.color}}>{p.name}: {cop(p.value)}</p>)}
    </div>
  }

  const TABS_L = [
    {id:'dashboard',label:'Dashboard'},
    {id:'tiempos',label:'Tiempos del equipo'},
    {id:'clientes',label:'Por Cliente'},
    {id:'presupuesto',label:'Presupuesto'}
  ]

  const tareasMedibles = [...tareasCentral.asignadas,...tareasCentral.personales]
    .filter(t=>t.estado==='Finalizada'&&Number(t.tiempoMinutos||0)>0)

  const notificacionesPaola = [
    ...tareasCentral.asignadas.map(t=>({
      id:'a-'+t.id,
      tipo:t.estado==='Finalizada'?'finalizada':'asignada',
      titulo:t.estado==='Finalizada'?'Tarea finalizada':'Tarea asignada',
      detalle:`${t.responsable||'Sin responsable'} · ${t.nombre||'Tarea'}`,
      fecha:t.fechaFinalizacion||t.fechaCreacion||'',
      estado:t.estado||'Pendiente'
    })),
    ...tareasCentral.personales.map(t=>({
      id:'p-'+t.id,
      tipo:t.estado==='Finalizada'?'finalizada':'personal',
      titulo:t.estado==='Finalizada'?'Tarea personal finalizada':'Nueva tarea personal',
      detalle:`${t.responsable||'Equipo'} · ${t.nombre||'Tarea'}`,
      fecha:t.fechaFinalizacion||t.fechaCreacion||'',
      estado:t.estado||'Pendiente'
    }))
  ].sort((a,b)=>String(b.fecha||'').localeCompare(String(a.fecha||'')))

  const resumenTiempos = RESPONSABLES.map(nombre=>{
    const lista = tareasMedibles.filter(t=>normalizarNombreResponsable(t.responsable)===normalizarNombreResponsable(nombre))
    const promedio = lista.length?lista.reduce((s,t)=>s+Number(t.tiempoMinutos||0),0)/lista.length:0
    return {nombre,cerradas:lista.length,promedio}
  }).filter(r=>r.cerradas>0)

  return (
    <div style={{display:'flex',flexDirection:'column',gap:20}}>
      <div style={{display:'flex',gap:8,alignItems:'center',flexWrap:'wrap'}}>
        <div style={{display:'flex',gap:4,background:'var(--bg2)',border:'1px solid var(--border)',borderRadius:10,padding:4}}>
          {TABS_L.map(t=>(
            <button key={t.id} onClick={()=>setTabLider(t.id)}
              style={{...S.btn(tabLider===t.id?'var(--accent)':'transparent',tabLider===t.id?'#fff':'var(--text2)'),padding:'6px 16px',fontSize:12,borderRadius:7}}>
              {t.label}
            </button>
          ))}
        </div>
        <select value={filtroUnidad} onChange={e=>setFiltroUnidad(e.target.value)} style={{width:190}}>
          <option value="">Todas las unidades</option>
          {unidades.map(u=><option key={u.id} value={u.id}>{u.nombre}</option>)}
        </select>
        <select value={filtroMes} onChange={e=>setFiltroMes(e.target.value)} style={{width:140}}>
          <option value="">Todos los meses</option>
          {MESES.map(m=><option key={m}>{m}</option>)}
        </select>
        {(filtroUnidad||filtroMes)&&<button onClick={()=>{setFiltroUnidad('');setFiltroMes('')}} style={{...S.btn('var(--bg3)','var(--text2)'),padding:'5px 10px',fontSize:12}}>✕</button>}
        
      </div>

      <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(190px,1fr))',gap:12,overflow:'visible'}}>
        <KpiCard icon={TrendingUp} label="Inversión total" value={cop(totalInv)} sub={filtroMes||'Acumulado'} accent="var(--accent2)"/>
        <KpiCard icon={ShoppingCart} label="Venta neta" value={cop(totalVenta)} accent="var(--green)"/>
        <KpiCard icon={BarChart2} label="% Inv/Venta" value={roiPct.toFixed(1)+'%'} accent={roiPct>15?'var(--red)':roiPct>10?'var(--yellow)':'var(--green)'}/>
        <KpiCard icon={DollarSign} label="Presupuesto" value={cop(totalPres)} sub={totalPres>0?((totalGastos/totalPres)*100).toFixed(1)+'% ejec.':''}/>
        <KpiCard icon={DollarSign} label="Gastado presup." value={cop(totalGastos)} accent={totalGastos>totalPres?'var(--red)':'var(--text)'}/>
        <KpiCard icon={ListTodo} label="Actividades abiertas" value={actFiltradas.filter(a=>a.estado!=='Listo'&&a.estado!=='Cancelado').length} accent="var(--yellow)"/>
      </div>

      {tabLider==='dashboard'&&(
        <div style={{display:'flex',flexDirection:'column',gap:16}}>
          <div style={{display:'grid',gridTemplateColumns:'1.4fr 1fr',gap:16}}>
            <div style={{...S.card,padding:20}}>
              <h4 style={{fontSize:11,fontWeight:600,color:'var(--text2)',textTransform:'uppercase',letterSpacing:'0.05em',marginBottom:14}}>Inversión vs Venta por mes</h4>
              <ResponsiveContainer width="100%" height={200}>
                <BarChart data={porMesChart} barGap={3}>
                  <XAxis dataKey="name" tick={{fill:'var(--text3)',fontSize:11}} axisLine={false} tickLine={false}/>
                  <YAxis tickFormatter={v=>(v/1000000).toFixed(0)+'M'} tick={{fill:'var(--text3)',fontSize:10}} axisLine={false} tickLine={false}/>
                  <Tooltip content={<CT2/>}/>
                  <Bar dataKey="venta" name="Venta neta" fill="var(--green)" radius={[4,4,0,0]} opacity={0.7}/>
                  <Bar dataKey="inv" name="Inversión" fill="var(--accent)" radius={[4,4,0,0]}/>
                  <Bar dataKey="gastos" name="Gastos presup." fill="var(--orange)" radius={[4,4,0,0]}/>
                </BarChart>
              </ResponsiveContainer>
            </div>
            <div style={{...S.card,padding:20}}>
              <h4 style={{fontSize:11,fontWeight:600,color:'var(--text2)',textTransform:'uppercase',letterSpacing:'0.05em',marginBottom:14}}>Inversión por unidad</h4>
              <ResponsiveContainer width="100%" height={200}>
                <PieChart>
                  <Pie data={porUnidad.filter(u=>u.inv>0)} cx="50%" cy="50%" innerRadius={45} outerRadius={75} dataKey="inv" nameKey="unidad" paddingAngle={3}>
                    {porUnidad.filter(u=>u.inv>0).map((u,i)=><Cell key={i} fill={u.color}/>)}
                  </Pie>
                  <Tooltip formatter={v=>cop(v)} contentStyle={{background:'var(--bg3)',border:'1px solid var(--border2)',borderRadius:10,fontSize:11}}/>
                  <Legend iconSize={7} iconType="circle" wrapperStyle={{fontSize:10,color:'var(--text2)'}}/>
                </PieChart>
              </ResponsiveContainer>
            </div>
          </div>
          <div style={S.card}>
            <div style={{padding:'12px 18px',borderBottom:'1px solid var(--border)'}}><h4 style={{fontSize:11,fontWeight:600,color:'var(--text2)',textTransform:'uppercase',letterSpacing:'0.05em'}}>Resumen por unidad de negocio</h4></div>
            <table style={{width:'100%',borderCollapse:'collapse'}}>
              <thead><tr>{['Unidad','Inversión','Venta Neta','% Inv/Venta','Presupuesto','Gastado','% Ejec.','Pendientes'].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
              <tbody>
                {porUnidad.map((u,i)=>(
                  <tr key={i} onClick={()=>setFiltroUnidad(filtroUnidad===u.id?'':u.id)} style={{cursor:'pointer',background:filtroUnidad===u.id?'rgba(108,99,255,0.06)':'transparent'}}>
                    <td style={{...S.td,fontWeight:600}}><div style={{display:'flex',alignItems:'center',gap:8}}><div style={{width:10,height:10,borderRadius:'50%',background:u.color}}/>{u.unidad}</div></td>
                    <td style={{...S.td,fontFamily:'var(--mono)'}}>{cop(u.inv)}</td>
                    <td style={{...S.td,fontFamily:'var(--mono)',color:'var(--green)'}}>{cop(u.venta)}</td>
                    <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:600,color:u.pctInv>15?'var(--red)':u.pctInv>10?'var(--yellow)':'var(--green)'}}>{u.pctInv.toFixed(1)}%</td>
                    <td style={{...S.td,fontFamily:'var(--mono)',color:'var(--text2)'}}>{cop(u.pres)}</td>
                    <td style={{...S.td,fontFamily:'var(--mono)'}}>{cop(u.gastos)}</td>
                    <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:600,color:u.pctEjec>100?'var(--red)':u.pctEjec>80?'var(--yellow)':'var(--green)'}}>{u.pres>0?u.pctEjec.toFixed(1)+'%':'—'}</td>
                    <td style={{...S.td,textAlign:'center'}}>{u.pendientes>0?<span style={{background:'var(--yellow-soft)',color:'var(--yellow)',padding:'2px 8px',borderRadius:6,fontSize:12,fontWeight:600}}>{u.pendientes}</span>:'—'}</td>
                  </tr>
                ))}
                <tr style={{borderTop:'2px solid var(--border2)',background:'var(--bg3)'}}>
                  <td style={{...S.td,fontWeight:700}}>TOTAL</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700,color:'var(--accent2)'}}>{cop(totalInv)}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700,color:'var(--green)'}}>{cop(totalVenta)}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700}}>{roiPct.toFixed(1)}%</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700}}>{cop(totalPres)}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700,color:totalGastos>totalPres?'var(--red)':'var(--text)'}}>{cop(totalGastos)}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700}}>{totalPres>0?((totalGastos/totalPres)*100).toFixed(1)+'%':'—'}</td>
                  <td style={S.td}/>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      )}

      {tabLider==='tiempos'&&(
        <div style={{display:'flex',flexDirection:'column',gap:16}}>
          <div style={{display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:12}}>
            <KpiCard icon={Check} label="Tareas medidas" value={tareasMedibles.length} accent="var(--green)"/>
            <KpiCard icon={ListTodo} label="Asignadas finalizadas" value={tareasCentral.asignadas.filter(t=>t.estado==='Finalizada').length} accent="var(--orange)"/>
            <KpiCard icon={ListTodo} label="Personales finalizadas" value={tareasCentral.personales.filter(t=>t.estado==='Finalizada').length} accent="var(--accent2)"/>
          </div>
          <div style={S.card}>
            <table style={{width:'100%',borderCollapse:'collapse'}}>
              <thead><tr>{['Persona','Tareas cerradas','Tiempo promedio'].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
              <tbody>
                {resumenTiempos.length===0&&<tr><td colSpan={3} style={{...S.td,textAlign:'center',padding:35,color:'var(--text3)'}}>Aún no hay tareas finalizadas con tiempo reportado.</td></tr>}
                {resumenTiempos.map(r=>(
                  <tr key={r.nombre}>
                    <td style={{...S.td,fontWeight:700}}>{r.nombre}</td>
                    <td style={S.td}>{r.cerradas}</td>
                    <td style={{...S.td,fontFamily:'var(--mono)',color:'var(--accent2)'}}>{formatoTiempoTarea({tiempoMinutos:r.promedio})}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {tabLider==='clientes'&&(
        <div style={S.card}>
          <div style={{padding:'12px 18px',borderBottom:'1px solid var(--border)'}}><h4 style={{fontSize:11,fontWeight:600,color:'var(--text2)',textTransform:'uppercase',letterSpacing:'0.05em'}}>Inversión por cliente/distribuidor</h4></div>
          <table style={{width:'100%',borderCollapse:'collapse'}}>
            <thead><tr>{['Distribuidor','Unidad','Inversión','Venta Neta','% Inv/Venta','Participación'].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
            <tbody>
              {topClientes.length===0&&<tr><td colSpan={6} style={{...S.td,textAlign:'center',color:'var(--text3)',padding:32}}>Sin datos</td></tr>}
              {topClientes.map((c,i)=>{
                const pct = c.venta>0?(c.inv/c.venta)*100:0
                return (
                  <tr key={i}>
                    <td style={{...S.td,fontWeight:500}}>{c.nombre}</td>
                    <td style={{...S.td}}><div style={{display:'flex',alignItems:'center',gap:6}}><div style={{width:8,height:8,borderRadius:'50%',background:c.color}}/><span style={{fontSize:12,color:'var(--text2)'}}>{c.unidad}</span></div></td>
                    <td style={{...S.td,fontFamily:'var(--mono)'}}>{cop(c.inv)}</td>
                    <td style={{...S.td,fontFamily:'var(--mono)',color:'var(--green)'}}>{c.venta>0?cop(c.venta):'—'}</td>
                    <td style={{...S.td}}>{c.venta>0?<span style={{fontFamily:'var(--mono)',fontWeight:600,color:pct>15?'var(--red)':pct>10?'var(--yellow)':'var(--green)'}}>{pct.toFixed(1)}%</span>:'—'}</td>
                    <td style={{...S.td,minWidth:130}}>
                      <div style={{display:'flex',alignItems:'center',gap:7}}>
                        <div style={{flex:1,height:5,background:'var(--bg4)',borderRadius:3}}><div style={{width:Math.min(totalInv>0?(c.inv/totalInv)*100:0,100)+'%',height:'100%',background:COLORES[i%COLORES.length],borderRadius:3}}/></div>
                        <span style={{fontSize:11,color:'var(--text2)',minWidth:34}}>{totalInv>0?((c.inv/totalInv)*100).toFixed(1):0}%</span>
                      </div>
                    </td>
                  </tr>
                )
              })}
              {topClientes.length>0&&(
                <tr style={{borderTop:'2px solid var(--border2)',background:'var(--bg3)'}}>
                  <td colSpan={2} style={{...S.td,fontWeight:700}}>TOTAL</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700,color:'var(--accent2)'}}>{cop(totalInv)}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700,color:'var(--green)'}}>{cop(totalVenta)}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700}}>{roiPct.toFixed(1)}%</td>
                  <td style={S.td}/>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {tabLider==='presupuesto'&&(
        <div style={{display:'flex',flexDirection:'column',gap:16}}>
          <div style={S.card}>
            <div style={{padding:'12px 18px',borderBottom:'1px solid var(--border)'}}><h4 style={{fontSize:11,fontWeight:600,color:'var(--text2)',textTransform:'uppercase',letterSpacing:'0.05em'}}>Presupuesto ejecutado vs asignado por mes</h4></div>
            <table style={{width:'100%',borderCollapse:'collapse'}}>
              <thead><tr>{['Mes','Asignado','Ejecutado','Disponible','% Ejec.','Barra'].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
              <tbody>
                {MESES.map(m=>{
                  const pres=filtered.reduce((s,d)=>s+d.presupuestos.filter(p=>normMes(p.mes)===m).reduce((ss,p)=>ss+(Number(p.monto)||0),0),0)
                  const gastado=filtered.reduce((s,d)=>s+(d.gastosPresupuesto||[]).filter(g=>normMes(g.mes)===m).reduce((ss,g)=>ss+(Number(g.valorFactura)||0),0),0)
                  if(pres===0&&gastado===0) return null
                  const ejec=pres>0?(gastado/pres)*100:0
                  return (
                    <tr key={m} onClick={()=>setFiltroMes(filtroMes===m?'':m)} style={{cursor:'pointer',background:filtroMes===m?'rgba(108,99,255,0.06)':'transparent'}}>
                      <td style={{...S.td,fontWeight:600,color:filtroMes===m?'var(--accent2)':'var(--text)'}}>{m}</td>
                      <td style={{...S.td,fontFamily:'var(--mono)',color:'var(--text2)'}}>{cop(pres)}</td>
                      <td style={{...S.td,fontFamily:'var(--mono)'}}>{cop(gastado)}</td>
                      <td style={{...S.td,fontFamily:'var(--mono)',color:pres-gastado>=0?'var(--green)':'var(--red)'}}>{cop(pres-gastado)}</td>
                      <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:600,color:ejec>100?'var(--red)':ejec>80?'var(--yellow)':'var(--green)'}}>{ejec.toFixed(1)}%</td>
                      <td style={{...S.td,minWidth:120}}><div style={{height:5,background:'var(--bg4)',borderRadius:3}}><div style={{width:Math.min(ejec,100)+'%',height:'100%',background:ejec>100?'var(--red)':ejec>80?'var(--yellow)':'var(--green)',borderRadius:3}}/></div></td>
                    </tr>
                  )
                })}
                <tr style={{borderTop:'2px solid var(--border2)',background:'var(--bg3)'}}>
                  <td style={{...S.td,fontWeight:700}}>TOTAL</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700}}>{cop(totalPres)}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700,color:'var(--accent2)'}}>{cop(totalGastos)}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700,color:totalPres-totalGastos>=0?'var(--green)':'var(--red)'}}>{cop(totalPres-totalGastos)}</td>
                  <td colSpan={2} style={S.td}/>
                </tr>
              </tbody>
            </table>
          </div>
          <div style={S.card}>
            <div style={{padding:'12px 18px',borderBottom:'1px solid var(--border)'}}><h4 style={{fontSize:11,fontWeight:600,color:'var(--text2)',textTransform:'uppercase',letterSpacing:'0.05em'}}>Por unidad de negocio</h4></div>
            <table style={{width:'100%',borderCollapse:'collapse'}}>
              <thead><tr>{['Unidad','Presupuesto','Ejecutado','Disponible','% Ejec.'].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
              <tbody>
                {filtered.map((d,i)=>{
                  const pres=d.presupuestos.filter(p=>!filtroMes||normMes(p.mes)===normMes(filtroMes)).reduce((s,p)=>s+(Number(p.monto)||0),0)
                  const gastado=(d.gastosPresupuesto||[]).filter(g=>!filtroMes||normMes(g.mes)===normMes(filtroMes)).reduce((s,g)=>s+(Number(g.valorFactura)||0),0)
                  const ejec=pres>0?(gastado/pres)*100:0
                  return (
                    <tr key={i}>
                      <td style={{...S.td,fontWeight:600}}><div style={{display:'flex',alignItems:'center',gap:8}}><div style={{width:10,height:10,borderRadius:'50%',background:d._color}}/>{d._unidad}</div></td>
                      <td style={{...S.td,fontFamily:'var(--mono)',color:'var(--text2)'}}>{cop(pres)}</td>
                      <td style={{...S.td,fontFamily:'var(--mono)'}}>{cop(gastado)}</td>
                      <td style={{...S.td,fontFamily:'var(--mono)',color:pres-gastado>=0?'var(--green)':'var(--red)'}}>{cop(pres-gastado)}</td>
                      <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:600,color:ejec>100?'var(--red)':ejec>80?'var(--yellow)':'var(--green)'}}>{pres>0?ejec.toFixed(1)+'%':'—'}</td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {modalPendiente&&(
        <Modal title="Asignar nueva tarea" onClose={()=>!guardandoPendiente&&setModalPendiente(false)} wide>
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:14}}>
            <Field label="Nombre *" span><input value={formPend.nombre} onChange={e=>setFormPend({...formPend,nombre:e.target.value})}/></Field>
            <Field label="Descripción" span><textarea rows={3} value={formPend.descripcion} onChange={e=>setFormPend({...formPend,descripcion:e.target.value})}/></Field>
            <Field label="Responsable *">
              <select value={formPend.responsable} onChange={e=>setFormPend({...formPend,responsable:e.target.value})}>
                <option value="">Selecciona responsable...</option>
                {RESPONSABLES.map(r=><option key={r}>{r}</option>)}
              </select>
            </Field>
            <Field label="Fecha límite"><input type="date" value={formPend.fechaLimite} onChange={e=>setFormPend({...formPend,fechaLimite:e.target.value})}/></Field>
            <Field label="Estado"><select value={formPend.estado} onChange={e=>setFormPend({...formPend,estado:e.target.value})}>{['Pendiente','En proceso'].map(s=><option key={s}>{s}</option>)}</select></Field>
            <Field label="Áreas que intervienen" span>
              <div style={{display:'flex',gap:8,flexWrap:'wrap'}}>
                {AREAS.map(a=>(
                  <label key={a} style={{fontSize:12,padding:'6px 10px',border:'1px solid var(--border2)',borderRadius:7,cursor:'pointer',background:formPend.areas.includes(a)?'var(--accent-soft)':'var(--bg3)'}}>
                    <input type="checkbox" style={{display:'none'}} checked={formPend.areas.includes(a)}
                      onChange={e=>setFormPend({...formPend,areas:e.target.checked?[...formPend.areas,a]:formPend.areas.filter(x=>x!==a)})}/>
                    {a}
                  </label>
                ))}
              </div>
            </Field>
          </div>
          <div style={{display:'flex',justifyContent:'flex-end',gap:8,marginTop:20}}>
            <button onClick={()=>setModalPendiente(false)} disabled={guardandoPendiente} style={S.btn('var(--bg3)','var(--text2)')}>Cancelar</button>
            <button onClick={crearPendiente} disabled={guardandoPendiente}
              style={S.btn(guardandoPendiente?'var(--bg3)':'var(--accent)',guardandoPendiente?'var(--text3)':'#fff')}>
              {guardandoPendiente?'⏳ Guardando tarea...':'Asignar tarea'}
            </button>
          </div>
        </Modal>
      )}
    </div>
  )
}

// ═══════════════════════════════════════════════════════
// PRESUPUESTO CONSOLIDADO
// ═══════════════════════════════════════════════════════
function PresupuestoConsolidado() {
  const allData = USUARIOS.filter(u=>u.rol==='normal').map(u=>{
    try { const d=localStorage.getItem(getStorageKey(u.id)); return d?{...JSON.parse(d),_unidad:u.nombre,_color:u.color}:null } catch { return null }
  }).filter(Boolean)

  const porMes = MESES.map(m=>{
    const pres = allData.reduce((s,d)=>s+d.presupuestos.filter(p=>normMes(p.mes)===m).reduce((ss,p)=>ss+(Number(p.monto)||0),0),0)
    const gastado = allData.reduce((s,d)=>s+(d.gastosPresupuesto||[]).filter(g=>normMes(g.mes)===m).reduce((ss,g)=>ss+(Number(g.valorFactura)||0),0),0)
    return {mes:m,pres,gastado,ejec:pres>0?(gastado/pres)*100:0}
  }).filter(r=>r.pres>0||r.gastado>0)

  const totalPres = porMes.reduce((s,r)=>s+r.pres,0)
  const totalGastado = porMes.reduce((s,r)=>s+r.gastado,0)

  return (
    <div style={{display:'flex',flexDirection:'column',gap:20}}>
      <div style={{padding:'12px 18px',background:'rgba(167,139,250,0.1)',borderRadius:12,border:'1px solid rgba(167,139,250,0.2)',fontSize:13,color:'#a78bfa'}}>
        💰 Presupuesto consolidado de todas las unidades de negocio
      </div>
      <div style={{display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:13}}>
        <KpiCard icon={DollarSign} label="Presupuesto total asignado" value={cop(totalPres)} accent="var(--text)"/>
        <KpiCard icon={TrendingUp} label="Total ejecutado" value={cop(totalGastado)} accent="var(--accent2)"/>
        <KpiCard icon={BarChart2} label="% Ejecutado global" value={totalPres>0?((totalGastado/totalPres)*100).toFixed(1)+'%':'—'} accent={totalGastado>totalPres?'var(--red)':'var(--green)'}/>
      </div>
      <div style={S.card}>
        <div style={{padding:'12px 18px',borderBottom:'1px solid var(--border)'}}><h4 style={{fontSize:11,fontWeight:600,color:'var(--text2)',textTransform:'uppercase',letterSpacing:'0.05em'}}>Ejecutado vs Presupuesto por mes — Consolidado</h4></div>
        <table style={{width:'100%',borderCollapse:'collapse'}}>
          <thead><tr>{['Mes','Presupuesto','Ejecutado','Disponible','% Ejec.','Barra'].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
          <tbody>
            {porMes.map((r,i)=>(
              <tr key={i}>
                <td style={{...S.td,fontWeight:600}}>{r.mes}</td>
                <td style={{...S.td,fontFamily:'var(--mono)',color:'var(--text2)'}}>{cop(r.pres)}</td>
                <td style={{...S.td,fontFamily:'var(--mono)'}}>{cop(r.gastado)}</td>
                <td style={{...S.td,fontFamily:'var(--mono)',color:r.pres-r.gastado>=0?'var(--green)':'var(--red)'}}>{cop(r.pres-r.gastado)}</td>
                <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:600,color:r.ejec>100?'var(--red)':r.ejec>80?'var(--yellow)':'var(--green)'}}>{r.ejec.toFixed(1)}%</td>
                <td style={{...S.td,minWidth:120}}><div style={{height:6,background:'var(--bg4)',borderRadius:3}}><div style={{width:Math.min(r.ejec,100)+'%',height:'100%',background:r.ejec>100?'var(--red)':r.ejec>80?'var(--yellow)':'var(--green)',borderRadius:3}}/></div></td>
              </tr>
            ))}
            <tr style={{borderTop:'2px solid var(--border2)',background:'var(--bg3)'}}>
              <td style={{...S.td,fontWeight:700}}>TOTAL</td>
              <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700}}>{cop(totalPres)}</td>
              <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700,color:'var(--accent2)'}}>{cop(totalGastado)}</td>
              <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:700,color:totalPres-totalGastado>=0?'var(--green)':'var(--red)'}}>{cop(totalPres-totalGastado)}</td>
              <td colSpan={2} style={S.td}/>
            </tr>
          </tbody>
        </table>
      </div>
      <div style={S.card}>
        <div style={{padding:'12px 18px',borderBottom:'1px solid var(--border)'}}><h4 style={{fontSize:11,fontWeight:600,color:'var(--text2)',textTransform:'uppercase',letterSpacing:'0.05em'}}>Por unidad de negocio</h4></div>
        <table style={{width:'100%',borderCollapse:'collapse'}}>
          <thead><tr>{['Unidad','Presupuesto','Ejecutado','Disponible','% Ejec.'].map(h=><th key={h} style={S.th}>{h}</th>)}</tr></thead>
          <tbody>
            {allData.map((d,i)=>{
              const pres=d.presupuestos.reduce((s,p)=>s+(Number(p.monto)||0),0)
              const gastado=(d.gastosPresupuesto||[]).reduce((s,g)=>s+(Number(g.valorFactura)||0),0)
              const ejec=pres>0?(gastado/pres)*100:0
              return (
                <tr key={i}>
                  <td style={{...S.td,fontWeight:600}}><div style={{display:'flex',alignItems:'center',gap:8}}><div style={{width:10,height:10,borderRadius:'50%',background:d._color}}/>{d._unidad}</div></td>
                  <td style={{...S.td,fontFamily:'var(--mono)',color:'var(--text2)'}}>{cop(pres)}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)'}}>{cop(gastado)}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',color:pres-gastado>=0?'var(--green)':'var(--red)'}}>{cop(pres-gastado)}</td>
                  <td style={{...S.td,fontFamily:'var(--mono)',fontWeight:600,color:ejec>100?'var(--red)':ejec>80?'var(--yellow)':'var(--green)'}}>{pres>0?ejec.toFixed(1)+'%':'—'}</td>
                </tr>
              )
            })}
          </tbody>
        </table>
      </div>
    </div>
  )
}

// ═══════════════════════════════════════════════════════
// APP PRINCIPAL
// ═══════════════════════════════════════════════════════
export default function App() {
  const [usuario, setUsuario] = useState(getSessionUser)
  const [tab, setTab] = useState('dashboard')
  const [importResult, setImportResult] = useState(null)
  const [importando, setImportando] = useState(false)
  const [chatOpen, setChatOpen] = useState(false)
  const [sheetsLoading, setSheetsLoading] = useState(false)
  const [sheetsSync, setSheetsSync] = useState(null)
  const [sidebarOpen, setSidebarOpen] = useState(true)

  const loadUser = (uid) => {
    if(!uid) { _currentUserKey = STORAGE_KEY; return load() }
    _currentUserKey = getStorageKey(uid)
    try { const d=localStorage.getItem(_currentUserKey); return d?JSON.parse(d):load() } catch { return load() }
  }
  const [data, setData] = useState(()=>loadUser(usuario?.id))

  useEffect(()=>{
    if(!usuario) return
    setData(loadUser(usuario.id))
    if(SHEETS_CONFIG[usuario.id]?.enabled) {
      setSheetsLoading(true)
      setSheetsSync('syncing')
      const loadTimeout = setTimeout(() => setSheetsLoading(false), 5000)
      cargarDesdeSheets(usuario.id).then(sheetData => {
        clearTimeout(loadTimeout)
        setSheetsLoading(false)
        if(sheetData) {
          // Salvaguarda: si Sheets devuelve presupuestos vacío pero localmente
          // ya teníamos datos, conservar los locales (evita pérdida por
          // desajuste de nombres de columna en el backend de Sheets)
          const localActual = loadUser(usuario.id)
          if((!sheetData.presupuestos || sheetData.presupuestos.length===0) && localActual.presupuestos?.length>0) {
            sheetData.presupuestos = localActual.presupuestos
          }
          if((!sheetData.proyectos || sheetData.proyectos.length===0) && localActual.proyectos?.length>0) {
            sheetData.proyectos = localActual.proyectos
          }
          if((!sheetData.tareasAsignadas || sheetData.tareasAsignadas.length===0) && localActual.tareasAsignadas?.length>0) {
            sheetData.tareasAsignadas = localActual.tareasAsignadas
          }
          setData(sheetData)
          _currentUserKey = getStorageKey(usuario.id)
          localStorage.setItem(_currentUserKey, JSON.stringify(sheetData))
          setSheetsSync('ok')
        } else {
          setSheetsSync('error')
        }
      })
    }
  },[usuario?.id])

  const saveUser = (d, opts={}) => {
    if(!usuario) return
    _currentUserKey = getStorageKey(usuario.id)
    localStorage.setItem(_currentUserKey, JSON.stringify(d))
    if(SHEETS_CONFIG[usuario.id]?.enabled && opts.insertar) {
      setSheetsSync('syncing')
      insertarFilaSheets(usuario.id, opts.tipo, opts.insertar)
        .then(()=>setSheetsSync('ok'))
        .catch(()=>setSheetsSync('error'))
    }
    if(SHEETS_CONFIG[usuario.id]?.enabled && opts.eliminar) {
      setSheetsSync('syncing')
      eliminarFilaSheets(usuario.id, opts.tipo, opts.eliminar)
        .then(()=>setSheetsSync('ok'))
        .catch(()=>setSheetsSync('error'))
    }
  }

  const setDataUser = (d, opts={}) => { setData(d); saveUser(d, opts) }
  window._sheetsSyncFn = (d, opts) => saveUser(d, opts)

  const cerrarSesion = () => { localStorage.removeItem(AUTH_KEY); setUsuario(null); setTab('dashboard') }

  const handleFile = e => {
    const file=e.target.files[0]; if(!file) return
    setImportando(true); setImportResult(null)
    parsearExcel(file,data,setDataUser,result=>{setImportResult(result);setImportando(false)})
    e.target.value=''
  }
  const totalImp = importResult?Object.values(importResult.importados).reduce((s,n)=>s+n,0):0

  if(!usuario) return <LoginScreen onLogin={u=>{setUsuario(u);setTab(u.rol==='colaborador'?'tareasasignadas':'dashboard')}}/>

  if(sheetsLoading) return (
    <div style={{minHeight:'100vh',display:'flex',flexDirection:'column',alignItems:'center',justifyContent:'center',background:'linear-gradient(135deg,#1e1040 0%,#2d1a6e 100%)',gap:20}}>
      <div style={{width:56,height:56,borderRadius:14,background:'var(--accent)',display:'flex',alignItems:'center',justifyContent:'center',boxShadow:'0 8px 32px rgba(91,82,240,0.3)'}}>
        <BarChart2 size={28} color="#fff"/>
      </div>
      <div style={{textAlign:'center'}}>
        <div style={{fontWeight:800,fontSize:20,color:'#fff',marginBottom:6,letterSpacing:'-0.01em'}}>Prolub Trade Marketing</div>
        <div style={{fontSize:13,color:'rgba(255,255,255,0.6)'}}>Sincronizando datos de <strong style={{color:'#c084fc'}}>{usuario.nombre}</strong>...</div>
      </div>
      <div style={{width:220,height:5,background:'var(--bg3)',borderRadius:4,overflow:'hidden'}}>
        <div style={{height:'100%',background:'var(--accent)',borderRadius:4,animation:'loadbar 1.8s ease-in-out infinite'}}/>
      </div>
      <button onClick={()=>setSheetsLoading(false)} style={{fontSize:12,color:'rgba(255,255,255,0.5)',background:'rgba(255,255,255,0.08)',border:'1px solid rgba(255,255,255,0.15)',borderRadius:8,padding:'7px 18px',cursor:'pointer',fontFamily:'var(--font)',marginTop:4}}>
        Entrar sin sincronizar →
      </button>
      <style>{`@keyframes loadbar{0%{width:0%;margin-left:0}50%{width:60%;margin-left:20%}100%{width:0%;margin-left:100%}}`}</style>
    </div>
  )

  const esLider = usuario.rol==='lider'
  const esPresupuesto = usuario.rol==='presupuesto'
  const esColaborador = usuario.rol==='colaborador'

  const TABS_LIDER = [
    {id:'dashboard',label:'Dashboard Consolidado',icon:LayoutDashboard},
    {id:'asignarTareas',label:'Asignar tareas al equipo',icon:ListTodo},
    {id:'notificaciones',label:'Notificaciones',icon:Bell},
    {id:'tareasProyectos',label:'Tareas y Proyectos',icon:BookOpen}
  ]
  const TABS_PRES = [{id:'dashboard',label:'Presupuesto Consolidado',icon:DollarSign}]
  const TABS_COLAB = [
    {id:'tareasasignadas',label:'Mis Tareas',icon:ListTodo},
    {id:'tareasProyectos',label:'Tareas y Proyectos',icon:BookOpen},
  ]
  const tareasAsignadasCount = !esLider&&!esPresupuesto&&!esColaborador ? (data.tareasAsignadas||[]).filter(p=>p.estado!=='Listo'&&p.estado!=='Cancelado').length : 0
  const TABS_NORMAL = [
    {id:'dashboard',  label:'Dashboard',   icon:LayoutDashboard},
    {id:'inversiones',label:'Inversiones', icon:TrendingUp},
    {id:'ventas',     label:'Ventas',      icon:ShoppingCart},
    {id:'planes',     label:'Planes Q',    icon:BookOpen},
    {id:'presupuesto',label:'Presupuesto', icon:DollarSign},
    {id:'pendientes', label:'Pendientes',  icon:ListTodo},
    {id:'apoyocierre', label:'Apoyo Cierre', icon:DollarSign},
    {id:'tareasasignadas', label:'Tareas Asignadas', icon:ListTodo, badge:tareasAsignadasCount},
    {id:'tareasProyectos', label:'Tareas y Proyectos', icon:BookOpen},
  ]
  const tabs = esLider?TABS_LIDER:esPresupuesto?TABS_PRES:esColaborador?TABS_COLAB:TABS_NORMAL

  return (
    <div style={{minHeight:'100vh',display:'flex',flexDirection:'column',background:'var(--bg)'}}>
      {/* TOPBAR */}
      <header style={{background:'var(--sidebar-bg)',padding:'0 20px',display:'flex',alignItems:'center',height:56,gap:14,position:'sticky',top:0,zIndex:20,boxShadow:'0 2px 16px rgba(30,16,64,0.2)'}}>
        <button onClick={()=>setSidebarOpen(o=>!o)} style={{width:34,height:34,borderRadius:9,background:'rgba(255,255,255,0.1)',border:'1px solid rgba(255,255,255,0.15)',display:'flex',alignItems:'center',justifyContent:'center',cursor:'pointer',flexShrink:0}}>
          <BarChart2 size={15} color="rgba(255,255,255,0.8)"/>
        </button>
        <div style={{display:'flex',alignItems:'center',gap:9}}>
          <div style={{width:30,height:30,borderRadius:8,background:'linear-gradient(135deg,#6c3fc4,#a855f7)',display:'flex',alignItems:'center',justifyContent:'center',boxShadow:'0 2px 8px rgba(108,63,196,0.4)'}}><BarChart2 size={15} color="#fff"/></div>
          <span style={{fontWeight:800,fontSize:15,color:'#fff',letterSpacing:'-0.01em'}}>Prolub</span>
          <span style={{color:'rgba(255,255,255,0.4)',fontSize:12}}>/</span>
          <span style={{color:'rgba(255,255,255,0.6)',fontSize:12}}>Trade Marketing</span>
        </div>
        <div style={{marginLeft:'auto',display:'flex',alignItems:'center',gap:8}}>
          {sheetsSync==='syncing'&&(
            <div style={{display:'flex',alignItems:'center',gap:6,background:'rgba(217,119,6,0.1)',border:'1px solid rgba(217,119,6,0.3)',borderRadius:20,padding:'4px 12px'}}>
              <span style={{fontSize:11,color:'var(--orange)',fontWeight:600,animation:'spin 1s linear infinite',display:'inline-block'}}>↻</span>
              <span style={{fontSize:11,color:'var(--orange)',fontWeight:500}}>Sincronizando...</span>
            </div>
          )}
          {sheetsSync==='error'&&(
            <button onClick={()=>{setSheetsSync('syncing');cargarDesdeSheets(usuario.id).then(sd=>{if(sd){setData(sd);localStorage.setItem(getStorageKey(usuario.id),JSON.stringify(sd));setSheetsSync('ok')}else setSheetsSync('error')})}} style={{display:'flex',alignItems:'center',gap:6,background:'rgba(224,60,60,0.1)',border:'1px solid rgba(224,60,60,0.3)',borderRadius:20,padding:'4px 12px',cursor:'pointer',fontFamily:'var(--font)'}}>
              <span style={{fontSize:11,color:'var(--red)',fontWeight:600}}>✗ Sin conexión — Reintentar</span>
            </button>
          )}
          <div style={{display:'flex',alignItems:'center',gap:7,background:'rgba(255,255,255,0.12)',border:'1px solid rgba(255,255,255,0.18)',borderRadius:20,padding:'5px 14px'}}>
            <div style={{width:8,height:8,borderRadius:'50%',background:usuario.color,boxShadow:'0 0 6px '+usuario.color}}/>
            <span style={{fontSize:12,fontWeight:600,color:'#fff'}}>{usuario.nombre}</span>
            {sheetsSync==='ok'&&<span style={{fontSize:10,color:'#4ade80',fontWeight:600}}>✓</span>}
          </div>
          <style>{`@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}`}</style>
          <button onClick={cerrarSesion} style={{fontSize:11,color:'rgba(255,255,255,0.5)',background:'none',border:'none',cursor:'pointer',padding:'4px 8px',fontFamily:'var(--font)',letterSpacing:'0.02em'}}>Salir</button>
        </div>
      </header>

      <div style={{display:'flex',flex:1,overflow:'hidden'}}>
        {/* SIDEBAR OSCURO */}
        <aside style={{width:sidebarOpen?220:60,flexShrink:0,background:'var(--sidebar-bg)',display:'flex',flexDirection:'column',padding:'16px 0',gap:2,transition:'width 0.22s cubic-bezier(.4,0,.2,1)',overflow:'hidden',position:'sticky',top:52,height:'calc(100vh - 52px)',boxShadow:'4px 0 24px rgba(30,16,64,0.18)'}}>
          {tabs.map(t=>{ const Icon=t.icon; const active=tab===t.id; return (
            <button key={t.id} onClick={()=>setTab(t.id)} title={t.label} style={{display:'flex',alignItems:'center',gap:11,padding:sidebarOpen?'10px 18px':'10px 0',justifyContent:sidebarOpen?'flex-start':'center',margin:'0 8px',borderRadius:10,fontSize:13,fontWeight:active?600:400,background:active?'rgba(255,255,255,0.13)':'transparent',color:active?'#fff':'rgba(255,255,255,0.55)',border:'none',cursor:'pointer',fontFamily:'var(--font)',whiteSpace:'nowrap',position:'relative',transition:'all 0.15s',borderLeft:active?'3px solid #a855f7':'3px solid transparent'}}>
              <Icon size={16} style={{flexShrink:0,color:active?'#c084fc':'rgba(255,255,255,0.5)'}}/>
              {sidebarOpen&&<span style={{overflow:'hidden',textOverflow:'ellipsis'}}>{t.label}</span>}
              {t.badge>0&&<span style={{background:'#a855f7',color:'#fff',borderRadius:10,fontSize:9,fontWeight:700,padding:'1px 6px',marginLeft:'auto',flexShrink:0}}>{t.badge}</span>}
            </button>
          )})}
        </aside>

        {/* CONTENIDO */}
        <main style={{flex:1,padding:'24px 28px',overflowY:'auto',minWidth:0}}>
        <div key={tab+usuario.id}>
          {esLider && tab==='dashboard' && <DashboardLider/>}
          {esLider && tab==='asignarTareas' && <AsignarTareasEquipoPaola/>}
          {esLider && tab==='notificaciones' && <NotificacionesEquipoPaola/>}
          {esLider && tab==='tareasProyectos' && <TareasYProyectos data={data} setData={setDataUser} usuario={usuario}/>}
          {esPresupuesto && <PresupuestoConsolidado/>}
          {esColaborador && (
            <>
              {tab==='tareasasignadas' && <TareasAsignadas data={data} setData={setDataUser} usuario={usuario}/>}
              {tab==='tareasProyectos' && <TareasYProyectos data={data} setData={setDataUser} usuario={usuario}/>}
            </>
          )}
          {!esLider&&!esPresupuesto&&!esColaborador&&(
            <>
              {(()=>{
                const tareasLider = (data.tareasAsignadas||[]).filter(p=>p.estado!=='Listo'&&p.estado!=='Cancelado')
                if(tareasLider.length===0) return null
                return (
                  <div style={{background:'rgba(91,82,240,0.08)',border:'1px solid rgba(91,82,240,0.25)',borderRadius:12,padding:'12px 18px',marginBottom:8,display:'flex',alignItems:'center',gap:12,cursor:'pointer'}} onClick={()=>setTab('tareasasignadas')}>
                    <div style={{width:32,height:32,borderRadius:8,background:'var(--accent)',display:'flex',alignItems:'center',justifyContent:'center',fontSize:16,flexShrink:0}}>📋</div>
                    <div style={{flex:1}}>
                      <div style={{fontWeight:600,fontSize:13,color:'var(--accent2)'}}>Tienes {tareasLider.length} actividad{tareasLider.length>1?'es':''} asignada{tareasLider.length>1?'s':''} por el Líder de Mercadeo</div>
                      <div style={{fontSize:11,color:'var(--text3)',marginTop:2}}>{tareasLider.map(t=>t.tarea).slice(0,2).join(' · ')}{tareasLider.length>2?' · ...':''} — Clic para ver</div>
                    </div>
                    <span style={{fontSize:20,color:'var(--accent2)'}}>→</span>
                  </div>
                )
              })()}
              {tab==='dashboard'        &&<Dashboard         data={data}/>}
              {tab==='inversiones'     &&<Inversiones       data={data} setData={setDataUser}/>}
              {tab==='ventas'          &&<Ventas            data={data} setData={setDataUser}/>}
              {tab==='planes'          &&<Planes            data={data} setData={setDataUser}/>}
              {tab==='presupuesto'     &&<Presupuesto       data={data} setData={setDataUser}/>}
              {tab==='pendientes'      &&<Pendientes        data={data} setData={setDataUser}/>}
              {tab==='apoyocierre'     &&<ApoyoCierre       data={data} setData={setDataUser} usuario={usuario}/>}
              {tab==='tareasasignadas' &&<TareasAsignadas   data={data} setData={setDataUser} usuario={usuario}/>}
              {tab==='tareasProyectos'  &&<TareasYProyectos   data={data} setData={setDataUser} usuario={usuario}/>}
            </>
          )}
        </div>
        </main>
      </div>

      {!esLider&&!esPresupuesto&&(
        <footer style={{padding:'12px 28px',borderTop:'1px solid var(--border)',display:'flex',justifyContent:'space-between',alignItems:'center',flexWrap:'wrap',gap:8}}>
          <span style={{fontSize:11,color:'var(--text3)'}}>Datos de <strong style={{color:usuario.color}}>{usuario.nombre}</strong> — guardados en tu navegador</span>
          <div style={{display:'flex',gap:10,alignItems:'center',flexWrap:'wrap'}}>
            <label style={{...S.btn('var(--accent-soft)','var(--accent2)'),fontSize:12,padding:'5px 14px',border:'1px solid rgba(108,99,255,0.25)',cursor:'pointer'}}>
              ↑ Importar Excel<input type="file" accept=".xlsx,.xls" onChange={handleFile} style={{display:'none'}}/>
            </label>
            <button onClick={()=>exportarExcel(data)} style={{...S.btn('var(--green-soft)','var(--green)'),fontSize:12,padding:'5px 14px',border:'1px solid rgba(61,214,140,0.2)'}}>
              <Download size={13}/> Exportar Excel
            </button>
            {SHEETS_CONFIG[usuario?.id]?.enabled&&(
              <button onClick={()=>{
                setSheetsSync('syncing')
                cargarDesdeSheets(usuario.id).then(sd=>{
                  if(sd){setData(sd);localStorage.setItem(getStorageKey(usuario.id),JSON.stringify(sd));setSheetsSync('ok')}
                  else setSheetsSync('error')
                })
              }} style={{...S.btn('rgba(6,182,212,0.15)','#06b6d4'),fontSize:12,padding:'5px 14px',border:'1px solid rgba(6,182,212,0.3)'}}>
                ↻ Recargar Sheets
              </button>
            )}
            {SHEETS_CONFIG[usuario?.id]?.enabled&&(
              <button onClick={async ()=>{
                setSheetsSync('syncing')
                const tipos=['inversiones','ventas','presupuestos','gastosPresupuesto','planes','apoyoCierre','redenciones','pendientes','tareasAsignadas','proyectos']
                for(const t of tipos){
                  guardarEnSheets(usuario.id,t,data[t]||[])
                  await new Promise(r=>setTimeout(r,600))
                }
                setTimeout(()=>setSheetsSync('ok'),2000)
              }} disabled={sheetsSync==='syncing'} style={{...S.btn('rgba(168,139,250,0.15)','#a78bfa'),fontSize:12,padding:'5px 14px',border:'1px solid rgba(168,139,250,0.3)',opacity:sheetsSync==='syncing'?0.6:1}}>
                {sheetsSync==='syncing'?'⏳ Subiendo...':'↑ Subir todo a Sheets'}
              </button>
            )}
            <button onClick={()=>{if(confirm('¿Borrar tus datos?')){localStorage.removeItem(getStorageKey(usuario.id));_currentUserKey=STORAGE_KEY;window.location.reload()}}} style={{fontSize:11,color:'var(--text3)',background:'none',border:'none',cursor:'pointer',padding:'4px 8px',fontFamily:'var(--font)'}}>Resetear</button>
          </div>
        </footer>
      )}

      {importando&&<div style={{position:'fixed',bottom:80,right:28,background:'var(--bg2)',border:'1px solid var(--border2)',borderRadius:12,padding:'14px 20px',zIndex:200,fontSize:13,color:'var(--accent2)'}}>⏳ Importando datos...</div>}
      {importResult&&(
        <div style={{position:'fixed',bottom:80,right:28,background:'var(--bg2)',border:'1px solid '+(importResult.errores?.length?'var(--red)':'var(--green)'),borderRadius:12,padding:'16px 20px',zIndex:200,maxWidth:320,boxShadow:'0 8px 32px rgba(0,0,0,0.4)'}}>
          <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:10}}>
            <span style={{fontWeight:600,fontSize:13,color:importResult.errores?.length?'var(--red)':'var(--green)'}}>{importResult.errores?.length?'⚠️ Error':'✅ '+totalImp+' registros importados'}</span>
            <button onClick={()=>setImportResult(null)} style={{background:'none',border:'none',color:'var(--text3)',cursor:'pointer'}}><X size={14}/></button>
          </div>
          {Object.entries(importResult.importados).map(([k,v])=><div key={k} style={{fontSize:12,color:'var(--text2)',marginBottom:3}}>✓ <strong>{k}</strong>: {v} filas</div>)}
          {importResult.errores?.map((e,i)=><div key={i} style={{fontSize:12,color:'var(--red)',marginTop:4}}>{e}</div>)}
        </div>
      )}

      {!esLider&&!esPresupuesto&&(
        <>
          <button onClick={()=>setChatOpen(o=>!o)} style={{position:'fixed',bottom:24,right:24,width:52,height:52,borderRadius:'50%',background:chatOpen?'var(--bg3)':'var(--accent)',color:'#fff',border:chatOpen?'1px solid var(--border2)':'none',cursor:'pointer',display:'flex',alignItems:'center',justifyContent:'center',boxShadow:'0 4px 20px rgba(108,99,255,0.4)',zIndex:250,transition:'all 0.2s'}}>
            {chatOpen?<X size={20}/>:<MessageCircle size={22}/>}
          </button>
          {chatOpen&&<Asistente data={data} setData={setDataUser} onClose={()=>setChatOpen(false)}/>}
        </>
      )}
    </div>
  )
}
