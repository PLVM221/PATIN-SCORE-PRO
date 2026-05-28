import { useEffect, useMemo, useRef, useState } from 'react'
import * as XLSX from 'xlsx'
import './App.css'

const STORAGE_KEY = 'patin-score-pro-v1'
const WHISTLE_SRC = '/sounds/silbato.mp3'

const temasSistema = [
  { id: 'verde', nombre: 'Verde clásico' },
  { id: 'azul', nombre: 'Azul nocturno' },
  { id: 'violeta', nombre: 'Violeta escenario' },
  { id: 'ambar', nombre: 'Ámbar deportivo' },
  { id: 'rojo', nombre: 'Rojo gala' },
  { id: 'hielo', nombre: 'Hielo' },
  { id: 'dorado', nombre: 'Dorado copa' },
]

const demo = {
  torneos: [{
    id: 't1',
    nombre: 'Copa Ciudad 2026',
    liga: 'Liga Metropolitana',
    ligaLogo: '',
    clubOrganizadorId: 'cl1',
    sede: 'Polideportivo Central',
    fechaDesde: '2026-05-26',
    fechaHasta: '2026-05-27',
    horarioInicio: '09:00',
    turnosPorDia: 2,
    turnos: ['Mañana', 'Tarde'],
  }],
  categorias: [
    { id: 'c1', nombre: 'Libre Infantil B', torneoId: 't1', dia: '2026-05-26', turno: 'Mañana', orden: 1 },
    { id: 'c2', nombre: 'Escuela Cadete C', torneoId: 't1', dia: '2026-05-26', turno: 'Tarde', orden: 2 },
    { id: 'c3', nombre: 'Show Juvenil', torneoId: 't1', dia: '2026-05-27', turno: 'Mañana', orden: 3 },
  ],
  clubes: [
    { id: 'cl1', nombre: 'Sol Patín', color: '#22c55e', logo: '' },
    { id: 'cl2', nombre: 'Estrella Norte', color: '#38bdf8', logo: '' },
    { id: 'cl3', nombre: 'Ruedas del Sur', color: '#f59e0b', logo: '' },
  ],
  tecnicas: [
    { id: 'te1', nombre: 'Valeria Costa', clubId: 'cl1', foto: '' },
    { id: 'te2', nombre: 'Marina Quiroga', clubId: 'cl2', foto: '' },
    { id: 'te3', nombre: 'Lucia Ferreyra', clubId: 'cl3', foto: '' },
  ],
  jueces: [
    { id: 'j1', nombre: 'Juez A', rol: 'Técnico' },
    { id: 'j2', nombre: 'Juez B', rol: 'Artístico' },
    { id: 'j3', nombre: 'Juez C', rol: 'General' },
  ],
  conceptosPuntaje: [
    { id: 'elemento-tecnico', nombre: 'Elemento técnico' },
    { id: 'componente', nombre: 'Componente' },
  ],
  canciones: [
    { id: 'libertango', nombre: 'Libertango', archivo: '' },
    { id: 'cinema-paradiso', nombre: 'Cinema Paradiso', archivo: '' },
    { id: 'experience', nombre: 'Experience', archivo: '' },
    { id: 'heroes', nombre: 'Heroes', archivo: '' },
  ],
  buffet: [
    { id: 'bf1', producto: 'Agua', precio: 1200 },
    { id: 'bf2', producto: 'Café', precio: 1000 },
    { id: 'bf3', producto: 'Sándwich', precio: 2500 },
  ],
  usuarios: [
    { id: 'admin', nombre: 'Admin', tabs: tabsBase() },
    { id: 'operador', nombre: 'Operador', tabs: ['Operador', 'Pública LED', 'Web pública', 'Reportes', 'Tanteador'] },
    { id: 'jueces', nombre: 'Jueces', tabs: ['Jueces'] },
  ],
  patinadoras: [
    { id: 'p1', orden: 1, ordenSalida: 1, dia: '2026-05-26', turno: 'Mañana', estado: 'pendiente', nombre: 'Sofia Benitez', edad: 10, clubId: 'cl1', tecnicaId: 'te1', categoriaId: 'c1', musicaId: 'libertango', musica: 'Libertango', audio: '', foto: '' },
    { id: 'p2', orden: 2, ordenSalida: 2, dia: '2026-05-26', turno: 'Mañana', estado: 'pendiente', nombre: 'Martina Lagos', edad: 11, clubId: 'cl2', tecnicaId: 'te2', categoriaId: 'c1', musicaId: 'cinema-paradiso', musica: 'Cinema Paradiso', audio: '', foto: '' },
    { id: 'p3', orden: 1, ordenSalida: 1, dia: '2026-05-26', turno: 'Tarde', estado: 'pendiente', nombre: 'Camila Rios', edad: 14, clubId: 'cl3', tecnicaId: 'te3', categoriaId: 'c2', musicaId: 'experience', musica: 'Experience', audio: '', foto: '' },
    { id: 'p4', orden: 1, ordenSalida: 1, dia: '2026-05-27', turno: 'Mañana', estado: 'pendiente', nombre: 'Equipo Aurora', edad: 15, clubId: 'cl1', tecnicaId: 'te1', categoriaId: 'c3', musicaId: 'heroes', musica: 'Heroes', audio: '', foto: '' },
  ],
  puntajes: {},
  logs: [{ id: 'l1', fecha: new Date().toISOString(), texto: 'Demo iniciado' }],
  tema: 'verde',
}

const basePista = {
  torneoId: 't1',
  categoriaId: 'c1',
  patinadoraId: 'p1',
  estado: 'Preparando',
  modo: 'jueces',
  recesoMin: 5,
  demoraRecesoSeg: 60,
  juecesPuntajeSeg: 5,
  inicioProgramadoHasta: null,
  recesoHasta: null,
  pruebaPistaMin: 5,
  pruebaPistaHasta: null,
  puntajeHasta: null,
  puntajePatinadoraId: null,
  juezCambioHasta: null,
  juezSiguientePatinadoraId: null,
  categoriaFinalHasta: null,
  autoRecesoHasta: null,
}

const tabs = tabsBase()

function tabsBase() {
  return ['Operador', 'Jueces', 'Pública LED', 'Datos', 'Web pública', 'Reportes', 'Tanteador', 'Ranking clubes', 'Actas', 'Registros']
}

function loadData() {
  try {
    const saved = localStorage.getItem(STORAGE_KEY)
    return normalizeData(saved ? { ...demo, ...JSON.parse(saved) } : demo)
  } catch {
    return normalizeData(demo)
  }
}

function normalizeData(data) {
  return {
    ...data,
    torneos: (data.torneos || demo.torneos).map((torneo) => ({
      liga: '',
      ligaLogo: '',
      clubOrganizadorId: '',
      activo: true,
      confirmado: false,
      fechaDesde: torneo.fecha || new Date().toISOString().slice(0, 10),
      fechaHasta: torneo.fecha || new Date().toISOString().slice(0, 10),
      horarioInicio: '',
      turnosPorDia: 1,
      turnos: ['Único'],
      ...torneo,
    })),
    buffet: data.buffet || demo.buffet,
    usuarios: data.usuarios || demo.usuarios,
    tema: data.tema || 'verde',
    categorias: (data.categorias || demo.categorias).map((item, index) => ({
      dia: data.torneos?.[0]?.fechaDesde || data.torneos?.[0]?.fecha || '',
      turno: data.torneos?.[0]?.turnos?.[0] || 'Único',
      orden: index + 1,
      ...item,
    })),
    canciones: data.canciones || [],
    conceptosPuntaje: data.conceptosPuntaje || demo.conceptosPuntaje,
    estadoPista: data.estadoPista || basePista,
    patinadoras: data.patinadoras.map((item, index) => ({
      orden: index + 1,
      ordenSalida: item.ordenSalida || item.orden || index + 1,
      dia: item.dia || data.torneos?.[0]?.fechaDesde || data.torneos?.[0]?.fecha || '',
      turno: item.turno || data.torneos?.[0]?.turnos?.[0] || 'Único',
      estado: 'pendiente',
      musicaId: slug(item.musica || item.nombre),
      ...item,
    })),
  }
}

function id(prefix) {
  return `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2)}`
}

function slug(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/(^-|-$)/g, '') || id('item')
}

function sumaConceptos(conceptos) {
  return Object.values(conceptos || {}).reduce((sum, valor) => {
    const number = Number(valor)
    return Number.isFinite(number) ? sum + number : sum
  }, 0)
}

function totalPuntaje(puntaje, modo, conceptosConfig = []) {
  if (!puntaje) return null
  const bonus = Number(puntaje.bonus || 0)
  if (modo === 'consenso') {
    if (puntaje.consensoConceptos) return sumaConceptos(puntaje.consensoConceptos) + bonus
    return puntaje.consenso ? Number(puntaje.consenso) + bonus : null
  }
  const tieneConceptos = Object.values(puntaje.jueces || {}).some((item) => item.conceptos)
  if (tieneConceptos) {
    const conceptos = conceptosConfig.length
      ? conceptosConfig
      : Array.from(new Set(Object.values(puntaje.jueces || {}).flatMap((item) => Object.keys(item.conceptos || {})))).map((conceptoId) => ({ id: conceptoId }))
    let total = 0
    let cargados = 0
    conceptos.forEach((concepto) => {
      const valores = Object.values(puntaje.jueces || {})
        .map((item) => Number(item.conceptos?.[concepto.id]))
        .filter((valor) => Number.isFinite(valor))
      if (valores.length) {
        total += valores.reduce((sum, valor) => sum + valor, 0) / valores.length
        cargados += 1
      }
    })
    return cargados ? total + bonus : null
  }
  const valores = Object.values(puntaje.jueces || {})
    .map((item) => Number(item.valor))
    .filter((valor) => Number.isFinite(valor) && valor > 0)
  if (!valores.length) return bonus || null
  return valores.reduce((sum, valor) => sum + valor, 0) / valores.length + bonus
}

function parcialesPuntaje(puntaje, modo, conceptosConfig = []) {
  if (!puntaje) return []
  if (modo === 'consenso') {
    if (puntaje.consensoConceptos) {
      return conceptosConfig.map((concepto) => ({
        id: concepto.id,
        nombre: concepto.nombre,
        valor: puntaje.consensoConceptos?.[concepto.id] === '' ? null : Number(puntaje.consensoConceptos?.[concepto.id]),
      }))
    }
    return [{ id: 'consenso', nombre: 'Puntaje', valor: puntaje.consenso === '' ? null : Number(puntaje.consenso) }]
  }
  const tieneConceptos = Object.values(puntaje.jueces || {}).some((item) => item.conceptos)
  if (!tieneConceptos) {
    return Object.entries(puntaje.jueces || {}).map(([juezId, item]) => ({
      id: juezId,
      nombre: juezId.toUpperCase(),
      valor: item.valor === '' ? null : Number(item.valor),
    }))
  }
  return conceptosConfig.map((concepto) => {
    const valores = Object.values(puntaje.jueces || {})
      .map((item) => Number(item.conceptos?.[concepto.id]))
      .filter((valor) => Number.isFinite(valor))
    return {
      id: concepto.id,
      nombre: concepto.nombre,
      valor: valores.length ? valores.reduce((sum, valor) => sum + valor, 0) / valores.length : null,
    }
  })
}

function rankingCategoria(data, categoriaId, modo) {
  return data.patinadoras
    .filter((patinadora) => patinadora.categoriaId === categoriaId && patinadora.estado !== 'ausente')
    .map((patinadora) => ({
      ...patinadora,
      total: totalPuntaje(data.puntajes[patinadora.id], modo, data.conceptosPuntaje),
      confirmado: Boolean(data.puntajes[patinadora.id]?.confirmado),
    }))
    .sort((a, b) => (b.total ?? -1) - (a.total ?? -1) || compararSalida(a, b))
    .map((row, index) => ({ ...row, puesto: row.total == null ? null : index + 1 }))
}

function rankingClubes(data, modo) {
  const tabla = new Map()
  data.categorias.forEach((categoria) => {
    rankingCategoria(data, categoria.id, modo)
      .filter((row) => row.puesto)
      .forEach((row) => {
        const actual = tabla.get(row.clubId) || { clubId: row.clubId, primero: 0, segundo: 0, tercero: 0, cuarto: 0, quinto: 0, puntos: 0, incentivos: 0 }
        if (row.puesto === 1) actual.primero += 1
        if (row.puesto === 2) actual.segundo += 1
        if (row.puesto === 3) actual.tercero += 1
        if (row.puesto === 4) actual.cuarto += 1
        if (row.puesto === 5) actual.quinto += 1
        if (row.puesto <= 5) actual.puntos += 6 - row.puesto
        else actual.incentivos += 1
        tabla.set(row.clubId, actual)
      })
  })
  return [...tabla.values()].sort((a, b) => b.puntos - a.puntos || b.primero - a.primero || b.segundo - a.segundo || b.tercero - a.tercero || b.cuarto - a.cuarto || b.quinto - a.quinto || b.incentivos - a.incentivos)
}

function cell(row, names) {
  const entries = Object.entries(row)
  for (const name of names) {
    const found = entries.find(([key]) => slug(key) === slug(name))
    if (found) return String(found[1] || '').trim()
  }
  return ''
}

function compararSalida(a, b) {
  return String(a.dia || '').localeCompare(String(b.dia || ''))
    || String(a.turno || '').localeCompare(String(b.turno || ''))
    || (Number(a.ordenSalida || a.orden || 0) - Number(b.ordenSalida || b.orden || 0))
    || String(a.nombre || '').localeCompare(String(b.nombre || ''))
}

function compararCategoria(a, b) {
  return String(a.dia || '').localeCompare(String(b.dia || ''))
    || String(a.turno || '').localeCompare(String(b.turno || ''))
    || (Number(a.orden || 0) - Number(b.orden || 0))
    || String(a.nombre || '').localeCompare(String(b.nombre || ''))
}

function compararReporte(data, a, b) {
  const categoriaA = data.categorias.find((categoria) => categoria.id === a.categoriaId) || {}
  const categoriaB = data.categorias.find((categoria) => categoria.id === b.categoriaId) || {}
  return String(a.dia || categoriaA.dia || '').localeCompare(String(b.dia || categoriaB.dia || ''))
    || String(a.turno || categoriaA.turno || '').localeCompare(String(b.turno || categoriaB.turno || ''))
    || String(categoriaA.nombre || '').localeCompare(String(categoriaB.nombre || ''))
    || (Number(a.ordenSalida || a.orden || 0) - Number(b.ordenSalida || b.orden || 0))
    || String(a.nombre || '').localeCompare(String(b.nombre || ''))
}

function diasTorneo(torneo) {
  const desde = torneo?.fechaDesde || torneo?.fecha
  const hasta = torneo?.fechaHasta || desde
  if (!desde) return []
  const start = new Date(`${desde}T00:00:00`)
  const end = new Date(`${hasta}T00:00:00`)
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) return [desde]
  const days = []
  for (let current = new Date(start); current <= end; current.setDate(current.getDate() + 1)) {
    days.push(current.toISOString().slice(0, 10))
  }
  return days
}

function nombreClub(data, clubId) {
  return data.clubes.find((club) => club.id === clubId)?.nombre || ''
}

function nombreCategoria(data, categoriaId) {
  return data.categorias.find((categoria) => categoria.id === categoriaId)?.nombre || ''
}

function siguienteCategoria(data, categoriaId) {
  const categorias = [...data.categorias].sort(compararCategoria)
  const index = categorias.findIndex((categoria) => categoria.id === categoriaId)
  return categorias[index + 1]
}

function readSheetRows(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => {
      try {
        const workbook = XLSX.read(reader.result, { type: 'array' })
        const sheet = workbook.Sheets[workbook.SheetNames[0]]
        resolve(XLSX.utils.sheet_to_json(sheet, { defval: '' }))
      } catch (error) {
        reject(error)
      }
    }
    reader.onerror = reject
    reader.readAsArrayBuffer(file)
  })
}

export default function App() {
  const [data, setData] = useState(loadData)
  const [pista, setPista] = useState(basePista)
  const [tab, setTab] = useState('Operador')
  const [juezId, setJuezId] = useState('j1')
  const [usuarioId, setUsuarioId] = useState('admin')
  const [persistencia, setPersistencia] = useState('localStorage')
  const audioRef = useRef(null)
  const sqlActivoRef = useRef(false)
  const cargandoSqlRef = useRef(true)
  const omitirGuardadoRef = useRef(false)
  const esVistaWeb = new URLSearchParams(window.location.search).get('vista') === 'web'

  useEffect(() => {
    let active = true
    async function cargarSql() {
      try {
        const response = await fetch('/api/state')
        if (!response.ok) throw new Error('SQL no disponible')
        const payload = await response.json()
        sqlActivoRef.current = true
        setPersistencia('SQL')
        if (payload.data) {
          omitirGuardadoRef.current = true
          const next = normalizeData({ ...demo, ...payload.data })
          setData(next)
          if (payload.data.estadoPista) setPista({ ...basePista, ...payload.data.estadoPista })
        } else {
          await fetch('/api/state', {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data),
          })
        }
      } catch {
        sqlActivoRef.current = false
        setPersistencia('localStorage')
      } finally {
        if (active) cargandoSqlRef.current = false
      }
    }
    cargarSql()
    return () => {
      active = false
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  useEffect(() => {
    if (esVistaWeb) return undefined
    // Sincroniza el estado operativo con SQL para que el QR vea receso/prueba en vivo.
    // eslint-disable-next-line react-hooks/set-state-in-effect
    setData((prev) => {
      const actualEstado = JSON.stringify(prev.estadoPista || {})
      const nuevoEstado = JSON.stringify(pista)
      if (actualEstado === nuevoEstado) return prev
      return { ...prev, estadoPista: pista }
    })
    return undefined
  }, [pista, esVistaWeb])

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data))
    if (omitirGuardadoRef.current) {
      omitirGuardadoRef.current = false
      return undefined
    }
    if (!sqlActivoRef.current || cargandoSqlRef.current) return undefined
    const timer = window.setTimeout(() => {
      fetch('/api/state', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
      }).catch(() => setPersistencia('localStorage'))
    }, 350)
    return () => window.clearTimeout(timer)
  }, [data])

  useEffect(() => {
    const onStorage = (event) => {
      if (event.key === STORAGE_KEY) setData(loadData())
    }
    window.addEventListener('storage', onStorage)
    return () => window.removeEventListener('storage', onStorage)
  }, [])

  useEffect(() => {
    if (!esVistaWeb) return undefined
    const timer = window.setInterval(() => {
      fetch('/api/state')
        .then((response) => response.ok ? response.json() : null)
        .then((payload) => {
          if (payload?.data) {
            omitirGuardadoRef.current = true
            const next = normalizeData({ ...demo, ...payload.data })
            setData(next)
          }
        })
        .catch(() => {})
    }, 2500)
    return () => window.clearInterval(timer)
  }, [esVistaWeb])

  const actual = useMemo(() => {
    const torneo = data.torneos.find((item) => item.id === pista.torneoId)
    const categoria = data.categorias.find((item) => item.id === pista.categoriaId)
    const patinadora = data.patinadoras.find((item) => item.id === pista.patinadoraId)
    const club = data.clubes.find((item) => item.id === patinadora?.clubId)
    const tecnica = data.tecnicas.find((item) => item.id === patinadora?.tecnicaId)
    const cancion = data.canciones.find((item) => item.id === patinadora?.musicaId)
    const puntaje = data.puntajes[patinadora?.id]
    return { torneo, categoria, patinadora, club, tecnica, cancion, puntaje }
  }, [data, pista])

  const patinadorasCategoria = data.patinadoras
    .filter((item) => item.categoriaId === pista.categoriaId && item.estado !== 'ausente')
    .sort(compararSalida)
  const totalActual = totalPuntaje(actual.puntaje, pista.modo, data.conceptosPuntaje)
  const clubOrganizador = data.clubes.find((club) => club.id === actual.torneo?.clubOrganizadorId)
  const usuarioActual = data.usuarios.find((item) => item.id === usuarioId) || data.usuarios[0]
  const esAdmin = usuarioActual?.id === 'admin'
  const tabsVisibles = tabs.filter((item) => usuarioActual?.tabs?.includes(item))

  const pistaPublica = data.estadoPista || pista
  const inicioProgramadoHasta = actual.torneo?.inicioProgramadoHasta || pista.inicioProgramadoHasta

  useEffect(() => {
    if (esVistaWeb || actual.torneo?.inicioTorneoEn || !inicioProgramadoHasta) return undefined
    const espera = Math.max(0, inicioProgramadoHasta - Date.now())
    const torneoId = actual.torneo?.id
    const torneoNombre = actual.torneo?.nombre || 'Torneo'
    const timer = window.setTimeout(() => {
      setData((prev) => {
        const next = structuredClone(prev)
        const torneo = next.torneos.find((item) => item.id === torneoId) || next.torneos[0]
        if (!torneo || torneo.inicioTorneoEn) return prev
        torneo.inicioTorneoEn = new Date().toISOString()
        torneo.inicioProgramadoHasta = null
        next.logs.unshift({ id: id('log'), fecha: new Date().toISOString(), texto: `Inicio de torneo: ${torneoNombre}` })
        return next
      })
      setPista((prev) => ({ ...prev, estado: 'Torneo iniciado', inicioProgramadoHasta: null }))
    }, espera)
    return () => window.clearTimeout(timer)
  }, [actual.torneo?.id, actual.torneo?.inicioTorneoEn, actual.torneo?.nombre, esVistaWeb, inicioProgramadoHasta])

  useEffect(() => {
    if (esVistaWeb || !pista.juezCambioHasta) return undefined
    const avanzar = () => {
      setPista((prev) => {
        if (!prev.juezCambioHasta || Date.now() < prev.juezCambioHasta) return prev
        const siguienteId = prev.juezSiguientePatinadoraId
        return {
          ...prev,
          patinadoraId: siguienteId || prev.patinadoraId,
          estado: siguienteId ? 'Preparando' : 'Categoría finalizada',
          juezCambioHasta: null,
          juezSiguientePatinadoraId: null,
        }
      })
    }
    const espera = Math.max(0, pista.juezCambioHasta - Date.now())
    const timer = window.setTimeout(avanzar, espera)
    return () => window.clearTimeout(timer)
  }, [esVistaWeb, pista.juezCambioHasta])

  if (esVistaWeb) {
    return <VistaWeb data={data} modo={pistaPublica.modo || pista.modo} pista={pistaPublica} />
  }

  function log(texto) {
    setData((prev) => ({ ...prev, logs: [{ id: id('log'), fecha: new Date().toISOString(), torneoId: pista.torneoId, texto }, ...prev.logs] }))
  }

  function mutate(fn, texto) {
    setData((prev) => {
      const next = structuredClone(prev)
      fn(next)
      if (texto) next.logs.unshift({ id: id('log'), fecha: new Date().toISOString(), torneoId: pista.torneoId, texto })
      return next
    })
  }

  function cambiarCategoria(categoriaId) {
    const primera = data.patinadoras
      .filter((item) => item.categoriaId === categoriaId && item.estado !== 'ausente')
      .sort(compararSalida)[0]
    setPista((prev) => ({ ...prev, categoriaId, patinadoraId: primera?.id || '', estado: 'Preparando', juezCambioHasta: null, juezSiguientePatinadoraId: null }))
  }

  function mover(delta) {
    const index = patinadorasCategoria.findIndex((item) => item.id === pista.patinadoraId)
    const next = patinadorasCategoria[index + delta]
    if (next) setPista((prev) => ({ ...prev, patinadoraId: next.id, estado: 'Preparando', juezCambioHasta: null, juezSiguientePatinadoraId: null }))
  }

  function guardarArchivo(tipo, entidad, itemId, campo, file) {
    if (!file) return
    const reader = new FileReader()
    reader.onload = () => {
      mutate((draft) => {
        const row = draft[entidad].find((item) => item.id === itemId)
        row[campo] = reader.result
        if (tipo === 'audio') {
          row.musica = file.name
          row.musicaId = slug(file.name.replace(/\.[^.]+$/, ''))
        }
      }, tipo === 'audio' ? 'Música cargada' : 'Imagen cargada')
    }
    reader.readAsDataURL(file)
  }

  function silbato() {
    reproducirSilbato()
    log('Silbato')
  }

  function actualizarPuntaje(patinadoraId, fn, texto) {
    mutate((draft) => {
      draft.puntajes[patinadoraId] ||= { jueces: {}, consenso: '', consensoConceptos: {}, bonus: 0, confirmado: null }
      fn(draft.puntajes[patinadoraId])
    }, texto)
  }

  function publicarPuntaje() {
    if (!actual.patinadora) return
    const total = totalPuntaje(actual.puntaje, pista.modo, data.conceptosPuntaje)
    if (total == null) return
    mutate((draft) => {
      draft.puntajes[actual.patinadora.id] ||= { jueces: {}, consenso: '', consensoConceptos: {}, bonus: 0, confirmado: null }
      draft.puntajes[actual.patinadora.id].confirmado = new Date().toISOString()
      const row = draft.patinadoras.find((item) => item.id === actual.patinadora.id)
      row.estado = 'finalizada'
    }, `Puntaje publicado: ${actual.patinadora.nombre}`)

    const index = patinadorasCategoria.findIndex((item) => item.id === actual.patinadora.id)
    const siguiente = patinadorasCategoria[index + 1]
    // eslint-disable-next-line react-hooks/purity
    const now = Date.now()
    const demoraRecesoMs = Math.max(0, Number(pista.demoraRecesoSeg) || 0) * 1000
    const demoraJuecesMs = Math.max(0, Number(pista.juecesPuntajeSeg) || 0) * 1000
    const categoriaFinalHasta = siguiente ? null : now + demoraRecesoMs
    setPista((prev) => ({
      ...prev,
      patinadoraId: actual.patinadora.id,
      estado: 'Puntaje publicado',
      puntajeHasta: now + 10000,
      puntajePatinadoraId: actual.patinadora.id,
      demoraRecesoSeg: prev.demoraRecesoSeg,
      juecesPuntajeSeg: prev.juecesPuntajeSeg,
      juezCambioHasta: now + demoraJuecesMs,
      juezSiguientePatinadoraId: siguiente?.id || null,
      categoriaFinalHasta,
      autoRecesoHasta: categoriaFinalHasta ? categoriaFinalHasta + Number(prev.recesoMin) * 60000 : null,
    }))
  }

  function pasarSiguienteCategoria() {
    const categoriasOrdenadas = [...data.categorias].sort(compararCategoria)
    const actualIndex = categoriasOrdenadas.findIndex((item) => item.id === pista.categoriaId)
    const siguienteCategoria = categoriasOrdenadas[actualIndex + 1]
    if (!siguienteCategoria) return
    const primera = data.patinadoras
      .filter((item) => item.categoriaId === siguienteCategoria.id && item.estado !== 'ausente')
      .sort(compararSalida)[0]
    setPista((prev) => ({
      ...prev,
      categoriaId: siguienteCategoria.id,
      patinadoraId: primera?.id || '',
      estado: 'Preparando',
      puntajeHasta: null,
      puntajePatinadoraId: null,
      juezCambioHasta: null,
      juezSiguientePatinadoraId: null,
      categoriaFinalHasta: null,
      autoRecesoHasta: null,
    }))
  }

  function iniciarPruebaPista(minutos) {
    const duration = Math.max(1, Number(minutos) || 1)
    setPista((prev) => ({
      ...prev,
      pruebaPistaMin: duration,
      estado: 'Probando pista',
      pruebaPistaHasta: Date.now() + duration * 60000,
      recesoHasta: null,
      categoriaFinalHasta: null,
      autoRecesoHasta: null,
    }))
    log(`Prueba de pista ${duration} min`)
  }

  function iniciarTorneo() {
    if (data.torneos[0]?.inicioTorneoEn) return
    mutate((draft) => {
      draft.torneos[0].inicioTorneoEn = new Date().toISOString()
      draft.torneos[0].inicioProgramadoHasta = null
    }, `Inicio de torneo: ${data.torneos[0]?.nombre || 'Torneo'}`)
    setPista((prev) => ({ ...prev, estado: 'Torneo iniciado', inicioProgramadoHasta: null }))
  }

  function programarInicioTorneo(minutos) {
    if (data.torneos[0]?.inicioTorneoEn) return
    const duration = Math.max(1, Number(minutos) || 1)
    const hasta = Date.now() + duration * 60000
    mutate((draft) => {
      draft.torneos[0].inicioProgramadoHasta = hasta
    }, `Inicio de torneo programado en ${duration} min`)
    setPista((prev) => ({ ...prev, estado: 'Inicio programado', inicioProgramadoHasta: hasta }))
  }

  function rehabilitarInicioTorneo() {
    mutate((draft) => {
      draft.torneos[0].inicioTorneoEn = null
      draft.torneos[0].inicioProgramadoHasta = null
    }, `Inicio de torneo rehabilitado: ${data.torneos[0]?.nombre || 'Torneo'}`)
    setPista((prev) => ({ ...prev, inicioProgramadoHasta: null, estado: 'Preparando' }))
  }

  function marcarAusente() {
    if (!actual.patinadora) return
    mutate((draft) => {
      const row = draft.patinadoras.find((item) => item.id === actual.patinadora.id)
      row.estado = 'ausente'
      delete draft.puntajes[actual.patinadora.id]
    }, `No se presentó: ${actual.patinadora.nombre}`)
    const siguiente = patinadorasCategoria.find((item) => item.id !== actual.patinadora.id)
    setPista((prev) => ({ ...prev, patinadoraId: siguiente?.id || '', estado: 'Preparando' }))
  }

  function postergar() {
    if (!actual.patinadora) return
    const maxOrden = Math.max(...data.patinadoras.filter((item) => item.categoriaId === pista.categoriaId).map((item) => item.ordenSalida || item.orden || 0), 0)
    mutate((draft) => {
      const row = draft.patinadoras.find((item) => item.id === actual.patinadora.id)
      row.orden = maxOrden + 1
      row.ordenSalida = maxOrden + 1
      row.estado = 'postergada'
    }, `Postergada al final: ${actual.patinadora.nombre}`)
    const siguiente = patinadorasCategoria.find((item) => item.id !== actual.patinadora.id)
    setPista((prev) => ({ ...prev, patinadoraId: siguiente?.id || actual.patinadora.id, estado: 'Preparando' }))
  }

  function playActual() {
    audioRef.current?.play()
    log(`Reproducir música: ${actual.patinadora?.nombre}`)
  }

  async function importarListado(file) {
    if (!file) return
    const rows = await readSheetRows(file)
    let primeraCategoriaId = ''
    let primeraPatinadoraId = ''
    setData((prev) => {
      const draft = structuredClone(prev)
      const clubes = [...draft.clubes]
      const categorias = [...draft.categorias]
      const patinadoras = rows
        .map((row, index) => {
          const nombre = cell(row, ['Nombre', 'Patinadora', 'Competidor'])
          if (!nombre) return null
          const clubNombre = cell(row, ['Club']) || 'Sin club'
          const categoriaNombre = cell(row, ['Categoria', 'Categoría']) || 'Sin categoría'
          const musicaTexto = cell(row, ['MusicaId', 'Musica ID', 'MúsicaId', 'Música ID', 'Musica', 'Música']) || nombre
          const dia = cell(row, ['Dia', 'Día', 'Fecha']) || draft.torneos.find((torneo) => torneo.id === pista.torneoId)?.fechaDesde || ''
          const turno = cell(row, ['Turno']) || draft.torneos.find((torneo) => torneo.id === pista.torneoId)?.turnos?.[0] || 'Único'
          const ordenSalida = Number(cell(row, ['Orden', 'Orden salida', 'Orden de salida', 'Ubicacion', 'Ubicación', 'Ubicacion de salida', 'Ubicación de salida'])) || index + 1
          let club = clubes.find((item) => slug(item.nombre) === slug(clubNombre))
          if (!club) {
            club = { id: id('club'), nombre: clubNombre, color: '#28e67a', logo: '' }
            clubes.push(club)
          }
          let categoria = categorias.find((item) => slug(item.nombre) === slug(categoriaNombre))
          if (!categoria) {
            categoria = { id: id('cat'), nombre: categoriaNombre, torneoId: pista.torneoId }
            categorias.push(categoria)
          }
          const musicaId = slug(musicaTexto)
          if (!draft.canciones.find((item) => item.id === musicaId)) {
            draft.canciones.push({ id: musicaId, nombre: musicaTexto, archivo: '' })
          }
          return {
            id: id('pat'),
            orden: index + 1,
            ordenSalida,
            dia,
            turno,
            estado: 'pendiente',
            nombre,
            edad: Number(cell(row, ['Edad'])) || '',
            clubId: club.id,
            tecnicaId: '',
            categoriaId: categoria.id,
            musicaId,
            musica: musicaTexto,
            audio: '',
            foto: '',
          }
        })
        .filter(Boolean)
      draft.clubes = clubes
      draft.categorias = categorias
      draft.patinadoras = patinadoras
      draft.puntajes = {}
      draft.logs.unshift({ id: id('log'), fecha: new Date().toISOString(), texto: `Listado importado: ${file.name}` })
      primeraCategoriaId = patinadoras[0]?.categoriaId || ''
      primeraPatinadoraId = patinadoras[0]?.id || ''
      return draft
    })
    if (primeraCategoriaId && primeraPatinadoraId) {
      setPista((prev) => ({ ...prev, categoriaId: primeraCategoriaId, patinadoraId: primeraPatinadoraId, estado: 'Preparando' }))
    }
  }

  function cargarCanciones(files) {
    Array.from(files || []).forEach((file) => {
      const reader = new FileReader()
      reader.onload = () => {
        const musicId = slug(file.name.replace(/\.[^.]+$/, ''))
        mutate((draft) => {
          const found = draft.canciones.find((item) => item.id === musicId)
          if (found) {
            found.archivo = reader.result
            found.nombre = found.nombre || file.name
          } else {
            draft.canciones.push({ id: musicId, nombre: file.name, archivo: reader.result })
          }
        }, `Canción cargada: ${file.name}`)
      }
      reader.readAsDataURL(file)
    })
  }

  function cargarEscudosClubes(files) {
    Array.from(files || []).forEach((file) => {
      const reader = new FileReader()
      reader.onload = () => {
        const nombreClubArchivo = file.name.replace(/\.[^.]+$/, '').trim()
        mutate((draft) => {
          let club = draft.clubes.find((item) => slug(item.nombre) === slug(nombreClubArchivo))
          if (!club) {
            club = { id: id('club'), nombre: nombreClubArchivo, color: '#28e67a', logo: '' }
            draft.clubes.push(club)
          }
          club.logo = reader.result
        }, `Escudo de club cargado: ${nombreClubArchivo}`)
      }
      reader.readAsDataURL(file)
    })
  }

  async function importarClubes(file) {
    if (!file) return
    const rows = await readSheetRows(file)
    mutate((draft) => {
      rows.forEach((row, index) => {
        const nombre = cell(row, ['Club', 'Nombre'])
        if (!nombre) return
        const existente = draft.clubes.find((item) => slug(item.nombre) === slug(nombre))
        if (existente) {
          existente.color = cell(row, ['Color']) || existente.color
        } else {
          draft.clubes.push({
            id: id('club'),
            nombre,
            color: cell(row, ['Color']) || ['#28e67a', '#38bdf8', '#f59e0b', '#a3e635'][index % 4],
            logo: '',
          })
        }
      })
    }, `Clubes importados: ${file.name}`)
  }

  async function importarTecnicas(file) {
    if (!file) return
    const rows = await readSheetRows(file)
    mutate((draft) => {
      rows.forEach((row) => {
        const nombre = cell(row, ['Tecnica', 'Técnica', 'Profesora', 'Nombre'])
        if (!nombre) return
        const clubNombre = cell(row, ['Club'])
        let clubId = ''
        if (clubNombre) {
          let club = draft.clubes.find((item) => slug(item.nombre) === slug(clubNombre))
          if (!club) {
            club = { id: id('club'), nombre: clubNombre, color: '#28e67a', logo: '' }
            draft.clubes.push(club)
          }
          clubId = club.id
        }
        const existente = draft.tecnicas.find((item) => slug(item.nombre) === slug(nombre))
        if (existente) {
          existente.clubId = clubId || existente.clubId
        } else {
          draft.tecnicas.push({ id: id('tec'), nombre, clubId, foto: '' })
        }
      })
    }, `Técnicas importadas: ${file.name}`)
  }

  async function importarCategorias(file) {
    if (!file) return
    const rows = await readSheetRows(file)
    mutate((draft) => {
      rows.forEach((row) => {
        const nombre = cell(row, ['Categoria', 'Categoría', 'Nombre'])
        if (!nombre) return
        const torneoNombre = cell(row, ['Torneo'])
        let torneoId = pista.torneoId
        if (torneoNombre) {
          let torneo = draft.torneos.find((item) => slug(item.nombre) === slug(torneoNombre))
          if (!torneo) {
            torneo = { id: id('tor'), nombre: torneoNombre, liga: '', clubOrganizadorId: '', sede: '', fechaDesde: new Date().toISOString().slice(0, 10), fechaHasta: new Date().toISOString().slice(0, 10), turnosPorDia: 1, turnos: ['Único'] }
            draft.torneos.push(torneo)
          }
          torneoId = torneo.id
        }
        const existente = draft.categorias.find((item) => slug(item.nombre) === slug(nombre) && item.torneoId === torneoId)
        const dia = cell(row, ['Dia', 'Día', 'Fecha']) || draft.torneos.find((torneo) => torneo.id === torneoId)?.fechaDesde || ''
        const turno = cell(row, ['Turno']) || draft.torneos.find((torneo) => torneo.id === torneoId)?.turnos?.[0] || 'Único'
        const orden = Number(cell(row, ['Orden', 'Orden categoria', 'Orden categoría', 'Orden de categoria', 'Orden de categoría'])) || draft.categorias.length + 1
        if (existente) {
          existente.dia = dia
          existente.turno = turno
          existente.orden = orden
        } else {
          draft.categorias.push({ id: id('cat'), nombre, torneoId, dia, turno, orden })
        }
      })
    }, `Categorías importadas: ${file.name}`)
  }

  async function importarJueces(file) {
    if (!file) return
    const rows = await readSheetRows(file)
    mutate((draft) => {
      rows.forEach((row) => {
        const nombre = cell(row, ['Juez', 'Nombre'])
        if (!nombre) return
        const rol = cell(row, ['Rol', 'Tipo']) || 'General'
        const existente = draft.jueces.find((item) => slug(item.nombre) === slug(nombre))
        if (existente) existente.rol = rol
        else draft.jueces.push({ id: id('juez'), nombre, rol })
      })
    }, `Jueces importados: ${file.name}`)
  }

  return (
    <div className="app" data-theme={data.tema || 'verde'}>
      <header className="top">
        <div>
          <span>Patín Score Pro</span>
          <h1>{actual.torneo?.nombre}</h1>
          <p>{actual.torneo?.liga} - {actual.torneo?.sede} - {actual.torneo?.fechaDesde} al {actual.torneo?.fechaHasta} - inicio {actual.torneo?.horarioInicio || '-'}</p>
        </div>
        <div className="organizer-crest">
          <Avatar src={clubOrganizador?.logo} label={clubOrganizador?.nombre} color={clubOrganizador?.color} />
        </div>
        <StatusInfo clima />
        <strong>{pista.estado}</strong>
      </header>

      <nav className="tabs">
        {tabsVisibles.map((item) => <button className={tab === item ? 'active' : ''} key={item} onClick={() => setTab(item)}>{item}</button>)}
      </nav>

      {tab === 'Operador' && (
        <main className="grid two operator-grid">
          <Panel title="Evento">
            <ConfigEvento data={data} mutate={mutate} torneoId={pista.torneoId} iniciarTorneo={iniciarTorneo} programarInicioTorneo={programarInicioTorneo} rehabilitarInicioTorneo={rehabilitarInicioTorneo} />
          </Panel>
          <Panel title="Torneos guardados">
            <TorneosGuardados data={data} mutate={mutate} pista={pista} setPista={setPista} esAdmin={esAdmin} />
          </Panel>
          <Panel title="Pista" className="operator-pista-panel">
            <div className="operator-sections">
              <section>
                <span>Orden de pista</span>
                <div className="fields">
                  <label>Categoría<select value={pista.categoriaId} onChange={(e) => cambiarCategoria(e.target.value)}>{[...data.categorias].sort(compararCategoria).map((item) => <option key={item.id} value={item.id}>{item.dia} {item.turno} - {item.orden}. {item.nombre}</option>)}</select></label>
                  <label>Patinadora<select value={pista.patinadoraId} onChange={(e) => setPista({ ...pista, patinadoraId: e.target.value, juezCambioHasta: null, juezSiguientePatinadoraId: null })}>{patinadorasCategoria.map((item) => <option key={item.id} value={item.id}>{item.dia} {item.turno} - {item.ordenSalida || item.orden}. {item.nombre}</option>)}</select></label>
                </div>
              </section>
              <section>
                <span>Configuración de puntuación y tiempos</span>
                <div className="fields">
                  <label>Modo<select value={pista.modo} onChange={(e) => setPista({ ...pista, modo: e.target.value })}><option value="jueces">Varios jueces</option><option value="consenso">Consensuado</option></select></label>
                  <label>Receso minutos<input type="number" min="1" value={pista.recesoMin} onChange={(e) => setPista({ ...pista, recesoMin: e.target.value })} /></label>
                  <label>Segundos visible para jueces<input type="number" min="0" value={pista.juecesPuntajeSeg ?? 5} onChange={(e) => setPista({ ...pista, juecesPuntajeSeg: e.target.value })} /></label>
                  <label>Segundos antes de cambiar a pantalla de receso<input type="number" min="0" value={pista.demoraRecesoSeg} onChange={(e) => setPista({ ...pista, demoraRecesoSeg: e.target.value })} /></label>
                </div>
              </section>
            </div>
            <Competidor actual={actual} total={totalActual} />
            <div className="actions operator-actions">
              <button onClick={() => setPista({ ...pista, estado: 'En pista' })}>En pista</button>
              <button onClick={() => mover(-1)}>Anterior</button>
              <button onClick={() => mover(1)}>Siguiente</button>
              <button onClick={pasarSiguienteCategoria}>Siguiente categoría</button>
              <button onClick={() => { setPista({ ...pista, estado: 'Receso', recesoHasta: Date.now() + Number(pista.recesoMin) * 60000 }); log(`Receso ${pista.recesoMin} min`) }}>Receso</button>
              <button className="danger" onClick={marcarAusente}>No se presentó</button>
              <button onClick={postergar}>Postergar al final</button>
            </div>
            <FlujoCategoria data={data} pista={pista} actual={actual} mostrarProxima={false} />
          </Panel>
          <Panel title="Música y confirmación">
            <div className="music-now">
              <small>Tema asignado</small>
              <strong>{actual.cancion?.nombre || actual.patinadora?.musica || 'Sin tema'}</strong>
              <span>ID: {actual.patinadora?.musicaId || '-'}</span>
            </div>
            {actual.cancion?.archivo || actual.patinadora?.audio ? <audio ref={audioRef} controls src={actual.cancion?.archivo || actual.patinadora.audio} /> : <div className="empty">El administrador carga audio con mismo ID de la patinadora.</div>}
            <div className="actions">
              <button onClick={silbato}>Silbato</button>
              <button className="primary" onClick={playActual}>Reproducir actual</button>
              <button onClick={() => audioRef.current?.pause()}>Pausa</button>
              <button className="danger" onClick={() => setPista({ ...pista, estado: 'Detenido' })}>Detener</button>
            </div>
            <div className="confirm">
              <b>{actual.puntaje?.confirmado ? 'Publicado por juez' : 'Pendiente de publicación del juez'}</b>
            </div>
          </Panel>
          <Panel title="Colores del sistema">
            <ConfigTema data={data} mutate={mutate} />
          </Panel>
        </main>
      )}

      {tab === 'Jueces' && <Jueces data={data} pista={pista} actual={actual} juezId={juezId} setJuezId={setJuezId} actualizarPuntaje={actualizarPuntaje} publicarPuntaje={publicarPuntaje} iniciarPruebaPista={iniciarPruebaPista} silbato={silbato} marcarAusente={marcarAusente} postergar={postergar} />}
      {tab === 'Pública LED' && <main className="screen-wrap"><FullscreenButton targetSelector=".screen-wrap" /><Publica pista={pista} actual={actual} total={totalActual} data={data} modo={pista.modo} iniciarTorneo={iniciarTorneo} /></main>}
      {tab === 'Datos' && <Datos data={data} mutate={mutate} torneoId={pista.torneoId} guardarArchivo={guardarArchivo} importarListado={importarListado} cargarCanciones={cargarCanciones} cargarEscudosClubes={cargarEscudosClubes} importarClubes={importarClubes} importarTecnicas={importarTecnicas} importarCategorias={importarCategorias} importarJueces={importarJueces} iniciarTorneo={iniciarTorneo} programarInicioTorneo={programarInicioTorneo} rehabilitarInicioTorneo={rehabilitarInicioTorneo} />}
      {tab === 'Web pública' && <WebPublica data={data} modo={pista.modo} pista={pistaPublica} />}
      {tab === 'Reportes' && <Reportes data={data} />}
      {tab === 'Tanteador' && <main className="screen-wrap"><FullscreenButton targetSelector=".screen-wrap" /><Tanteador data={data} categoriaId={pista.categoriaId} modo={pista.modo} /></main>}
      {tab === 'Ranking clubes' && <RankingClubes data={data} modo={pista.modo} />}
      {tab === 'Actas' && <Actas data={data} modo={pista.modo} />}
      {tab === 'Registros' && <Logs data={data} />}

      <footer>
        <button className="ghost danger" onClick={() => { localStorage.removeItem(STORAGE_KEY); setData(demo); setPista(basePista) }}>Reiniciar demo</button>
        <label className="user-select">Usuario<select value={usuarioId} onChange={(event) => { setUsuarioId(event.target.value); setTab((data.usuarios.find((item) => item.id === event.target.value)?.tabs || tabs)[0]) }}>{data.usuarios.map((usuario) => <option key={usuario.id} value={usuario.id}>{usuario.nombre}</option>)}</select></label>
        <span>Persistencia {persistencia}{persistencia === 'localStorage' ? ' (iniciá npm run server para usar SQLite).' : ''}</span>
      </footer>
    </div>
  )
}

function Jueces({ data, pista, actual, juezId, setJuezId, actualizarPuntaje, publicarPuntaje, iniciarPruebaPista, silbato, marcarAusente, postergar }) {
  const puntaje = actual.puntaje || { jueces: {} }
  const [minutosPrueba, setMinutosPrueba] = useState(pista.pruebaPistaMin || 5)
  const miPuntaje = puntaje.jueces[juezId] || {}
  const totalTecnico = totalPuntaje(puntaje, pista.modo, data.conceptosPuntaje)
  const patinadorasOrdenadas = data.patinadoras
    .filter((item) => item.categoriaId === pista.categoriaId && item.estado !== 'ausente')
    .sort(compararSalida)
  const indiceActual = patinadorasOrdenadas.findIndex((item) => item.id === actual.patinadora?.id)
  const proxima = patinadorasOrdenadas[indiceActual + 1]
  const proximaClub = data.clubes.find((item) => item.id === proxima?.clubId)
  const categoriaSiguiente = siguienteCategoria(data, pista.categoriaId)
  const mismoTurnoSiguiente = categoriaSiguiente && categoriaSiguiente.dia === actual.categoria?.dia && categoriaSiguiente.turno === actual.categoria?.turno
  const tieneValor = pista.modo === 'consenso'
    ? Object.values(puntaje.consensoConceptos || {}).some((valor) => valor !== '')
    : Object.values(miPuntaje.conceptos || {}).some((valor) => valor !== '')

  function setConcepto(conceptoId, valor) {
    actualizarPuntaje(actual.patinadora.id, (draft) => {
      if (pista.modo === 'consenso') {
        draft.consensoConceptos ||= {}
        draft.consensoConceptos[conceptoId] = valor
      } else {
        draft.jueces[juezId] ||= {}
        draft.jueces[juezId].conceptos ||= {}
        draft.jueces[juezId].conceptos[conceptoId] = valor
        draft.jueces[juezId].enviado = new Date().toISOString()
      }
    }, `Puntaje cargado: ${actual.patinadora.nombre}`)
  }

  function guardarConEnter(event) {
    if (event.key !== 'Enter') return
    event.preventDefault()
    const inputs = Array.from(event.currentTarget.closest('.concept-list')?.querySelectorAll('input') || [])
    const index = inputs.indexOf(event.currentTarget)
    const next = inputs[index + 1]
    if (next) next.focus()
    else event.currentTarget.blur()
  }

  return (
    <main className="judge-screen">
      <Panel title="Solo puntuación">
        <div className="judge-head">
          {pista.modo === 'consenso' ? <div className="mode-lock">Modo consensuado</div> : <label>Juez<select value={juezId} onChange={(e) => setJuezId(e.target.value)}>{data.jueces.map((item) => <option key={item.id} value={item.id}>{item.nombre} - {item.rol}</option>)}</select></label>}
          <div>
            <small>{actual.categoria?.nombre}</small>
            <h2>{actual.patinadora?.nombre}</h2>
            <p>{actual.club?.nombre}</p>
            <b>Ubicación de salida {actual.patinadora?.ordenSalida || actual.patinadora?.orden || '-'}</b>
          </div>
        </div>
        <div className="concept-list">
          {data.conceptosPuntaje.map((concepto) => (
            <label key={concepto.id}>
              {concepto.nombre}
              <input
                type="number"
                min="0"
                max="10"
                step="0.1"
                value={pista.modo === 'consenso' ? puntaje.consensoConceptos?.[concepto.id] || '' : miPuntaje.conceptos?.[concepto.id] || ''}
                onChange={(event) => setConcepto(concepto.id, event.target.value)}
                onKeyDown={guardarConEnter}
              />
            </label>
          ))}
        </div>
        <div className="technical-total">
          <span>Total técnico preliminar</span>
          <strong>{totalTecnico == null ? '--' : totalTecnico.toFixed(2)}</strong>
        </div>
        <div className="publish-box">
          <button className="primary" disabled={!tieneValor || actual.puntaje?.confirmado} onClick={publicarPuntaje}>Publicar puntaje</button>
          <span>{actual.puntaje?.confirmado ? 'Publicado' : tieneValor ? 'Listo para publicar' : 'Esperando puntuación'}</span>
        </div>
        <div className="judge-actions">
          <button onClick={silbato}>Silbato</button>
          <button className="danger" onClick={marcarAusente}>No se presentó</button>
          <button onClick={postergar}>Postergar al final</button>
          <label>Minutos prueba pista<input type="number" min="1" value={minutosPrueba} onChange={(event) => setMinutosPrueba(event.target.value)} /></label>
          <button className="primary" onClick={() => iniciarPruebaPista(minutosPrueba)}>Probando pista</button>
        </div>
        <div className="next-skater">
          <span>Siguiente patinadora</span>
          <strong>{proxima?.nombre || 'Última patinadora de la categoría'}</strong>
          {proxima && <small>{proximaClub?.nombre} - {proxima.dia} - {proxima.turno} - salida {proxima.ordenSalida || proxima.orden}</small>}
        </div>
        <FlujoCategoria data={data} pista={pista} actual={actual} mostrarProxima={false} />
        {!proxima && (
          <div className="category-end-actions judge-end-status">
            <div className="end-status">
              <span>{mismoTurnoSiguiente ? 'Categoría finalizada' : 'Turno finalizado'}</span>
              <strong>{mismoTurnoSiguiente ? categoriaSiguiente?.nombre : `${actual.categoria?.dia || ''} ${actual.categoria?.turno || ''}`}</strong>
            </div>
          </div>
        )}
        <div className="legacy-score"><ScorePad value={pista.modo === 'consenso' ? puntaje.consenso || '' : miPuntaje.valor || ''} onChange={(valor) => actualizarPuntaje(actual.patinadora.id, (draft) => {
          if (pista.modo === 'consenso') draft.consenso = valor
          else {
            draft.jueces[juezId] ||= {}
            draft.jueces[juezId].valor = valor
            draft.jueces[juezId].enviado = new Date().toISOString()
          }
        }, `Puntaje cargado: ${actual.patinadora.nombre}`)} />
        <div className="sent-state">{miPuntaje.valor || puntaje.consenso ? 'Puntaje enviado' : 'Esperando puntuación'}</div>
        </div>
      </Panel>
    </main>
  )
}

function Publica({ pista, actual, total, data, modo, iniciarTorneo }) {
  const [now, setNow] = useState(() => Date.now())
  useEffect(() => {
    const timer = setInterval(() => setNow(Date.now()), 500)
    return () => clearInterval(timer)
  }, [])

  const inicioProgramadoHasta = pista.inicioProgramadoHasta || actual.torneo?.inicioProgramadoHasta
  useEffect(() => {
    if (!actual.torneo?.inicioTorneoEn && inicioProgramadoHasta && now >= inicioProgramadoHasta) {
      iniciarTorneo?.()
    }
  }, [actual.torneo?.inicioTorneoEn, inicioProgramadoHasta, iniciarTorneo, now])

  const mostrandoPuntaje = pista.puntajeHasta && now < pista.puntajeHasta && pista.puntajePatinadoraId
  const mostrandoPrueba = pista.pruebaPistaHasta && now < pista.pruebaPistaHasta
  const mostrandoRecesoAuto = pista.categoriaFinalHasta && pista.autoRecesoHasta && now >= pista.categoriaFinalHasta && now < pista.autoRecesoHasta
  const patinadoraPuntaje = data.patinadoras.find((item) => item.id === pista.puntajePatinadoraId)
  const clubPuntaje = data.clubes.find((item) => item.id === patinadoraPuntaje?.clubId)
  const totalPuntajeReciente = totalPuntaje(data.puntajes[pista.puntajePatinadoraId], modo, data.conceptosPuntaje)
  const recesoHasta = pista.recesoHasta || pista.autoRecesoHasta
  const proximaCategoria = siguienteCategoria(data, pista.categoriaId)
  const clubOrganizador = data.clubes.find((club) => club.id === actual.torneo?.clubOrganizadorId)

  if (!actual.torneo?.inicioTorneoEn && inicioProgramadoHasta && now < inicioProgramadoHasta) {
    return <InicioTorneoPantalla actual={actual} hasta={inicioProgramadoHasta} clubOrganizador={clubOrganizador} />
  }

  if (mostrandoPrueba) {
    const patinadoras = data.patinadoras
      .filter((item) => item.categoriaId === pista.categoriaId && item.estado !== 'ausente')
      .sort(compararSalida)
    return <PausaPantalla tipo="Probando pista" actual={actual} hasta={pista.pruebaPistaHasta} patinadoras={patinadoras} clubes={data.clubes} />
  }

  if (pista.estado === 'Receso' || mostrandoRecesoAuto) {
    return (
      <main className="break-screen">
        {actual.torneo?.ligaLogo && <img className="league-logo" src={actual.torneo.ligaLogo} alt={actual.torneo?.liga || 'Liga'} />}
        <section>
          <div className="break-crest">
            <Avatar src={clubOrganizador?.logo} label={clubOrganizador?.nombre} color={clubOrganizador?.color} />
          </div>
          <h2>Receso</h2>
        </section>
        <div className="break-timer">
          {recesoHasta ? <Timer hasta={recesoHasta} /> : <b>--:--</b>}
          <small>Cuenta regresiva</small>
        </div>
        <div className="break-info">
          <div><span>Torneo</span><strong>{actual.torneo?.nombre || '-'}</strong></div>
          <div><span>Liga</span><strong>{actual.torneo?.liga || '-'}</strong></div>
          <div><span>Siguiente categoría</span><strong>{proximaCategoria?.nombre || 'A confirmar'}</strong></div>
          <div><span>Tiempo de receso</span><strong>{pista.recesoMin || 0} min</strong></div>
        </div>
      </main>
    )
  }

  if (mostrandoPuntaje) {
    return (
      <main className="led score-reveal">
        <div className="led-competitor">
          <div className="led-meta">
            <span>Categoría</span>
            <strong>{actual.categoria?.nombre || '-'}</strong>
          </div>
          <p>Puntaje publicado</p>
          <h2>{patinadoraPuntaje?.nombre}</h2>
          <div className="led-club">
            <Avatar src={clubPuntaje?.logo} label={clubPuntaje?.nombre} color={clubPuntaje?.color} />
            <span>Club</span>
            <strong>{clubPuntaje?.nombre || '-'}</strong>
          </div>
        </div>
        <aside className="score-card">
          <span>Puntuación</span>
          <strong>{totalPuntajeReciente == null ? '--' : totalPuntajeReciente.toFixed(2)}</strong>
          <small>Publicado por jueces</small>
        </aside>
      </main>
    )
  }

  const rows = rankingCategoria(data, pista.categoriaId, modo).filter((row) => row.confirmado && row.total != null)
  if (rows.some((row) => row.total != null)) {
    return <TanteadorLed categoria={actual.categoria} rows={rows} clubes={data.clubes} data={data} modo={modo} />
  }

  const puestoActual = rankingCategoria(data, pista.categoriaId, modo).find((row) => row.id === actual.patinadora?.id)?.puesto

  return (
    <main className="led">
      <div className="led-competitor">
        <div className="led-meta">
          <span>Categoría</span>
          <strong>{actual.categoria?.nombre || '-'}</strong>
        </div>
        <h2>{actual.patinadora?.nombre}</h2>
        <div className="led-club">
          <Avatar src={actual.club?.logo} label={actual.club?.nombre} color={actual.club?.color} />
          <span>Club</span>
          <strong>{actual.club?.nombre || '-'}</strong>
        </div>
      </div>
      <aside className="score-card">
        <div className="score-slot">
          <span>Puntuación</span>
          <strong>{total == null ? '--' : total.toFixed(2)}</strong>
          <small>{total == null ? 'Esperando puntaje' : 'Puntaje cargado'}</small>
        </div>
        <div className="score-slot">
          <span>Posición</span>
          <strong>{puestoActual || '--'}</strong>
          <small>{puestoActual ? 'Tanteador' : 'En espera'}</small>
        </div>
      </aside>
    </main>
  )
}

function InicioTorneoPantalla({ actual, hasta, clubOrganizador }) {
  return (
    <main className="break-screen tournament-start-screen">
      {actual.torneo?.ligaLogo && <img className="league-logo" src={actual.torneo.ligaLogo} alt={actual.torneo?.liga || 'Liga'} />}
      <section>
        <div className="break-crest">
          <Avatar src={clubOrganizador?.logo} label={clubOrganizador?.nombre} color={clubOrganizador?.color} />
        </div>
        <span className="start-kicker">Inicio de torneo</span>
        <h2>{actual.torneo?.nombre || 'Torneo'}</h2>
        <div className="start-meta">
          {actual.torneo?.liga && <span>{actual.torneo.liga}</span>}
          {actual.torneo?.sede && <span>{actual.torneo.sede}</span>}
        </div>
      </section>
      <div className="break-timer">
        <Timer hasta={hasta} />
        <small>Comienza en</small>
      </div>
    </main>
  )
}

function PausaPantalla({ tipo, actual, hasta, patinadoras, clubes }) {
  return (
    <main className="break-screen practice-screen">
      <section>
        <span>{actual.torneo?.nombre}</span>
        <h2>{tipo}</h2>
        <p>{actual.categoria?.nombre}</p>
      </section>
      <div className="break-timer">
        <Timer hasta={hasta} />
        <small>Cuenta regresiva</small>
      </div>
      <div className="practice-list">
        <span>Patinadoras en pista</span>
        <div>
          {patinadoras.map((item) => {
            const club = clubes.find((clubItem) => clubItem.id === item.clubId)
            return (
              <article key={item.id}>
                <b>{item.ordenSalida || item.orden}</b>
                <Avatar src={club?.logo} label={club?.nombre} color={club?.color} />
                <strong>{item.nombre}</strong>
                <small>{club?.nombre || '-'}</small>
              </article>
            )
          })}
        </div>
      </div>
    </main>
  )
}

function TanteadorLed({ categoria, rows, clubes, data, modo }) {
  const pageSize = 8
  const [page, setPage] = useState(0)
  const pageCount = Math.max(1, Math.ceil(rows.length / pageSize))
  const currentPage = page % pageCount
  const visibles = rows.slice(currentPage * pageSize, currentPage * pageSize + pageSize)
  const columnas = rows[0] ? parcialesPuntaje(data.puntajes[rows[0].id], modo, data.conceptosPuntaje).filter((item) => item.valor != null) : []

  useEffect(() => {
    if (pageCount <= 1) {
      return undefined
    }
    const timer = setInterval(() => setPage((prev) => (prev + 1) % pageCount), 7000)
    return () => clearInterval(timer)
  }, [pageCount])

  return (
    <main className="led-board">
      <header>
        <span>Tanteador</span>
        <h2>{categoria?.nombre}</h2>
      </header>
      <div className="board-table" style={{ '--score-cols': columnas.length }}>
        <div className="board-row board-head">
          <span>Pos</span>
          <span>Patinadora</span>
          <span>Club</span>
          {columnas.map((columna) => <span key={columna.id}>{columna.nombre}</span>)}
          <span>Total</span>
        </div>
        {visibles.map((row) => {
          const parciales = parcialesPuntaje(data.puntajes[row.id], modo, data.conceptosPuntaje)
          const club = clubes.find((item) => item.id === row.clubId)
          return (
            <div className="board-row" key={row.id}>
              <strong>{row.puesto || '-'}</strong>
              <span>{row.nombre}</span>
              <div className="board-club">
                <Avatar src={club?.logo} label={club?.nombre} color={club?.color} />
                <small>{club?.nombre || '-'}</small>
              </div>
              {columnas.map((columna) => {
                const parcial = parciales.find((item) => item.id === columna.id)
                return <b key={columna.id}>{parcial?.valor == null ? '-' : parcial.valor.toFixed(2)}</b>
              })}
              <b>{row.total == null ? '--' : row.total.toFixed(2)}</b>
            </div>
          )
        })}
      </div>
      {pageCount > 1 && <em>Página {currentPage + 1} de {pageCount}</em>}
    </main>
  )
}

function Datos({ data, mutate, torneoId, guardarArchivo, importarListado, cargarCanciones, cargarEscudosClubes, importarClubes, importarTecnicas, importarCategorias, importarJueces, iniciarTorneo, programarInicioTorneo, rehabilitarInicioTorneo }) {
  return (
    <main className="grid admin-grid">
      <Panel title="Configurar evento">
        <ConfigEvento data={data} mutate={mutate} torneoId={torneoId} iniciarTorneo={iniciarTorneo} programarInicioTorneo={programarInicioTorneo} rehabilitarInicioTorneo={rehabilitarInicioTorneo} />
      </Panel>
      <Panel title="Importar patinadoras">
        <div className="import-box">
          <p>Columnas: Nombre, Club, Categoría, Música, Día, Turno, Orden de salida. Opcional: Edad.</p>
          <label>Excel o CSV<input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => importarListado(e.target.files[0])} /></label>
        </div>
      </Panel>
      <Panel title="Importar clubes">
        <div className="import-box">
          <p>Columnas: Club o Nombre. Opcional: Color.</p>
          <label>Excel o CSV<input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => importarClubes(e.target.files[0])} /></label>
        </div>
      </Panel>
      <Panel title="Importar categorías">
        <div className="import-box">
          <p>Columnas: Categoría o Nombre. Opcional: Torneo, Día, Turno, Orden de categoría.</p>
          <label>Excel o CSV<input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => importarCategorias(e.target.files[0])} /></label>
        </div>
      </Panel>
      <Panel title="Importar profesoras">
        <div className="import-box">
          <p>Columnas: Profesora o Técnica o Nombre. Opcional: Club.</p>
          <label>Excel o CSV<input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => importarTecnicas(e.target.files[0])} /></label>
        </div>
      </Panel>
      <Panel title="Importar jueces">
        <div className="import-box">
          <p>Columnas: Juez o Nombre. Opcional: Rol.</p>
          <label>Excel o CSV<input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => importarJueces(e.target.files[0])} /></label>
        </div>
      </Panel>
      <Panel title="Cargar canciones">
        <div className="import-box">
          <p>Nombre de archivo identifica música. Ejemplo: libertango.mp3 coincide con MúsicaId libertango.</p>
          <label>Audios<input type="file" multiple accept="audio/*" onChange={(e) => cargarCanciones(e.target.files)} /></label>
          <div className="song-list">{data.canciones.map((item) => <div key={item.id}><b>{item.id}</b><span>{item.archivo ? 'cargada' : 'sin audio'}</span></div>)}</div>
        </div>
      </Panel>
      <Panel title="Cargar escudos de clubes">
        <div className="import-box">
          <p>El nombre del archivo debe coincidir con el club. Ejemplo: Sol Patín.png se asocia al club Sol Patín.</p>
          <label>Escudos<input type="file" multiple accept="image/*" onChange={(e) => cargarEscudosClubes(e.target.files)} /></label>
        </div>
      </Panel>
      <Panel title="Configurar puntaje">
        <ConfigPuntaje data={data} mutate={mutate} />
      </Panel>
      <Panel title="Clubes">
        <EditorEntidades
          placeholder="Nombre del club"
          ayudaArchivo="Logo del club"
          onFile={(itemId, file) => guardarArchivo('img', 'clubes', itemId, 'logo', file)}
          onAdd={(nombre, itemId) => mutate((draft) => { draft.clubes.push({ id: itemId, nombre, color: '#28e67a', logo: '' }) }, `Club agregado: ${nombre}`)}
        />
      </Panel>
      <Panel title="Profesoras">
        <EditorEntidades
          placeholder="Nombre de la profesora"
          ayudaArchivo="Foto de profesora"
          onFile={(itemId, file) => guardarArchivo('img', 'tecnicas', itemId, 'foto', file)}
          onAdd={(nombre, itemId) => mutate((draft) => { draft.tecnicas.push({ id: itemId, nombre, clubId: data.clubes[0]?.id || '', foto: '' }) }, `Profesora agregada: ${nombre}`)}
        />
      </Panel>
      <Panel title="Orden de categorías">
        <OrdenCategorias data={data} mutate={mutate} />
      </Panel>
      <Panel title="Jueces">
        <EditorJueces data={data} mutate={mutate} />
      </Panel>
      <Panel title="Buffet">
        <ConfigBuffet data={data} mutate={mutate} />
      </Panel>
      <Panel title="Usuarios">
        <ConfigUsuarios data={data} mutate={mutate} />
      </Panel>
      <Panel title="Patinadoras">
        <EditorPatinadoras data={data} mutate={mutate} />
      </Panel>
    </main>
  )
}

function ConfigEvento({ data, mutate, torneoId, iniciarTorneo, programarInicioTorneo, rehabilitarInicioTorneo }) {
  const torneo = data.torneos.find((item) => item.id === torneoId) || data.torneos[0]
  const [ok, setOk] = useState(false)
  const [minutosInicio, setMinutosInicio] = useState(5)
  const bloqueado = Boolean(torneo.confirmado)
  const grabado = torneo.eventoGrabado
  const inicioBloqueado = Boolean(torneo.inicioTorneoEn)
  const inicioProgramado = Boolean(torneo.inicioProgramadoHasta && !inicioBloqueado)

  function cambiar(campo, valor) {
    if (bloqueado) return
    mutate((draft) => {
      const actual = draft.torneos.find((item) => item.id === torneo.id) || draft.torneos[0]
      actual[campo] = valor
    })
  }

  function cambiarTurnos(cantidad) {
    if (bloqueado) return
    const total = Math.max(1, Number(cantidad) || 1)
    mutate((draft) => {
      const actual = draft.torneos.find((item) => item.id === torneo.id) || draft.torneos[0]
      actual.turnosPorDia = total
      const nombresBase = ['Mañana', 'Tarde', 'Noche']
      actual.turnos = Array.from({ length: total }, (_, index) => actual.turnos?.[index] || nombresBase[index] || `Turno ${index + 1}`)
    }, 'Turnos del evento actualizados')
  }

  function cambiarNombreTurno(index, valor) {
    if (bloqueado) return
    mutate((draft) => {
      const actual = draft.torneos.find((item) => item.id === torneo.id) || draft.torneos[0]
      actual.turnos[index] = valor
    })
  }

  function confirmar() {
    mutate((draft) => {
      const actual = draft.torneos.find((item) => item.id === torneo.id) || draft.torneos[0]
      actual.confirmado = true
      actual.activo = true
      actual.confirmadoEn = new Date().toISOString()
      actual.eventoGrabado = {
        nombre: actual.nombre,
        liga: actual.liga,
        clubOrganizadorId: actual.clubOrganizadorId,
        sede: actual.sede,
        fechaDesde: actual.fechaDesde,
        fechaHasta: actual.fechaHasta,
        horarioInicio: actual.horarioInicio,
        turnos: actual.turnos,
      }
    }, `Evento confirmado: ${torneo.nombre}`)
    setOk(true)
    window.setTimeout(() => setOk(false), 1300)
  }

  function habilitarEdicion(checked) {
    mutate((draft) => {
      const actual = draft.torneos.find((item) => item.id === torneo.id) || draft.torneos[0]
      actual.confirmado = !checked
    }, checked ? `Edición de evento habilitada: ${torneo.nombre}` : `Evento bloqueado: ${torneo.nombre}`)
  }

  function cargarEscudoOrganizador(file) {
    if (bloqueado || !file || !torneo.clubOrganizadorId) return
    const reader = new FileReader()
    reader.onload = () => {
      mutate((draft) => {
        const actual = draft.torneos.find((item) => item.id === torneo.id) || draft.torneos[0]
        const club = draft.clubes.find((item) => item.id === actual.clubOrganizadorId)
        if (club) club.logo = reader.result
      }, 'Escudo del club organizador cargado')
    }
    reader.readAsDataURL(file)
  }

  function cargarLogoLiga(file) {
    if (bloqueado || !file) return
    const reader = new FileReader()
    reader.onload = () => {
      mutate((draft) => {
        const actual = draft.torneos.find((item) => item.id === torneo.id) || draft.torneos[0]
        actual.ligaLogo = reader.result
      }, 'Logo de liga cargado')
    }
    reader.readAsDataURL(file)
  }

  return (
    <div className="event-config">
      <label className="edit-toggle">Editar evento<input type="checkbox" checked={!bloqueado} onChange={(event) => habilitarEdicion(event.target.checked)} /></label>
      <div className="fields">
        <label>Nombre del evento<input disabled={bloqueado} value={torneo.nombre} onChange={(event) => cambiar('nombre', event.target.value)} /></label>
        <label>Liga<input disabled={bloqueado} value={torneo.liga || ''} onChange={(event) => cambiar('liga', event.target.value)} /></label>
        <label>Club organizador<select disabled={bloqueado} value={torneo.clubOrganizadorId || ''} onChange={(event) => cambiar('clubOrganizadorId', event.target.value)}><option value="">Sin asignar</option>{data.clubes.map((club) => <option key={club.id} value={club.id}>{club.nombre}</option>)}</select></label>
        <label>Sede<input disabled={bloqueado} value={torneo.sede || ''} onChange={(event) => cambiar('sede', event.target.value)} /></label>
        <label>Fecha desde<input disabled={bloqueado} type="date" value={torneo.fechaDesde || ''} onChange={(event) => cambiar('fechaDesde', event.target.value)} /></label>
        <label>Fecha hasta<input disabled={bloqueado} type="date" value={torneo.fechaHasta || ''} onChange={(event) => cambiar('fechaHasta', event.target.value)} /></label>
        <label>Horario de inicio<input disabled={bloqueado} type="time" value={torneo.horarioInicio || ''} onChange={(event) => cambiar('horarioInicio', event.target.value)} /></label>
        <label>Turnos por día<input disabled={bloqueado} type="number" min="1" max="4" value={torneo.turnosPorDia || 1} onChange={(event) => cambiarTurnos(event.target.value)} /></label>
        <label>Logo liga<input disabled={bloqueado} type="file" accept="image/*" onChange={(event) => cargarLogoLiga(event.target.files[0])} /></label>
        <label>Escudo club organizador<input disabled={bloqueado} type="file" accept="image/*" onChange={(event) => cargarEscudoOrganizador(event.target.files[0])} /></label>
      </div>
      <div className="turnos-config">
        {(torneo.turnos || ['Único']).map((turno, index) => (
          <label key={index}>Turno {index + 1}<input disabled={bloqueado} value={turno} onChange={(event) => cambiarNombreTurno(index, event.target.value)} /></label>
        ))}
      </div>
      <small>Días configurados: {diasTorneo(torneo).join(', ') || 'sin fecha'}</small>
      <div className="start-row">
        <button className="start-button" disabled={inicioBloqueado} onClick={iniciarTorneo}>{inicioBloqueado ? 'Torneo iniciado' : 'Inicio de Torneo'}</button>
        <label>Minutos para iniciar<input disabled={inicioBloqueado} type="number" min="1" value={minutosInicio} onChange={(event) => setMinutosInicio(event.target.value)} /></label>
        <button className="start-button ghost-start" disabled={inicioBloqueado} onClick={() => programarInicioTorneo(minutosInicio)}>{inicioProgramado ? 'Inicio programado' : 'Programar inicio'}</button>
        {torneo.inicioTorneoEn && <button className="ghost" disabled={bloqueado} onClick={rehabilitarInicioTorneo}>Rehabilitar inicio</button>}
      </div>
      {inicioProgramado && (
        <div className="scheduled-start">
          <span>Cuenta regresiva para inicio</span>
          <strong><Timer hasta={torneo.inicioProgramadoHasta} /></strong>
        </div>
      )}
      <button className="ok-button" disabled={bloqueado} onClick={confirmar}>{ok ? 'OK ✓' : 'OK'}</button>
      {grabado && (
        <div className="event-summary">
          <span>Evento grabado</span>
          <strong>{grabado.nombre}</strong>
          <p>{grabado.liga || '-'} · {grabado.sede || '-'}</p>
          <p>{grabado.fechaDesde} al {grabado.fechaHasta} · inicio {grabado.horarioInicio || '-'} · turnos: {(grabado.turnos || []).join(', ')}</p>
        </div>
      )}
    </div>
  )
}

function TorneosGuardados({ data, mutate, pista, setPista, esAdmin }) {
  const torneosGrabados = data.torneos.filter((torneo) => torneo.eventoGrabado || torneo.confirmado)

  function primeraCategoria(torneoId) {
    return [...data.categorias].filter((categoria) => categoria.torneoId === torneoId).sort(compararCategoria)[0]
  }

  function primeraPatinadora(categoriaId) {
    return [...data.patinadoras].filter((patinadora) => patinadora.categoriaId === categoriaId && patinadora.estado !== 'ausente').sort(compararSalida)[0]
  }

  function activar(torneoId) {
    if (!esAdmin) return
    const categoria = primeraCategoria(torneoId)
    const patinadora = primeraPatinadora(categoria?.id)
    mutate((draft) => {
      draft.torneos.forEach((torneo) => {
        torneo.activo = torneo.id === torneoId
      })
    }, `Torneo activado: ${data.torneos.find((torneo) => torneo.id === torneoId)?.nombre || torneoId}`)
    setPista((prev) => ({
      ...prev,
      torneoId,
      categoriaId: categoria?.id || '',
      patinadoraId: patinadora?.id || '',
      estado: 'Preparando',
      puntajeHasta: null,
      puntajePatinadoraId: null,
      recesoHasta: null,
      pruebaPistaHasta: null,
    }))
  }

  function desactivar(torneoId) {
    if (!esAdmin) return
    mutate((draft) => {
      const torneo = draft.torneos.find((item) => item.id === torneoId)
      if (torneo) torneo.activo = false
    }, `Torneo desactivado: ${data.torneos.find((torneo) => torneo.id === torneoId)?.nombre || torneoId}`)
  }

  function eliminar(torneoId) {
    if (!esAdmin || data.torneos.length <= 1) return
    if (!window.confirm('¿Eliminar torneo y datos asociados?')) return
    const categorias = new Set(data.categorias.filter((categoria) => categoria.torneoId === torneoId).map((categoria) => categoria.id))
    const patinadoras = new Set(data.patinadoras.filter((patinadora) => categorias.has(patinadora.categoriaId)).map((patinadora) => patinadora.id))
    mutate((draft) => {
      draft.torneos = draft.torneos.filter((torneo) => torneo.id !== torneoId)
      draft.categorias = draft.categorias.filter((categoria) => !categorias.has(categoria.id))
      draft.patinadoras = draft.patinadoras.filter((patinadora) => !patinadoras.has(patinadora.id))
      patinadoras.forEach((patinadoraId) => { delete draft.puntajes[patinadoraId] })
    }, `Torneo eliminado: ${data.torneos.find((torneo) => torneo.id === torneoId)?.nombre || torneoId}`)
    if (pista.torneoId === torneoId) {
      const siguiente = data.torneos.find((torneo) => torneo.id !== torneoId)
      if (siguiente) activar(siguiente.id)
    }
  }

  function crearTorneo() {
    if (!esAdmin) return
    const nuevo = {
      id: id('tor'),
      nombre: `Nuevo torneo ${data.torneos.length + 1}`,
      liga: '',
      ligaLogo: '',
      clubOrganizadorId: '',
      sede: '',
      fechaDesde: new Date().toISOString().slice(0, 10),
      fechaHasta: new Date().toISOString().slice(0, 10),
      horarioInicio: '09:00',
      turnosPorDia: 1,
      turnos: ['Único'],
      activo: true,
      confirmado: false,
    }
    mutate((draft) => {
      draft.torneos.forEach((torneo) => { torneo.activo = false })
      draft.torneos.push(nuevo)
    }, `Torneo creado: ${nuevo.nombre}`)
    setPista((prev) => ({ ...prev, torneoId: nuevo.id, categoriaId: '', patinadoraId: '', estado: 'Preparando' }))
  }

  function descargarLogs(torneoId) {
    const torneo = data.torneos.find((item) => item.id === torneoId)
    const rows = data.logs.filter((log) => !log.torneoId || log.torneoId === torneoId)
    const csv = [
      ['Fecha', 'Torneo', 'Registro'],
      ...rows.map((log) => [new Date(log.fecha).toLocaleString('es-AR'), torneo?.nombre || '', log.texto]),
    ].map((row) => row.map((cell) => `"${String(cell || '').replaceAll('"', '""')}"`).join(',')).join('\n')
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' })
    const link = document.createElement('a')
    link.href = URL.createObjectURL(blob)
    link.download = `logs-${slug(torneo?.nombre || torneoId)}.csv`
    link.click()
    URL.revokeObjectURL(link.href)
  }

  return (
    <div className="saved-events">
      <div className="saved-events-actions">
        <button disabled={!esAdmin} onClick={crearTorneo}>Nuevo torneo</button>
        {!esAdmin && <small>Solo el administrador puede administrar torneos.</small>}
      </div>
      <div>
        {(torneosGrabados.length ? torneosGrabados : data.torneos).map((torneo) => (
          <article className={pista.torneoId === torneo.id ? 'active' : ''} key={torneo.id}>
            <div>
              <span>{torneo.activo ? 'Activo' : 'Guardado'}</span>
              <strong>{torneo.nombre}</strong>
              <small>{torneo.liga || '-'} · {torneo.fechaDesde || '-'} al {torneo.fechaHasta || '-'}</small>
            </div>
            <button disabled={!esAdmin || pista.torneoId === torneo.id} onClick={() => activar(torneo.id)}>Activar</button>
            <button disabled={!esAdmin || !torneo.activo} onClick={() => desactivar(torneo.id)}>Desactivar</button>
            <button onClick={() => descargarLogs(torneo.id)}>Bajar logs</button>
            <button className="danger" disabled={!esAdmin || data.torneos.length <= 1} onClick={() => eliminar(torneo.id)}>Eliminar</button>
          </article>
        ))}
      </div>
    </div>
  )
}

function OrdenCategorias({ data, mutate }) {
  const torneo = data.torneos[0]
  const [form, setForm] = useState({
    categoriaId: data.categorias[0]?.id || '',
    nombre: '',
    dia: torneo.fechaDesde || '',
    turno: torneo.turnos?.[0] || 'Único',
    orden: 1,
  })
  const [ok, setOk] = useState(false)

  function cambiar(campo, valor) {
    setForm((prev) => ({ ...prev, [campo]: valor }))
  }

  function guardar() {
    const nombre = form.nombre.trim()
    if (!form.categoriaId && !nombre) return
    mutate((draft) => {
      let categoria = draft.categorias.find((item) => item.id === form.categoriaId)
      if (!categoria) {
        categoria = { id: id('cat'), nombre, torneoId: draft.torneos[0]?.id || 't1' }
        draft.categorias.push(categoria)
      }
      categoria.nombre = nombre || categoria.nombre
      categoria.dia = form.dia
      categoria.turno = form.turno
      categoria.orden = Number(form.orden) || 1
    }, 'Orden de categoría actualizado')
    setForm((prev) => ({ ...prev, nombre: '', orden: Number(prev.orden) + 1 }))
    setOk(true)
    window.setTimeout(() => setOk(false), 1300)
  }

  return (
    <div className="event-config">
      <div className="fields">
        <label>Categoría existente<select value={form.categoriaId} onChange={(event) => cambiar('categoriaId', event.target.value)}><option value="">Nueva categoría</option>{[...data.categorias].sort(compararCategoria).map((categoria) => <option key={categoria.id} value={categoria.id}>{categoria.nombre}</option>)}</select></label>
        <label>Nombre si es nueva<input value={form.nombre} onChange={(event) => cambiar('nombre', event.target.value)} /></label>
        <label>Día<select value={form.dia} onChange={(event) => cambiar('dia', event.target.value)}>{diasTorneo(torneo).map((dia) => <option key={dia} value={dia}>{dia}</option>)}</select></label>
        <label>Turno<select value={form.turno} onChange={(event) => cambiar('turno', event.target.value)}>{(torneo.turnos || ['Único']).map((turno) => <option key={turno} value={turno}>{turno}</option>)}</select></label>
        <label>Orden de categoría<input type="number" min="1" value={form.orden} onChange={(event) => cambiar('orden', event.target.value)} /></label>
      </div>
      <button className="ok-button" onClick={guardar}>{ok ? 'OK ✓' : 'OK'}</button>
    </div>
  )
}

function ConfigPuntaje({ data, mutate }) {
  const [nombre, setNombre] = useState('')
  const [ok, setOk] = useState({})

  function actualizar(conceptoId, valor) {
    mutate((draft) => {
      const concepto = draft.conceptosPuntaje.find((item) => item.id === conceptoId)
      concepto.nombre = valor
    })
  }

  function agregar() {
    if (!nombre.trim()) return
    mutate((draft) => {
      draft.conceptosPuntaje.push({ id: slug(nombre), nombre: nombre.trim() })
    }, `Concepto de puntaje agregado: ${nombre}`)
    setOk((prev) => ({ ...prev, nuevo: true }))
    window.setTimeout(() => setOk((prev) => ({ ...prev, nuevo: false })), 1300)
    setNombre('')
  }

  function quitar(conceptoId) {
    mutate((draft) => {
      draft.conceptosPuntaje = draft.conceptosPuntaje.filter((item) => item.id !== conceptoId)
    }, 'Concepto de puntaje eliminado')
  }

  return (
    <div className="score-config">
      <p>Preestablecé cuántos ítems se cargan y nombre de cada uno antes de competir.</p>
      {data.conceptosPuntaje.map((concepto) => (
        <div key={concepto.id}>
          <input value={concepto.nombre} onChange={(event) => actualizar(concepto.id, event.target.value)} />
          <button className="ok-button" onClick={() => {
            setOk((prev) => ({ ...prev, [concepto.id]: true }))
            window.setTimeout(() => setOk((prev) => ({ ...prev, [concepto.id]: false })), 1300)
          }}>{ok[concepto.id] ? 'OK ✓' : 'OK'}</button>
          <button className="danger" onClick={() => quitar(concepto.id)}>Quitar</button>
        </div>
      ))}
      <div>
        <input placeholder="Nuevo concepto: elemento técnico, componente..." value={nombre} onChange={(event) => setNombre(event.target.value)} />
        <button className="ok-button" onClick={agregar}>{ok.nuevo ? 'OK ✓' : 'OK'}</button>
      </div>
    </div>
  )
}

function EditorEntidades({ placeholder, ayudaArchivo, onFile, onAdd }) {
  const [nuevo, setNuevo] = useState('')
  const [file, setFile] = useState(null)
  const [ok, setOk] = useState(false)

  function agregar() {
    if (!nuevo.trim()) return
    const itemId = id('manual')
    onAdd(nuevo.trim(), itemId)
    if (file) onFile(itemId, file)
    setOk(true)
    window.setTimeout(() => setOk(false), 1300)
    setNuevo('')
    setFile(null)
  }

  return (
    <div className="entity-editor">
      <div className="entity-add manual-only">
        <input placeholder={placeholder} value={nuevo} onChange={(event) => setNuevo(event.target.value)} />
        <label>{ayudaArchivo}<input type="file" accept="image/*" onChange={(event) => setFile(event.target.files[0])} /></label>
        <button className="ok-button" onClick={agregar}>{ok ? 'OK ✓' : 'OK'}</button>
      </div>
      {ok && <small>Dato confirmado</small>}
    </div>
  )
}

function EditorPatinadoras({ data, mutate }) {
  const [form, setForm] = useState({ nombre: '', categoriaId: data.categorias[0]?.id || '', clubId: data.clubes[0]?.id || '', dia: data.torneos[0]?.fechaDesde || '', turno: data.torneos[0]?.turnos?.[0] || 'Único', ordenSalida: 1, musica: '' })
  const [ok, setOk] = useState(false)
  const torneo = data.torneos[0]
  const dias = diasTorneo(torneo)
  const turnos = torneo.turnos || ['Único']

  function cambiar(campo, valor) {
    setForm((prev) => ({ ...prev, [campo]: valor }))
  }

  function agregar() {
    if (!form.nombre.trim()) return
    const itemId = id('pat')
    mutate((draft) => {
      draft.patinadoras.push({
        id: itemId,
        orden: Number(form.ordenSalida) || 1,
        ordenSalida: Number(form.ordenSalida) || 1,
        dia: form.dia,
        turno: form.turno,
        estado: 'pendiente',
        nombre: form.nombre.trim(),
        edad: '',
        clubId: form.clubId,
        tecnicaId: draft.tecnicas[0]?.id || '',
        categoriaId: form.categoriaId,
        musicaId: slug(form.musica || form.nombre),
        musica: form.musica || form.nombre.trim(),
        audio: '',
        foto: '',
      })
    }, `Patinadora agregada: ${form.nombre}`)
    setOk(true)
    window.setTimeout(() => setOk(false), 1300)
    setForm((prev) => ({ ...prev, nombre: '', musica: '', ordenSalida: Number(prev.ordenSalida) + 1 }))
  }

  return (
    <div className="skater-editor">
      <div className="skater-row manual-only">
        <label>Nombre<input value={form.nombre} onChange={(event) => cambiar('nombre', event.target.value)} /></label>
        <label>Categoría<select value={form.categoriaId} onChange={(event) => cambiar('categoriaId', event.target.value)}>{[...data.categorias].sort(compararCategoria).map((categoria) => <option key={categoria.id} value={categoria.id}>{categoria.nombre}</option>)}</select></label>
        <label>Club<select value={form.clubId} onChange={(event) => cambiar('clubId', event.target.value)}>{data.clubes.map((club) => <option key={club.id} value={club.id}>{club.nombre}</option>)}</select></label>
        <label>Día<select value={form.dia} onChange={(event) => cambiar('dia', event.target.value)}>{dias.map((dia) => <option key={dia} value={dia}>{dia}</option>)}</select></label>
        <label>Turno<select value={form.turno} onChange={(event) => cambiar('turno', event.target.value)}>{turnos.map((turno) => <option key={turno} value={turno}>{turno}</option>)}</select></label>
        <label>Ubicación de salida<input type="number" min="1" value={form.ordenSalida} onChange={(event) => cambiar('ordenSalida', event.target.value)} /></label>
        <label>Música<input value={form.musica} onChange={(event) => cambiar('musica', event.target.value)} /></label>
        <button className="ok-button" onClick={agregar}>{ok ? 'OK ✓' : 'OK'}</button>
      </div>
      {ok && <small>Patinadora confirmada</small>}
    </div>
  )
}

function Reportes({ data }) {
  const [campos, setCampos] = useState({
    musica: false,
    profesora: false,
    puntaje: false,
  })
  const opciones = [
    ['categoria', 'Categoría'],
    ['dia', 'Día'],
    ['musica', 'Música'],
    ['profesora', 'Profesora'],
    ['puntaje', 'Puntuación'],
  ]
  const rows = [...data.patinadoras]
    .sort((a, b) => compararReporte(data, a, b))
  const grupos = rows.reduce((acc, row) => {
    const key = `${row.dia}|${row.turno}|${row.categoriaId}`
    if (!acc[key]) acc[key] = { dia: row.dia, turno: row.turno, categoria: nombreCategoria(data, row.categoriaId), rows: [] }
    acc[key].rows.push(row)
    return acc
  }, {})

  function valorCampo(row, campo) {
    if (campo === 'club') return nombreClub(data, row.clubId)
    if (campo === 'categoria') return nombreCategoria(data, row.categoriaId)
    if (campo === 'orden') return row.ordenSalida || row.orden || ''
    if (campo === 'musica') return row.musica || row.musicaId || ''
    if (campo === 'profesora') return data.tecnicas.find((tecnica) => tecnica.id === row.tecnicaId)?.nombre || ''
    if (campo === 'puntaje') return ''
    return row[campo] || ''
  }

  const activos = opciones.filter(([campo]) => campos[campo])
  const columnas = [
    ['orden', 'Orden de ingreso'],
    ['nombre', 'Nombre'],
    ['club', 'Club'],
    ...activos,
  ]

  return (
    <main className="print">
      <Panel title="Reporte de patinadoras">
        <div className="report-controls no-print">
          {opciones.map(([campo, label]) => (
            <label key={campo} className="check-line"><input type="checkbox" checked={campos[campo]} onChange={(event) => setCampos((prev) => ({ ...prev, [campo]: event.target.checked }))} />{label}</label>
          ))}
          <button onClick={() => window.print()}>Imprimir reporte</button>
        </div>
        <div className="report-groups">
          {Object.values(grupos).map((grupo) => (
            <section key={`${grupo.dia}-${grupo.turno}-${grupo.categoria}`} className="report-group">
              <header>
                <span>{grupo.dia}</span>
                <span>{grupo.turno}</span>
                <strong>{grupo.categoria}</strong>
              </header>
              <table>
                <thead><tr>{columnas.map(([, label]) => <th key={label}>{label}</th>)}</tr></thead>
                <tbody>
                  {grupo.rows.map((row) => (
                    <tr key={row.id}>{columnas.map(([campo]) => <td key={campo}>{valorCampo(row, campo)}</td>)}</tr>
                  ))}
                </tbody>
              </table>
            </section>
          ))}
        </div>
      </Panel>
    </main>
  )
}

function EditorJueces({ mutate }) {
  const [form, setForm] = useState({ nombre: '', rol: 'General' })
  const [ok, setOk] = useState(false)
  function agregar() {
    if (!form.nombre.trim()) return
    mutate((draft) => {
      draft.jueces.push({ id: id('juez'), nombre: form.nombre.trim(), rol: form.rol || 'General' })
    }, `Juez agregado: ${form.nombre}`)
    setOk(true)
    window.setTimeout(() => setOk(false), 1300)
    setForm({ nombre: '', rol: 'General' })
  }
  return (
    <div className="entity-add manual-only">
      <input placeholder="Nombre del juez" value={form.nombre} onChange={(event) => setForm((prev) => ({ ...prev, nombre: event.target.value }))} />
      <input placeholder="Rol" value={form.rol} onChange={(event) => setForm((prev) => ({ ...prev, rol: event.target.value }))} />
      <button className="ok-button" onClick={agregar}>{ok ? 'OK ✓' : 'OK'}</button>
    </div>
  )
}

function ConfigBuffet({ data, mutate }) {
  const [form, setForm] = useState({ producto: '', precio: '' })
  const [ok, setOk] = useState(false)
  function agregar() {
    if (!form.producto.trim()) return
    mutate((draft) => {
      draft.buffet.push({ id: id('buffet'), producto: form.producto.trim(), precio: Number(form.precio) || 0 })
    }, `Buffet agregado: ${form.producto}`)
    setOk(true)
    window.setTimeout(() => setOk(false), 1300)
    setForm({ producto: '', precio: '' })
  }
  return (
    <div className="buffet-config">
      <div className="entity-add manual-only">
        <input placeholder="Producto" value={form.producto} onChange={(event) => setForm((prev) => ({ ...prev, producto: event.target.value }))} />
        <input type="number" min="0" placeholder="Valor" value={form.precio} onChange={(event) => setForm((prev) => ({ ...prev, precio: event.target.value }))} />
        <button className="ok-button" onClick={agregar}>{ok ? 'OK ✓' : 'OK'}</button>
      </div>
      <div className="mini-list">{data.buffet.map((item) => <div key={item.id}><span>{item.producto}</span><strong>${Number(item.precio || 0).toLocaleString('es-AR')}</strong></div>)}</div>
    </div>
  )
}

function ConfigUsuarios({ data, mutate }) {
  const [nombre, setNombre] = useState('')
  function toggle(usuarioId, tabName, checked) {
    mutate((draft) => {
      const usuario = draft.usuarios.find((item) => item.id === usuarioId)
      usuario.tabs = checked ? [...new Set([...usuario.tabs, tabName])] : usuario.tabs.filter((item) => item !== tabName)
    }, 'Permisos de usuario actualizados')
  }
  function agregar() {
    if (!nombre.trim()) return
    mutate((draft) => {
      draft.usuarios.push({ id: id('user'), nombre: nombre.trim(), tabs: ['Operador'] })
    }, `Usuario agregado: ${nombre}`)
    setNombre('')
  }
  return (
    <div className="user-config">
      {data.usuarios.map((usuario) => (
        <div key={usuario.id}>
          <strong>{usuario.nombre}</strong>
          <div>
            {tabs.map((tabName) => <label className="check-line" key={tabName}><input type="checkbox" checked={usuario.tabs.includes(tabName)} onChange={(event) => toggle(usuario.id, tabName, event.target.checked)} />{tabName}</label>)}
          </div>
        </div>
      ))}
      <div className="entity-add">
        <input placeholder="Nuevo usuario" value={nombre} onChange={(event) => setNombre(event.target.value)} />
        <button onClick={agregar}>OK</button>
      </div>
    </div>
  )
}

function WebPublica({ data, modo, pista }) {
  const url = `${window.location.origin}${window.location.pathname}?vista=web`
  const qr = `https://api.qrserver.com/v1/create-qr-code/?size=260x260&data=${encodeURIComponent(url)}`

  return (
    <main className="grid two">
      <Panel title="QR para tanteador web">
        <div className="qr-box">
          <img src={qr} alt="QR tanteador web" />
          <div>
            <strong>Escanear para seguir el tanteador</strong>
            <p>Desde la web se elige la categoría y se ven posiciones, parciales y puntaje final.</p>
            <input readOnly value={url} onFocus={(event) => event.currentTarget.select()} />
          </div>
        </div>
      </Panel>
      <Panel title="Vista previa web">
        <VistaWeb data={data} modo={modo} pista={pista} compacto />
      </Panel>
    </main>
  )
}

function VistaWeb({ data, modo, pista, compacto = false }) {
  const categorias = [...data.categorias].sort(compararCategoria)
  const [categoriaId, setCategoriaId] = useState(categorias[0]?.id || '')
  const [vista, setVista] = useState('tanteador')
  const categoriaPantalla = data.categorias.find((item) => item.id === pista?.categoriaId)
  const categoria = categorias.find((item) => item.id === categoriaId) || categoriaPantalla || categorias[0]
  const rows = rankingCategoria(data, categoria?.id || '', modo).filter((row) => row.confirmado && row.total != null)
  const patinadorasPantalla = data.patinadoras
    .filter((item) => item.categoriaId === pista?.categoriaId && item.estado !== 'ausente')
    .sort(compararSalida)

  return (
    <main className={compacto ? 'web-score compact' : 'web-score'}>
      <header>
        <div className="web-actions">
          <button className={vista === 'tanteador' ? 'active' : ''} onClick={() => setVista('tanteador')}>Tanteador</button>
          <button className={vista === 'pantalla' ? 'active' : ''} onClick={() => setVista('pantalla')}>Pantalla</button>
          <button className={vista === 'buffet' ? 'active' : ''} onClick={() => setVista('buffet')}>Buffet</button>
          {!compacto && <button onClick={() => document.querySelector('.web-score')?.requestFullscreen?.()}>Maximizar</button>}
        </div>
        <span>Tanteador online</span>
        <h1>{data.torneos[0]?.nombre}</h1>
        <label>Categoría<select value={categoria?.id || ''} onChange={(event) => setCategoriaId(event.target.value)}>{categorias.map((item) => <option key={item.id} value={item.id}>{item.nombre}</option>)}</select></label>
      </header>
      {vista === 'buffet' && (
        <section className="buffet-web">
          <h2>Buffet</h2>
          <div>{data.buffet.map((item) => <article key={item.id}><span>{item.producto}</span><strong>${Number(item.precio || 0).toLocaleString('es-AR')}</strong></article>)}</div>
        </section>
      )}
      {vista === 'pantalla' && (
        <PantallaWeb data={data} pista={pista} categoria={categoriaPantalla || categoria} patinadoras={patinadorasPantalla} />
      )}
      {vista === 'tanteador' && <section>
        <h2>{categoria?.nombre}</h2>
        {rows.length ? (
          <div className="web-list">
            {rows.map((row) => {
              const parciales = parcialesPuntaje(data.puntajes[row.id], modo, data.conceptosPuntaje).filter((item) => item.valor != null)
              return (
                <article key={row.id}>
                  <strong>{row.puesto}</strong>
                  <div>
                    <h3>{row.nombre}</h3>
                    <p>{nombreClub(data, row.clubId)}</p>
                    <small>{parciales.map((item) => `${item.nombre}: ${item.valor.toFixed(2)}`).join(' · ')}</small>
                  </div>
                  <b>{row.total.toFixed(2)}</b>
                </article>
              )
            })}
          </div>
        ) : <div className="empty">Todavía no hay puntajes publicados en esta categoría.</div>}
      </section>}
    </main>
  )
}

function PantallaWeb({ data, pista, categoria, patinadoras }) {
  const [now, setNow] = useState(() => Date.now())
  useEffect(() => {
    const timer = window.setInterval(() => setNow(Date.now()), 1000)
    return () => window.clearInterval(timer)
  }, [])

  const inicioProgramadoHasta = pista?.inicioProgramadoHasta || data.torneos[0]?.inicioProgramadoHasta
  const enInicioProgramado = !data.torneos[0]?.inicioTorneoEn && inicioProgramadoHasta && now < inicioProgramadoHasta
  const enPrueba = pista?.pruebaPistaHasta && now < pista.pruebaPistaHasta
  const enRecesoManual = pista?.estado === 'Receso' && pista?.recesoHasta && now < pista.recesoHasta
  const enRecesoAuto = pista?.categoriaFinalHasta && pista?.autoRecesoHasta && now >= pista.categoriaFinalHasta && now < pista.autoRecesoHasta
  const hasta = enInicioProgramado ? inicioProgramadoHasta : enPrueba ? pista.pruebaPistaHasta : enRecesoManual ? pista.recesoHasta : enRecesoAuto ? pista.autoRecesoHasta : null
  const estado = enInicioProgramado ? 'Inicio de torneo' : enPrueba ? 'Probando pista' : (enRecesoManual || enRecesoAuto) ? 'Receso' : pista?.estado || 'Pantalla pública'
  const siguiente = siguienteCategoria(data, pista?.categoriaId)

  return (
    <section className="screen-web">
      <h2>{hasta ? estado : 'En pista'}</h2>
      <p>{data.torneos[0]?.nombre}</p>
      {hasta ? (
        <>
          <strong>{categoria?.nombre || '-'}</strong>
          <div className="web-timer"><Timer hasta={hasta} /><span>Cuenta regresiva</span></div>
        </>
      ) : (
        <div className="web-current-next">
          <article><span>Categoría actual</span><strong>{categoria?.nombre || '-'}</strong></article>
          <article><span>Siguiente categoría</span><strong>{siguiente?.nombre || 'A confirmar'}</strong></article>
        </div>
      )}
      {enPrueba && (
        <div className="web-skater-list">
          <b>Patinadoras</b>
          {patinadoras.map((item) => <span key={item.id}>{item.ordenSalida || item.orden}. {item.nombre}</span>)}
        </div>
      )}
    </section>
  )
}

function Tanteador({ data, categoriaId, modo }) {
  const categoria = data.categorias.find((item) => item.id === categoriaId)
  const rows = rankingCategoria(data, categoriaId, modo).filter((row) => row.confirmado && row.total != null)
  return (
    <main>
      <Panel title={`Tanteador: ${categoria?.nombre || ''}`}>
        <Tabla rows={rows} clubes={data.clubes} />
      </Panel>
    </main>
  )
}

function RankingClubes({ data, modo }) {
  return (
    <main>
      <Panel title="Ranking clubes">
        <table><thead><tr><th>Club</th><th>1</th><th>2</th><th>3</th><th>4</th><th>5</th><th>Puntos</th><th>Incentivos</th></tr></thead><tbody>{rankingClubes(data, modo).map((row) => <tr key={row.clubId}><td>{data.clubes.find((c) => c.id === row.clubId)?.nombre}</td><td>{row.primero}</td><td>{row.segundo}</td><td>{row.tercero}</td><td>{row.cuarto}</td><td>{row.quinto}</td><td>{row.puntos}</td><td>{row.incentivos}</td></tr>)}</tbody></table>
      </Panel>
    </main>
  )
}

function Rankings({ data, modo }) {
  return (
    <main className="grid two">
      <Panel title="Tanteadores por categoría">{[...data.categorias].sort(compararCategoria).map((categoria) => <Tabla key={categoria.id} titulo={categoria.nombre} rows={rankingCategoria(data, categoria.id, modo).filter((row) => row.confirmado && row.total != null)} clubes={data.clubes} />)}</Panel>
      <RankingClubes data={data} modo={modo} />
    </main>
  )
}

function Actas({ data, modo }) {
  return (
    <main className="print">
      <button className="no-print" onClick={() => window.print()}>Imprimir actas</button>
      {[...data.categorias].sort(compararCategoria).map((categoria) => <article className="sheet" key={categoria.id}><h2>Acta categoría: {categoria.nombre}</h2><Tabla rows={rankingCategoria(data, categoria.id, modo)} clubes={data.clubes} /><Firmas jueces={data.jueces} /></article>)}
      <article className="sheet"><h2>Premiación final por clubes</h2><Rankings data={data} modo={modo} /></article>
    </main>
  )
}

function Logs({ data }) {
  const [torneoId, setTorneoId] = useState(data.torneos[0]?.id || '')
  const logs = data.logs.filter((log) => !torneoId || !log.torneoId || log.torneoId === torneoId)

  function descargar() {
    const torneo = data.torneos.find((item) => item.id === torneoId)
    const csv = [
      ['Fecha', 'Torneo', 'Registro'],
      ...logs.map((log) => [new Date(log.fecha).toLocaleString('es-AR'), torneo?.nombre || '', log.texto]),
    ].map((row) => row.map((cell) => `"${String(cell || '').replaceAll('"', '""')}"`).join(',')).join('\n')
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' })
    const link = document.createElement('a')
    link.href = URL.createObjectURL(blob)
    link.download = `logs-${slug(torneo?.nombre || 'todos')}.csv`
    link.click()
    URL.revokeObjectURL(link.href)
  }

  return (
    <main>
      <Panel title="Registros básicos">
        <div className="log-toolbar">
          <label>Torneo<select value={torneoId} onChange={(event) => setTorneoId(event.target.value)}><option value="">Todos</option>{data.torneos.map((torneo) => <option key={torneo.id} value={torneo.id}>{torneo.nombre}</option>)}</select></label>
          <button onClick={descargar}>Bajar logs</button>
        </div>
        <div className="logs">{logs.map((log) => <div key={log.id}><time>{new Date(log.fecha).toLocaleString()}</time><span>{log.texto}</span></div>)}</div>
      </Panel>
    </main>
  )
}

function FlujoCategoria({ data, pista, actual, mostrarProxima = true }) {
  const actualCategoria = data.categorias.find((categoria) => categoria.id === pista.categoriaId)
  const siguiente = siguienteCategoria(data, pista.categoriaId)
  const patinadorasActuales = data.patinadoras
    .filter((item) => item.categoriaId === pista.categoriaId && item.estado !== 'ausente')
    .sort(compararSalida)
  const patinadorasSiguientes = data.patinadoras
    .filter((item) => item.categoriaId === siguiente?.id && item.estado !== 'ausente')
    .sort(compararSalida)
  const indiceActual = patinadorasActuales.findIndex((item) => item.id === actual.patinadora?.id)
  const proxima = patinadorasActuales[indiceActual + 1]
  const proximaClub = data.clubes.find((club) => club.id === proxima?.clubId)
  const mismoTurnoSiguiente = siguiente && actualCategoria && siguiente.dia === actualCategoria.dia && siguiente.turno === actualCategoria.turno

  function estadoFila(item) {
    if (item.id === actual.patinadora?.id) return 'En pista'
    if (item.estado === 'finalizada') return 'Ya patinó'
    return 'Pendiente'
  }

  return (
    <section className={`flow-board ${mostrarProxima ? '' : 'flow-board-no-next'}`}>
      {mostrarProxima && <article className="next-card">
        <span>Siguiente patinadora</span>
        <strong>{proxima?.nombre || 'Última patinadora de la categoría'}</strong>
        <small>{proxima ? `${proximaClub?.nombre || '-'} · salida ${proxima.ordenSalida || proxima.orden}` : actualCategoria?.nombre}</small>
      </article>}
      <article>
        <h3>Compitiendo</h3>
        <p>{actualCategoria?.nombre}</p>
        <div>
          {patinadorasActuales.map((item) => (
            <div className={item.estado === 'finalizada' ? 'done' : item.id === actual.patinadora?.id ? 'current' : ''} key={item.id}>
              <b>{item.ordenSalida || item.orden}</b>
              <span>{item.nombre}</span>
              <small>{nombreClub(data, item.clubId)} · {estadoFila(item)}</small>
            </div>
          ))}
        </div>
      </article>
      <article>
        <h3>{mismoTurnoSiguiente ? 'Siguiente categoría' : siguiente ? 'Turno finalizado' : 'Torneo finalizado'}</h3>
        <p>{mismoTurnoSiguiente ? siguiente?.nombre : siguiente ? `${actualCategoria?.dia || ''} ${actualCategoria?.turno || ''}` : 'Sin más categorías'}</p>
        <div>
          {mismoTurnoSiguiente && patinadorasSiguientes.length ? patinadorasSiguientes.map((item) => (
            <div key={item.id}>
              <b>{item.ordenSalida || item.orden}</b>
              <span>{item.nombre}</span>
              <small>{nombreClub(data, item.clubId)}</small>
            </div>
          )) : <div className="operator-continues"><span>{mismoTurnoSiguiente ? 'Sin listado cargado' : 'El operador continúa con el próximo turno.'}</span></div>}
        </div>
      </article>
    </section>
  )
}

function ConfigTema({ data, mutate }) {
  return (
    <div className="theme-picker">
      {temasSistema.map((tema) => (
        <button key={tema.id} className={data.tema === tema.id ? 'active' : ''} onClick={() => mutate((draft) => { draft.tema = tema.id }, `Tema seleccionado: ${tema.nombre}`)}>
          <span className={`theme-swatch ${tema.id}`}></span>
          {tema.nombre}
        </button>
      ))}
    </div>
  )
}

function FullscreenButton({ targetSelector }) {
  function abrirPantalla() {
    const target = document.querySelector(targetSelector) || document.documentElement
    target.requestFullscreen?.()
  }
  return <button className="fullscreen-button" onClick={abrirPantalla}>Maximizar pantalla</button>
}

function reproducirSilbato() {
  const audio = new Audio(WHISTLE_SRC)
  audio.currentTime = 0
  audio.play().catch(() => {})
}

function Competidor({ actual, total }) {
  return (
    <section className="competidor">
      <Avatar src={actual.club?.logo} label={actual.club?.nombre || actual.patinadora?.nombre} color={actual.club?.color} />
      <div><small>{actual.categoria?.nombre}</small><h2>{actual.patinadora?.nombre}</h2><p>{actual.club?.nombre}</p><b>Ubicación de salida {actual.patinadora?.ordenSalida || actual.patinadora?.orden || '-'}</b></div>
      <strong>{total == null ? '--' : total.toFixed(2)}</strong>
    </section>
  )
}

function ScorePad({ value, onChange }) {
  const values = ['5.0', '5.5', '6.0', '6.5', '7.0', '7.5', '8.0', '8.5', '9.0', '9.5', '10']
  return <><input className="score-input" type="number" min="0" max="10" step="0.1" value={value} onChange={(e) => onChange(e.target.value)} /><div className="pad">{values.map((item) => <button key={item} onClick={() => onChange(item)}>{item}</button>)}</div></>
}

function Tabla({ titulo, rows, clubes }) {
  return <div className="tabla">{titulo && <h3>{titulo}</h3>}<table><thead><tr><th>Pos</th><th>Nombre</th><th>Club</th><th>Total</th></tr></thead><tbody>{rows.map((row) => <tr key={row.id}><td>{row.puesto || '-'}</td><td>{row.nombre}</td><td>{clubes.find((club) => club.id === row.clubId)?.nombre}</td><td>{row.total == null ? '-' : row.total.toFixed(2)}</td></tr>)}</tbody></table></div>
}

function Avatar({ src, label, color = '#22c55e' }) {
  return src ? <img className="avatar" src={src} alt={label || ''} /> : <div className="avatar fallback" style={{ '--avatar': color }}>{(label || '?').slice(0, 2).toUpperCase()}</div>
}

function Panel({ title, children, className = '' }) {
  return <section className={`panel ${className}`}><h2>{title}</h2>{children}</section>
}

function StatusInfo({ clima }) {
  const [horaBase, setHoraBase] = useState({ date: new Date(), source: 'equipo' })
  const [now, setNow] = useState(new Date())
  const [weather, setWeather] = useState('')

  useEffect(() => {
    let active = true
    fetch('https://worldtimeapi.org/api/ip')
      .then((response) => response.ok ? response.json() : null)
      .then((json) => {
        if (!active || !json?.datetime) return
        setHoraBase({ date: new Date(json.datetime), source: 'internet' })
        setNow(new Date(json.datetime))
      })
      .catch(() => {
        if (active) setHoraBase({ date: new Date(), source: 'equipo' })
      })
    return () => {
      active = false
    }
  }, [])

  useEffect(() => {
    const startedAt = Date.now()
    const timer = setInterval(() => {
      setNow(new Date(horaBase.date.getTime() + (Date.now() - startedAt)))
    }, 1000)
    return () => clearInterval(timer)
  }, [horaBase])

  useEffect(() => {
    if (!clima || !navigator.geolocation) {
      return
    }
    navigator.geolocation.getCurrentPosition((position) => {
      const { latitude, longitude } = position.coords
      fetch(`https://api.open-meteo.com/v1/forecast?latitude=${latitude}&longitude=${longitude}&current=temperature_2m,weather_code,wind_speed_10m`)
        .then((response) => response.ok ? response.json() : null)
        .then((json) => {
          const current = json?.current
          setWeather(current ? `${Math.round(current.temperature_2m)}°C - viento ${Math.round(current.wind_speed_10m)} km/h` : '')
        })
        .catch(() => setWeather(''))
    }, () => setWeather(''), { timeout: 5000 })
  }, [clima])

  return (
    <aside className="status-info">
      <strong>{now.toLocaleTimeString('es-AR', { hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false })}</strong>
      {clima && weather && <small>{weather}</small>}
    </aside>
  )
}

function Timer({ hasta }) {
  const [now, setNow] = useState(() => Date.now())
  const sonoRef = useRef(false)
  useEffect(() => {
    const timer = setInterval(() => setNow(Date.now()), 500)
    return () => clearInterval(timer)
  }, [])
  useEffect(() => {
    sonoRef.current = false
  }, [hasta])
  const left = Math.max(0, hasta - now)
  useEffect(() => {
    if (left <= 0 && hasta && !sonoRef.current) {
      sonoRef.current = true
      reproducirSilbato()
    }
  }, [hasta, left])
  const minutes = Math.floor(left / 60000)
  const seconds = Math.floor((left % 60000) / 1000)
  return <b>{String(minutes).padStart(2, '0')}:{String(seconds).padStart(2, '0')}</b>
}

function Firmas({ jueces }) {
  return <div className="firmas">{jueces.map((juez) => <div key={juez.id}><span></span><p>{juez.nombre}</p></div>)}</div>
}
