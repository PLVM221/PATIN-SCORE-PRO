import { useEffect, useMemo, useRef, useState } from 'react'
import * as XLSX from 'xlsx'
import './App.css'

const STORAGE_KEY = 'patin-score-pro-v1'

const demo = {
  torneos: [{ id: 't1', nombre: 'Copa Ciudad 2026', sede: 'Polideportivo Central', fecha: '2026-05-26' }],
  categorias: [
    { id: 'c1', nombre: 'Libre Infantil B', torneoId: 't1' },
    { id: 'c2', nombre: 'Escuela Cadete C', torneoId: 't1' },
    { id: 'c3', nombre: 'Show Juvenil', torneoId: 't1' },
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
  patinadoras: [
    { id: 'p1', orden: 1, estado: 'pendiente', nombre: 'Sofia Benitez', edad: 10, clubId: 'cl1', tecnicaId: 'te1', categoriaId: 'c1', musicaId: 'libertango', musica: 'Libertango', audio: '', foto: '' },
    { id: 'p2', orden: 2, estado: 'pendiente', nombre: 'Martina Lagos', edad: 11, clubId: 'cl2', tecnicaId: 'te2', categoriaId: 'c1', musicaId: 'cinema-paradiso', musica: 'Cinema Paradiso', audio: '', foto: '' },
    { id: 'p3', orden: 1, estado: 'pendiente', nombre: 'Camila Rios', edad: 14, clubId: 'cl3', tecnicaId: 'te3', categoriaId: 'c2', musicaId: 'experience', musica: 'Experience', audio: '', foto: '' },
    { id: 'p4', orden: 1, estado: 'pendiente', nombre: 'Equipo Aurora', edad: 15, clubId: 'cl1', tecnicaId: 'te1', categoriaId: 'c3', musicaId: 'heroes', musica: 'Heroes', audio: '', foto: '' },
  ],
  puntajes: {},
  logs: [{ id: 'l1', fecha: new Date().toISOString(), texto: 'Demo iniciado' }],
}

const basePista = {
  torneoId: 't1',
  categoriaId: 'c1',
  patinadoraId: 'p1',
  estado: 'Preparando',
  modo: 'jueces',
  recesoMin: 5,
  recesoHasta: null,
  puntajeHasta: null,
  puntajePatinadoraId: null,
  categoriaFinalHasta: null,
  autoRecesoHasta: null,
}

const tabs = ['Operador', 'Jueces', 'Pública LED', 'Datos', 'Tanteador', 'Ranking clubes', 'Actas', 'Registros']

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
    canciones: data.canciones || [],
    conceptosPuntaje: data.conceptosPuntaje || demo.conceptosPuntaje,
    patinadoras: data.patinadoras.map((item, index) => ({
      orden: index + 1,
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

function rankingCategoria(data, categoriaId, modo) {
  return data.patinadoras
    .filter((patinadora) => patinadora.categoriaId === categoriaId && patinadora.estado !== 'ausente')
    .map((patinadora) => ({
      ...patinadora,
      total: totalPuntaje(data.puntajes[patinadora.id], modo, data.conceptosPuntaje),
      confirmado: Boolean(data.puntajes[patinadora.id]?.confirmado),
    }))
    .sort((a, b) => (b.total ?? -1) - (a.total ?? -1) || (a.orden || 0) - (b.orden || 0))
    .map((row, index) => ({ ...row, puesto: row.total == null ? null : index + 1 }))
}

function rankingClubes(data, modo) {
  const tabla = new Map()
  data.categorias.forEach((categoria) => {
    rankingCategoria(data, categoria.id, modo)
      .filter((row) => row.puesto)
      .forEach((row) => {
        const actual = tabla.get(row.clubId) || { clubId: row.clubId, primero: 0, segundo: 0, tercero: 0, podios: 0, incentivos: 0 }
        if (row.puesto === 1) actual.primero += 1
        if (row.puesto === 2) actual.segundo += 1
        if (row.puesto === 3) actual.tercero += 1
        if (row.puesto <= 3) actual.podios += 1
        if ((row.total || 0) >= 7.5) actual.incentivos += 1
        tabla.set(row.clubId, actual)
      })
  })
  return [...tabla.values()].sort((a, b) => b.primero - a.primero || b.segundo - a.segundo || b.tercero - a.tercero || b.podios - a.podios || b.incentivos - a.incentivos)
}

function cell(row, names) {
  const entries = Object.entries(row)
  for (const name of names) {
    const found = entries.find(([key]) => slug(key) === slug(name))
    if (found) return String(found[1] || '').trim()
  }
  return ''
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
  const audioRef = useRef(null)

  useEffect(() => localStorage.setItem(STORAGE_KEY, JSON.stringify(data)), [data])

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
    .sort((a, b) => (a.orden || 0) - (b.orden || 0))
  const totalActual = totalPuntaje(actual.puntaje, pista.modo, data.conceptosPuntaje)

  function log(texto) {
    setData((prev) => ({ ...prev, logs: [{ id: id('log'), fecha: new Date().toISOString(), texto }, ...prev.logs] }))
  }

  function mutate(fn, texto) {
    setData((prev) => {
      const next = structuredClone(prev)
      fn(next)
      if (texto) next.logs.unshift({ id: id('log'), fecha: new Date().toISOString(), texto })
      return next
    })
  }

  function cambiarCategoria(categoriaId) {
    const primera = data.patinadoras
      .filter((item) => item.categoriaId === categoriaId && item.estado !== 'ausente')
      .sort((a, b) => (a.orden || 0) - (b.orden || 0))[0]
    setPista((prev) => ({ ...prev, categoriaId, patinadoraId: primera?.id || '', estado: 'Preparando' }))
  }

  function mover(delta) {
    const index = patinadorasCategoria.findIndex((item) => item.id === pista.patinadoraId)
    const next = patinadorasCategoria[index + delta]
    if (next) setPista((prev) => ({ ...prev, patinadoraId: next.id, estado: 'Preparando' }))
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
    const AudioContext = window.AudioContext || window.webkitAudioContext
    const ctx = new AudioContext()
    const osc = ctx.createOscillator()
    const gain = ctx.createGain()
    osc.type = 'square'
    osc.frequency.value = 1650
    gain.gain.setValueAtTime(0.001, ctx.currentTime)
    gain.gain.exponentialRampToValueAtTime(0.45, ctx.currentTime + 0.03)
    gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.55)
    osc.connect(gain)
    gain.connect(ctx.destination)
    osc.start()
    osc.stop(ctx.currentTime + 0.6)
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
    const categoriaFinalHasta = siguiente ? null : now + 60000
    setPista((prev) => ({
      ...prev,
      patinadoraId: siguiente?.id || actual.patinadora.id,
      estado: siguiente ? 'Preparando' : 'Categoría finalizada',
      puntajeHasta: now + 10000,
      puntajePatinadoraId: actual.patinadora.id,
      categoriaFinalHasta,
      autoRecesoHasta: categoriaFinalHasta ? categoriaFinalHasta + Number(prev.recesoMin) * 60000 : null,
    }))
  }

  function pasarSiguienteCategoria() {
    const actualIndex = data.categorias.findIndex((item) => item.id === pista.categoriaId)
    const siguienteCategoria = data.categorias[actualIndex + 1]
    if (!siguienteCategoria) return
    const primera = data.patinadoras
      .filter((item) => item.categoriaId === siguienteCategoria.id && item.estado !== 'ausente')
      .sort((a, b) => (a.orden || 0) - (b.orden || 0))[0]
    setPista((prev) => ({
      ...prev,
      categoriaId: siguienteCategoria.id,
      patinadoraId: primera?.id || '',
      estado: 'Preparando',
      puntajeHasta: null,
      puntajePatinadoraId: null,
      categoriaFinalHasta: null,
      autoRecesoHasta: null,
    }))
  }

  function iniciarRecesoDesdeJuez(minutos) {
    const duration = Math.max(1, Number(minutos) || 1)
    setPista((prev) => ({
      ...prev,
      recesoMin: duration,
      estado: 'Receso',
      recesoHasta: Date.now() + duration * 60000,
      categoriaFinalHasta: null,
      autoRecesoHasta: null,
    }))
    log(`Receso ${duration} min`)
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
    const maxOrden = Math.max(...data.patinadoras.filter((item) => item.categoriaId === pista.categoriaId).map((item) => item.orden || 0), 0)
    mutate((draft) => {
      const row = draft.patinadoras.find((item) => item.id === actual.patinadora.id)
      row.orden = maxOrden + 1
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
            torneo = { id: id('tor'), nombre: torneoNombre, sede: '', fecha: new Date().toISOString().slice(0, 10) }
            draft.torneos.push(torneo)
          }
          torneoId = torneo.id
        }
        const existente = draft.categorias.find((item) => slug(item.nombre) === slug(nombre) && item.torneoId === torneoId)
        if (!existente) draft.categorias.push({ id: id('cat'), nombre, torneoId })
      })
    }, `Categorías importadas: ${file.name}`)
  }

  return (
    <div className="app">
      <header className="top">
        <div>
          <span>Patín Score Pro</span>
          <h1>{actual.torneo?.nombre}</h1>
          <p>{actual.torneo?.sede} - {actual.torneo?.fecha}</p>
        </div>
        <strong>{pista.estado}</strong>
      </header>

      <nav className="tabs">
        {tabs.map((item) => <button className={tab === item ? 'active' : ''} key={item} onClick={() => setTab(item)}>{item}</button>)}
      </nav>

      {tab === 'Operador' && (
        <main className="grid two">
          <Panel title="Pista">
            <div className="fields">
              <label>Categoría<select value={pista.categoriaId} onChange={(e) => cambiarCategoria(e.target.value)}>{data.categorias.map((item) => <option key={item.id} value={item.id}>{item.nombre}</option>)}</select></label>
              <label>Patinadora<select value={pista.patinadoraId} onChange={(e) => setPista({ ...pista, patinadoraId: e.target.value })}>{patinadorasCategoria.map((item) => <option key={item.id} value={item.id}>{item.nombre}</option>)}</select></label>
              <label>Modo<select value={pista.modo} onChange={(e) => setPista({ ...pista, modo: e.target.value })}><option value="jueces">Varios jueces</option><option value="consenso">Consensuado</option></select></label>
              <label>Receso minutos<input type="number" min="1" value={pista.recesoMin} onChange={(e) => setPista({ ...pista, recesoMin: e.target.value })} /></label>
            </div>
            <Competidor actual={actual} total={totalActual} />
            <div className="actions">
              <button onClick={() => setPista({ ...pista, estado: 'En pista' })}>En pista</button>
              <button onClick={silbato}>Silbato</button>
              <button onClick={() => mover(-1)}>Anterior</button>
              <button onClick={() => mover(1)}>Siguiente</button>
              <button onClick={() => { setPista({ ...pista, estado: 'Receso', recesoHasta: Date.now() + Number(pista.recesoMin) * 60000 }); log(`Receso ${pista.recesoMin} min`) }}>Receso</button>
              <button className="danger" onClick={marcarAusente}>No se presentó</button>
              <button onClick={postergar}>Postergar al final</button>
            </div>
          </Panel>
          <Panel title="Música y confirmación">
            <div className="music-now">
              <small>Tema asignado</small>
              <strong>{actual.cancion?.nombre || actual.patinadora?.musica || 'Sin tema'}</strong>
              <span>ID: {actual.patinadora?.musicaId || '-'}</span>
            </div>
            {actual.cancion?.archivo || actual.patinadora?.audio ? <audio ref={audioRef} controls src={actual.cancion?.archivo || actual.patinadora.audio} /> : <div className="empty">El administrador carga audio con mismo ID de la patinadora.</div>}
            <div className="actions">
              <button className="primary" onClick={playActual}>Reproducir actual</button>
              <button onClick={() => audioRef.current?.pause()}>Pausa</button>
              <button className="danger" onClick={() => setPista({ ...pista, estado: 'Detenido' })}>Detener</button>
            </div>
            <div className="confirm">
              <b>{actual.puntaje?.confirmado ? 'Publicado por juez' : 'Pendiente de publicación del juez'}</b>
            </div>
          </Panel>
        </main>
      )}

      {tab === 'Jueces' && <Jueces data={data} pista={pista} actual={actual} juezId={juezId} setJuezId={setJuezId} actualizarPuntaje={actualizarPuntaje} publicarPuntaje={publicarPuntaje} pasarSiguienteCategoria={pasarSiguienteCategoria} iniciarReceso={iniciarRecesoDesdeJuez} />}
      {tab === 'Pública LED' && <Publica pista={pista} actual={actual} total={totalActual} data={data} modo={pista.modo} />}
      {tab === 'Datos' && <Datos data={data} mutate={mutate} guardarArchivo={guardarArchivo} importarListado={importarListado} cargarCanciones={cargarCanciones} importarClubes={importarClubes} importarTecnicas={importarTecnicas} importarCategorias={importarCategorias} />}
      {tab === 'Tanteador' && <Tanteador data={data} categoriaId={pista.categoriaId} modo={pista.modo} />}
      {tab === 'Ranking clubes' && <RankingClubes data={data} modo={pista.modo} />}
      {tab === 'Actas' && <Actas data={data} modo={pista.modo} />}
      {tab === 'Registros' && <Logs data={data} />}

      <footer>
        <button className="ghost danger" onClick={() => { localStorage.removeItem(STORAGE_KEY); setData(demo); setPista(basePista) }}>Reiniciar demo</button>
        <span>Persistencia localStorage. Sin SQL.</span>
      </footer>
    </div>
  )
}

function Jueces({ data, pista, actual, juezId, setJuezId, actualizarPuntaje, publicarPuntaje, pasarSiguienteCategoria, iniciarReceso }) {
  const puntaje = actual.puntaje || { jueces: {} }
  const [minutosReceso, setMinutosReceso] = useState(pista.recesoMin || 5)
  const miPuntaje = puntaje.jueces[juezId] || {}
  const totalTecnico = totalPuntaje(puntaje, pista.modo, data.conceptosPuntaje)
  const patinadorasOrdenadas = data.patinadoras
    .filter((item) => item.categoriaId === pista.categoriaId && item.estado !== 'ausente')
    .sort((a, b) => (a.orden || 0) - (b.orden || 0))
  const indiceActual = patinadorasOrdenadas.findIndex((item) => item.id === actual.patinadora?.id)
  const proxima = patinadorasOrdenadas[indiceActual + 1]
  const proximaClub = data.clubes.find((item) => item.id === proxima?.clubId)
  const pendientesPista = patinadorasOrdenadas.slice(Math.max(indiceActual + 1, 0))
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
        <div className="next-skater">
          <span>Se prepara</span>
          <strong>{proxima?.nombre || 'Última patinadora de la categoría'}</strong>
          {proxima && <small>{proximaClub?.nombre}</small>}
        </div>
        <div className="warmup-list">
          <span>Próximas para ingresar a pista</span>
          {pendientesPista.length ? pendientesPista.map((item) => (
            <div key={item.id}>{item.nombre}</div>
          )) : <div>Sin patinadoras pendientes en esta categoría</div>}
        </div>
        {!proxima && (
          <div className="category-end-actions">
            <button onClick={pasarSiguienteCategoria}>Pasar a siguiente categoría</button>
            <label>
              Minutos de receso
              <input type="number" min="1" value={minutosReceso} onChange={(event) => setMinutosReceso(event.target.value)} />
            </label>
            <button onClick={() => iniciarReceso(minutosReceso)}>Receso</button>
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

function Publica({ pista, actual, total, data, modo }) {
  const [now, setNow] = useState(() => Date.now())
  useEffect(() => {
    const timer = setInterval(() => setNow(Date.now()), 500)
    return () => clearInterval(timer)
  }, [])

  const mostrandoPuntaje = pista.puntajeHasta && now < pista.puntajeHasta && pista.puntajePatinadoraId
  const mostrandoRecesoAuto = pista.categoriaFinalHasta && pista.autoRecesoHasta && now >= pista.categoriaFinalHasta && now < pista.autoRecesoHasta
  const patinadoraPuntaje = data.patinadoras.find((item) => item.id === pista.puntajePatinadoraId)
  const clubPuntaje = data.clubes.find((item) => item.id === patinadoraPuntaje?.clubId)
  const totalPuntajeReciente = totalPuntaje(data.puntajes[pista.puntajePatinadoraId], modo, data.conceptosPuntaje)

  if (pista.estado === 'Receso' || mostrandoRecesoAuto) {
    return (
      <main className="led">
        <div>
          <p>{actual.categoria?.nombre}</p>
          <h2>Receso</h2>
          <span>{actual.torneo?.nombre}</span>
        </div>
        <strong>--</strong>
        {(pista.recesoHasta || pista.autoRecesoHasta) && <Timer hasta={pista.recesoHasta || pista.autoRecesoHasta} />}
      </main>
    )
  }

  if (mostrandoPuntaje) {
    return (
      <main className="led score-reveal">
        <div>
          <p>Puntaje recibido</p>
          <h2>{patinadoraPuntaje?.nombre}</h2>
          <span>{clubPuntaje?.nombre}</span>
        </div>
        <strong>{totalPuntajeReciente == null ? '--' : totalPuntajeReciente.toFixed(2)}</strong>
        <em>Mostrando puntaje</em>
      </main>
    )
  }

  const rows = rankingCategoria(data, pista.categoriaId, modo).filter((row) => row.confirmado && row.total != null)
  if (rows.some((row) => row.total != null)) {
    return <TanteadorLed categoria={actual.categoria} rows={rows} clubes={data.clubes} />
  }

  return (
    <main className="led">
      <div>
        <p>{actual.categoria?.nombre}</p>
        <h2>{actual.patinadora?.nombre}</h2>
        <span>{actual.club?.nombre}</span>
      </div>
      <strong>{total == null ? '--' : total.toFixed(2)}</strong>
      <em>{pista.estado}</em>
    </main>
  )
}

function TanteadorLed({ categoria, rows, clubes }) {
  return (
    <main className="led-board">
      <header>
        <span>Tanteador</span>
        <h2>{categoria?.nombre}</h2>
      </header>
      <div className="board-list">
        {rows.map((row) => (
          <div className={row.total == null ? 'pending' : ''} key={row.id}>
            <strong>{row.puesto || '-'}</strong>
            <span>{row.nombre}</span>
            <small>{clubes.find((club) => club.id === row.clubId)?.nombre}</small>
            <b>{row.total == null ? '--' : row.total.toFixed(2)}</b>
          </div>
        ))}
      </div>
    </main>
  )
}

function Datos({ data, mutate, guardarArchivo, importarListado, cargarCanciones, importarClubes, importarTecnicas, importarCategorias }) {
  const update = (entidad, itemId, campo, valor) => mutate((draft) => {
    draft[entidad].find((item) => item.id === itemId)[campo] = valor
  })
  return (
    <main className="grid admin-grid">
      <Panel title="Importar patinadoras">
        <div className="import-box">
          <p>Columnas: Nombre, Club, Categoría, Música. Opcional: Edad.</p>
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
          <p>Columnas: Categoría o Nombre. Opcional: Torneo.</p>
          <label>Excel o CSV<input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => importarCategorias(e.target.files[0])} /></label>
        </div>
      </Panel>
      <Panel title="Importar profesoras">
        <div className="import-box">
          <p>Columnas: Profesora o Técnica o Nombre. Opcional: Club.</p>
          <label>Excel o CSV<input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => importarTecnicas(e.target.files[0])} /></label>
        </div>
      </Panel>
      <Panel title="Cargar canciones">
        <div className="import-box">
          <p>Nombre de archivo identifica música. Ejemplo: libertango.mp3 coincide con MúsicaId libertango.</p>
          <label>Audios<input type="file" multiple accept="audio/*" onChange={(e) => cargarCanciones(e.target.files)} /></label>
          <div className="song-list">{data.canciones.map((item) => <div key={item.id}><b>{item.id}</b><span>{item.archivo ? 'cargada' : 'sin audio'}</span></div>)}</div>
        </div>
      </Panel>
      <Panel title="Configurar puntaje">
        <ConfigPuntaje data={data} mutate={mutate} />
      </Panel>
      <Panel title="Clubes">
        <EditorEntidades
          items={data.clubes}
          srcCampo="logo"
          ayudaArchivo="Subir logo del club"
          onName={(itemId, valor) => update('clubes', itemId, 'nombre', valor)}
          onFile={(itemId, file) => guardarArchivo('img', 'clubes', itemId, 'logo', file)}
          onAdd={(nombre) => mutate((draft) => { draft.clubes.push({ id: id('club'), nombre, color: '#28e67a', logo: '' }) }, `Club agregado: ${nombre}`)}
        />
      </Panel>
      <Panel title="Profesoras">
        <EditorEntidades
          items={data.tecnicas}
          srcCampo="foto"
          ayudaArchivo="Subir foto de profesora"
          onName={(itemId, valor) => update('tecnicas', itemId, 'nombre', valor)}
          onFile={(itemId, file) => guardarArchivo('img', 'tecnicas', itemId, 'foto', file)}
          onAdd={(nombre) => mutate((draft) => { draft.tecnicas.push({ id: id('tec'), nombre, clubId: data.clubes[0]?.id || '', foto: '' }) }, `Profesora agregada: ${nombre}`)}
        />
      </Panel>
      <Panel title="Patinadoras">
        <EditorEntidades
          items={data.patinadoras}
          srcCampo="foto"
          ayudaArchivo="Subir foto de patinadora"
          extra={(item) => `${item.estado} - ${item.musicaId}`}
          onName={(itemId, valor) => update('patinadoras', itemId, 'nombre', valor)}
          onFile={(itemId, file) => guardarArchivo('img', 'patinadoras', itemId, 'foto', file)}
          onAdd={(nombre) => mutate((draft) => {
            const categoriaId = draft.categorias[0]?.id || ''
            draft.patinadoras.push({
              id: id('pat'),
              orden: draft.patinadoras.filter((item) => item.categoriaId === categoriaId).length + 1,
              estado: 'pendiente',
              nombre,
              edad: '',
              clubId: draft.clubes[0]?.id || '',
              tecnicaId: draft.tecnicas[0]?.id || '',
              categoriaId,
              musicaId: slug(nombre),
              musica: nombre,
              audio: '',
              foto: '',
            })
          }, `Patinadora agregada: ${nombre}`)}
        />
      </Panel>
    </main>
  )
}

function ConfigPuntaje({ data, mutate }) {
  const [nombre, setNombre] = useState('')

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
          <button className="danger" onClick={() => quitar(concepto.id)}>Quitar</button>
        </div>
      ))}
      <div>
        <input placeholder="Nuevo concepto: elemento técnico, componente..." value={nombre} onChange={(event) => setNombre(event.target.value)} />
        <button onClick={agregar}>Agregar</button>
      </div>
    </div>
  )
}

function EditorEntidades({ items, srcCampo, ayudaArchivo, extra, onName, onFile, onAdd }) {
  const [nuevo, setNuevo] = useState('')
  const [guardados, setGuardados] = useState({})

  function ok(itemId) {
    setGuardados((prev) => ({ ...prev, [itemId]: true }))
    window.setTimeout(() => {
      setGuardados((prev) => ({ ...prev, [itemId]: false }))
    }, 1200)
  }

  function agregar() {
    if (!nuevo.trim()) return
    onAdd(nuevo.trim())
    setNuevo('')
  }

  return (
    <div className="entity-editor">
      {items.map((item) => (
        <div className="entity-row" key={item.id}>
          <Avatar src={item[srcCampo]} label={item.nombre} color={item.color} />
          <input value={item.nombre} onChange={(event) => onName(item.id, event.target.value)} />
          <button className="ok-button" onClick={() => ok(item.id)}>{guardados[item.id] ? 'OK ✓' : 'OK'}</button>
          {extra && <small>{extra(item)}</small>}
          <label className="file-hint">{ayudaArchivo}<input type="file" accept="image/*" onChange={(event) => onFile(item.id, event.target.files[0])} /></label>
        </div>
      ))}
      <div className="entity-add">
        <input placeholder="Agregar manual" value={nuevo} onChange={(event) => setNuevo(event.target.value)} />
        <button onClick={agregar}>Agregar</button>
      </div>
    </div>
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
        <table><thead><tr><th>Club</th><th>1</th><th>2</th><th>3</th><th>Podios</th><th>Incentivos</th></tr></thead><tbody>{rankingClubes(data, modo).map((row) => <tr key={row.clubId}><td>{data.clubes.find((c) => c.id === row.clubId)?.nombre}</td><td>{row.primero}</td><td>{row.segundo}</td><td>{row.tercero}</td><td>{row.podios}</td><td>{row.incentivos}</td></tr>)}</tbody></table>
      </Panel>
    </main>
  )
}

function Rankings({ data, modo }) {
  return (
    <main className="grid two">
      <Panel title="Tanteadores por categoría">{data.categorias.map((categoria) => <Tabla key={categoria.id} titulo={categoria.nombre} rows={rankingCategoria(data, categoria.id, modo).filter((row) => row.confirmado && row.total != null)} clubes={data.clubes} />)}</Panel>
      <RankingClubes data={data} modo={modo} />
    </main>
  )
}

function Actas({ data, modo }) {
  return (
    <main className="print">
      <button className="no-print" onClick={() => window.print()}>Imprimir actas</button>
      {data.categorias.map((categoria) => <article className="sheet" key={categoria.id}><h2>Acta categoría: {categoria.nombre}</h2><Tabla rows={rankingCategoria(data, categoria.id, modo)} clubes={data.clubes} /><Firmas jueces={data.jueces} /></article>)}
      <article className="sheet"><h2>Premiación final por clubes</h2><Rankings data={data} modo={modo} /></article>
    </main>
  )
}

function Logs({ data }) {
  return <main><Panel title="Registros básicos"><div className="logs">{data.logs.map((log) => <div key={log.id}><time>{new Date(log.fecha).toLocaleString()}</time><span>{log.texto}</span></div>)}</div></Panel></main>
}

function Competidor({ actual, total }) {
  return (
    <section className="competidor">
      <Avatar src={actual.patinadora?.foto} label={actual.patinadora?.nombre} color={actual.club?.color} />
      <div><small>{actual.categoria?.nombre}</small><h2>{actual.patinadora?.nombre}</h2><p>{actual.club?.nombre} - {actual.tecnica?.nombre || 'Sin técnica'} - {actual.patinadora?.edad || '-'} años</p></div>
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

function Panel({ title, children }) {
  return <section className="panel"><h2>{title}</h2>{children}</section>
}

function Timer({ hasta }) {
  const [now, setNow] = useState(() => Date.now())
  useEffect(() => {
    const timer = setInterval(() => setNow(Date.now()), 500)
    return () => clearInterval(timer)
  }, [])
  const left = Math.max(0, hasta - now)
  return <b>{Math.floor(left / 60000)}:{String(Math.floor((left % 60000) / 1000)).padStart(2, '0')}</b>
}

function Firmas({ jueces }) {
  return <div className="firmas">{jueces.map((juez) => <div key={juez.id}><span></span><p>{juez.nombre}</p></div>)}</div>
}
