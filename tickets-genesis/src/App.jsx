import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  Headphones, Code2, Wrench, Plus, Search, Clock, CheckCircle2, XCircle,
  RotateCw, LayoutDashboard, ListChecks, User, Building2, X, ArrowLeft,
  MessageSquare, Send, Trash2, AlertTriangle, Ticket as TicketIcon,
  Users, LogOut, Mail, Lock, ShieldCheck, Eye, Bell, Download, Paperclip, FileText, Image as ImageIcon, Pencil
} from 'lucide-react';

/* ---------- Configuración de dominio ---------- */

const TIPOS = {
  soporte: { label: 'Soporte', color: '#2E6F95', bg: '#EAF2F6', Icon: Headphones, desc: 'Equipos, accesos, redes, correo' },
  programacion: { label: 'Programación', color: '#6B4E9E', bg: '#F1ECF7', Icon: Code2, desc: 'Sistemas, desarrollos, integraciones' },
  mantenimiento: { label: 'Mantenimiento', color: '#C97A2B', bg: '#FBF0E4', Icon: Wrench, desc: 'Instalaciones, mobiliario, infraestructura' },
};

const PRIORIDADES = {
  baja: { label: 'Baja', color: '#7C93A6' },
  media: { label: 'Media', color: '#B8912F' },
  alta: { label: 'Alta', color: '#C1583A' },
  urgente: { label: 'Urgente', color: '#A3223A' },
};

const ESTADOS = {
  pendiente: { label: 'Pendiente', color: '#8A8F98', Icon: Clock },
  en_proceso: { label: 'En proceso', color: '#2E6F95', Icon: RotateCw },
  resuelto: { label: 'Resuelto', color: '#3F8F5F', Icon: CheckCircle2 },
  cerrado: { label: 'Cerrado', color: '#4B5563', Icon: XCircle },
};

const ESTADO_ORDEN = ['pendiente', 'en_proceso', 'resuelto', 'cerrado'];
const STORAGE_TICKETS = 'ticket-system-store';
const STORAGE_USERS = 'ticket-system-users';
const STORAGE_NOTIFICACIONES = 'ticket-system-notifications';
const DOMINIO = '@genesisig.com';
const MAX_ADJUNTO_MB = 3;
const MAX_ADJUNTOS_TICKET = 6;
const TIPOS_ADJUNTO_ACEPTADOS = 'image/*,.pdf,.doc,.docx,.xls,.xlsx,.csv,.txt';
const ADMIN_SEED = { email: 'admin@genesisig.com', nombre: 'Administrador General', area: 'Sistemas', password: 'Genesis2026', rol: 'admin' };

function formatFecha(iso) {
  const d = new Date(iso);
  const fecha = d.toLocaleDateString('es-ES', { day: '2-digit', month: 'short', year: 'numeric' });
  const hora = d.toLocaleTimeString('es-ES', { hour: '2-digit', minute: '2-digit' });
  return `${fecha} · ${hora}`;
}

function nuevoId(seq) {
  return `TK-${String(seq).padStart(6, '0')}`;
}

function formatBytes(bytes) {
  if (!bytes) return '0 KB';
  const kb = bytes / 1024;
  if (kb < 1024) return `${kb.toFixed(0)} KB`;
  return `${(kb / 1024).toFixed(1)} MB`;
}

function esImagenAdjunto(tipo) {
  return typeof tipo === 'string' && tipo.startsWith('image/');
}

function luminanciaRelativa(hex) {
  const c = hex.replace('#', '');
  const [r, g, b] = [0, 2, 4].map(i => parseInt(c.substring(i, i + 2), 16) / 255);
  const lin = v => (v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4));
  return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b);
}

function contrasteEntre(hexA, hexB) {
  const [claro, oscuro] = [luminanciaRelativa(hexA), luminanciaRelativa(hexB)].sort((a, b) => b - a);
  return (claro + 0.05) / (oscuro + 0.05);
}

function textoLegibleSobre(hexFondo) {
  return contrasteEntre(hexFondo, '#FFFFFF') >= contrasteEntre(hexFondo, '#1B2430') ? '#FFFFFF' : '#1B2430';
}

function leerArchivoComoDataURL(archivo) {
  return new Promise((resolve, reject) => {
    const lector = new FileReader();
    lector.onload = () => resolve(lector.result);
    lector.onerror = () => reject(lector.error);
    lector.readAsDataURL(archivo);
  });
}

function validarCorreo(email) {
  return typeof email === 'string' && email.trim().toLowerCase().endsWith(DOMINIO);
}

function validarPassword(pw) {
  pw = pw || '';
  if (pw.length < 6 || pw.length > 12) return { ok: false, error: 'La contraseña debe tener entre 6 y 12 caracteres.' };
  if (!/[A-Z]/.test(pw)) return { ok: false, error: 'La contraseña debe incluir al menos una letra mayúscula.' };
  if (!/[0-9]/.test(pw)) return { ok: false, error: 'La contraseña debe incluir al menos un número.' };
  if (!/[^A-Za-z0-9]/.test(pw)) return { ok: false, error: 'La contraseña debe incluir al menos un carácter especial.' };
  return { ok: true };
}

const NIVELES_FUERZA = [
  { max: 1, label: 'Muy débil', color: '#A3223A' },
  { max: 3, label: 'Débil', color: '#C1583A' },
  { max: 4, label: 'Media', color: '#B8912F' },
  { max: 5, label: 'Fuerte', color: '#3F8F5F' },
  { max: 6, label: 'Muy fuerte', color: '#2E6F95' },
];

function evaluarFuerzaPassword(pw) {
  pw = pw || '';
  let puntos = 0;
  if (pw.length >= 6) puntos++;
  if (pw.length >= 9) puntos++;
  if (/[a-z]/.test(pw)) puntos++;
  if (/[A-Z]/.test(pw)) puntos++;
  if (/[0-9]/.test(pw)) puntos++;
  if (/[^A-Za-z0-9]/.test(pw)) puntos++;
  const nivel = NIVELES_FUERZA.find(n => puntos <= n.max) || NIVELES_FUERZA[NIVELES_FUERZA.length - 1];
  return { puntos, ...nivel };
}

function exportarTicketsExcel(tickets) {
  const filas = tickets.map(t => ({
    ID: t.id,
    Título: t.titulo,
    Descripción: t.descripcion,
    Área: TIPOS[t.tipo]?.label || t.tipo,
    Prioridad: PRIORIDADES[t.prioridad]?.label || t.prioridad,
    Estado: ESTADOS[t.estado]?.label || t.estado,
    'Área solicitante': t.area,
    Solicitante: t.solicitante,
    'Creado por': t.creadoPor,
    'Asignado a': t.asignadoA || 'Sin asignar',
    Creado: formatFecha(t.creado),
    Actualizado: formatFecha(t.actualizado),
    Comentarios: t.notas.length,
    'Adjuntos evidencia': (t.adjuntos || []).length,
  }));
  const hoja = XLSX.utils.json_to_sheet(filas);
  hoja['!cols'] = Object.keys(filas[0] || {}).map(k => ({ wch: Math.max(k.length + 2, 16) }));
  const libro = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(libro, hoja, 'Tickets');
  const fecha = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(libro, `tickets-${fecha}.xlsx`);
}

/* ---------- Notificaciones ----------
   Envío de correo real pendiente: conectar un servicio (ej. EmailJS) aquí.
   Por ahora la notificación queda registrada in-app y en consola. */
async function enviarNotificacionEmail(destinatario, asunto, mensaje) {
  console.info(`[correo pendiente de configurar] Para: ${destinatario} · ${asunto} · ${mensaje}`);
}

/* ---------- Piezas visuales pequeñas ---------- */

function Pill({ label, color, Icon, uppercase }) {
  return (
    <span
      className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium"
      style={{ background: color + '1A', color, border: `1px solid ${color}40`, textTransform: uppercase ? 'uppercase' : 'none', letterSpacing: uppercase ? '0.03em' : 'normal' }}
    >
      {Icon && <Icon size={12} />}
      {label}
    </span>
  );
}

function TipoBadge({ tipo, size = 14 }) {
  const t = TIPOS[tipo];
  const Icon = t.Icon;
  return (
    <span className="inline-flex items-center gap-1.5 px-2 py-1 rounded-md text-xs font-semibold" style={{ background: t.bg, color: t.color }}>
      <Icon size={size} />
      {t.label}
    </span>
  );
}

function EmptyState({ texto, sub }) {
  return (
    <div className="w-full flex flex-col items-center justify-center py-16 text-center" style={{ color: 'var(--muted)' }}>
      <TicketIcon size={32} style={{ opacity: 0.35, marginBottom: 10 }} />
      <p className="font-medium" style={{ color: 'var(--ink)' }}>{texto}</p>
      {sub && <p className="text-sm" style={{ marginTop: 4 }}>{sub}</p>}
    </div>
  );
}

function MedidorFuerzaPassword({ password }) {
  if (!password) return null;
  const { puntos, label, color } = evaluarFuerzaPassword(password);
  const porcentaje = Math.min(100, (puntos / 6) * 100);
  const requisitos = [
    { ok: password.length >= 6 && password.length <= 12, texto: '6 a 12 caracteres' },
    { ok: /[A-Z]/.test(password), texto: 'Una letra mayúscula' },
    { ok: /[0-9]/.test(password), texto: 'Un número' },
    { ok: /[^A-Za-z0-9]/.test(password), texto: 'Un carácter especial' },
  ];
  return (
    <div className="mt-1.5">
      <div className="rounded-full overflow-hidden" style={{ height: 5, background: 'var(--line)' }}>
        <div style={{ width: `${porcentaje}%`, height: '100%', background: color, transition: 'width .2s ease, background .2s ease' }} />
      </div>
      <p className="text-xs font-medium mt-1" style={{ color }}>{label}</p>
      <ul className="grid grid-cols-2 gap-x-2 gap-y-0.5 mt-1">
        {requisitos.map((r, i) => (
          <li key={i} className="flex items-center gap-1 text-xs" style={{ color: r.ok ? '#3F8F5F' : 'var(--muted)' }}>
            {r.ok ? <CheckCircle2 size={11} /> : <XCircle size={11} />}
            {r.texto}
          </li>
        ))}
      </ul>
    </div>
  );
}

function AdjuntoItem({ adjunto, onQuitar }) {
  const esImg = esImagenAdjunto(adjunto.tipo);
  return (
    <div className="flex items-center gap-2 p-2 rounded-md" style={{ background: 'var(--bg)', border: '1px solid var(--line)' }}>
      {esImg ? (
        <a href={adjunto.dataUrl} target="_blank" rel="noreferrer" className="flex-shrink-0">
          <img src={adjunto.dataUrl} alt={adjunto.nombre} style={{ width: 40, height: 40, objectFit: 'cover', borderRadius: 6, display: 'block' }} />
        </a>
      ) : (
        <a href={adjunto.dataUrl} download={adjunto.nombre} className="flex items-center justify-center flex-shrink-0" style={{ width: 40, height: 40, borderRadius: 6, background: 'var(--card)', border: '1px solid var(--line)' }}>
          <FileText size={18} style={{ color: 'var(--muted)' }} />
        </a>
      )}
      <div className="min-w-0 flex-1">
        <a href={adjunto.dataUrl} target={esImg ? '_blank' : undefined} download={esImg ? undefined : adjunto.nombre} rel="noreferrer" className="text-xs font-medium truncate block" style={{ color: 'var(--ink)' }}>
          {adjunto.nombre}
        </a>
        <p className="text-xs" style={{ color: 'var(--muted)' }}>
          {formatBytes(adjunto.tamano)}{adjunto.subidoNombre ? ` · ${adjunto.subidoNombre}` : ''}
        </p>
      </div>
      {onQuitar && (
        <button type="button" onClick={onQuitar} className="icon-btn flex-shrink-0" style={{ width: 26, height: 26 }} title="Quitar">
          <X size={13} />
        </button>
      )}
    </div>
  );
}

function SelectorAdjuntos({ onSeleccionar, restantes, disabled }) {
  const [error, setError] = useState('');
  const inputRef = useRef(null);

  async function manejarSeleccion(e) {
    const archivos = Array.from(e.target.files || []);
    e.target.value = '';
    if (!archivos.length) return;
    setError('');
    if (archivos.length > restantes) {
      setError(`Solo puedes agregar ${restantes} archivo${restantes === 1 ? '' : 's'} más en este ticket.`);
      return;
    }
    const validos = [];
    for (const archivo of archivos) {
      if (archivo.size > MAX_ADJUNTO_MB * 1024 * 1024) {
        setError(`"${archivo.name}" supera ${MAX_ADJUNTO_MB} MB.`);
        continue;
      }
      try {
        const dataUrl = await leerArchivoComoDataURL(archivo);
        validos.push({ nombre: archivo.name, tipo: archivo.type || 'application/octet-stream', tamano: archivo.size, dataUrl });
      } catch {
        setError(`No se pudo leer "${archivo.name}".`);
      }
    }
    if (validos.length) onSeleccionar(validos);
  }

  const agotado = restantes <= 0;

  return (
    <div>
      <button
        type="button"
        disabled={disabled || agotado}
        onClick={() => inputRef.current?.click()}
        className="btn-secondary flex items-center gap-1.5"
        style={{ padding: '7px 12px', opacity: (disabled || agotado) ? 0.5 : 1 }}
      >
        <Paperclip size={13} /> Adjuntar evidencia
      </button>
      <input ref={inputRef} type="file" multiple accept={TIPOS_ADJUNTO_ACEPTADOS} onChange={manejarSeleccion} style={{ display: 'none' }} />
      <p className="text-xs mt-1" style={{ color: 'var(--muted)' }}>
        {agotado ? `Máximo de ${MAX_ADJUNTOS_TICKET} archivos alcanzado.` : `Imágenes o documentos, máx. ${MAX_ADJUNTO_MB} MB c/u · hasta ${MAX_ADJUNTOS_TICKET} por ticket.`}
      </p>
      {error && <p className="text-xs mt-1" style={{ color: PRIORIDADES.urgente.color }}>{error}</p>}
    </div>
  );
}

function Toast({ message }) {
  return (
    <div
      className="fixed left-1/2 z-50 px-4 py-2 rounded-lg text-sm font-medium shadow-xl"
      style={{ bottom: 24, transform: 'translateX(-50%)', background: 'var(--ink)', color: '#fff' }}
    >
      {message}
    </div>
  );
}

function NotificationBell({ notificaciones, onAbrir, onMarcarTodas }) {
  const [abierto, setAbierto] = useState(false);
  const ref = useRef(null);
  const sinLeer = notificaciones.filter(n => !n.leida).length;

  useEffect(() => {
    function onClickFuera(e) { if (ref.current && !ref.current.contains(e.target)) setAbierto(false); }
    document.addEventListener('mousedown', onClickFuera);
    return () => document.removeEventListener('mousedown', onClickFuera);
  }, []);

  return (
    <div className="relative" ref={ref}>
      <button onClick={() => setAbierto(v => !v)} className="icon-btn relative" title="Notificaciones">
        <Bell size={16} />
        {sinLeer > 0 && (
          <span className="absolute flex items-center justify-center rounded-full text-white font-semibold"
            style={{ top: -3, right: -3, width: 16, height: 16, fontSize: 9.5, background: PRIORIDADES.urgente.color }}>
            {sinLeer > 9 ? '9+' : sinLeer}
          </span>
        )}
      </button>
      {abierto && (
        <div className="absolute right-0 mt-2 rounded-lg overflow-hidden z-50" style={{ width: 320, maxHeight: 380, background: 'var(--card)', border: '1px solid var(--line)', boxShadow: '0 8px 24px rgba(27,36,48,0.18)' }}>
          <div className="flex items-center justify-between px-3 py-2.5" style={{ borderBottom: '1px solid var(--line)' }}>
            <p className="text-sm font-semibold" style={{ color: 'var(--ink)' }}>Notificaciones</p>
            {sinLeer > 0 && (
              <button onClick={onMarcarTodas} className="text-xs font-medium" style={{ color: 'var(--muted)' }}>Marcar todas leídas</button>
            )}
          </div>
          <div className="overflow-y-auto" style={{ maxHeight: 330 }}>
            {notificaciones.length === 0 ? (
              <p className="text-xs text-center py-8" style={{ color: 'var(--muted)' }}>Sin notificaciones todavía.</p>
            ) : (
              notificaciones.slice(0, 30).map(n => (
                <button key={n.id} onClick={() => { onAbrir(n); setAbierto(false); }} className="w-full text-left px-3 py-2.5 flex gap-2 items-start" style={{ borderBottom: '1px solid var(--line)', background: n.leida ? 'transparent' : 'var(--bg)' }}>
                  {!n.leida && <span className="rounded-full flex-shrink-0" style={{ width: 6, height: 6, marginTop: 5, background: PRIORIDADES.urgente.color }} />}
                  <span className="min-w-0" style={{ marginLeft: n.leida ? 14 : 0 }}>
                    <p className="text-xs" style={{ color: 'var(--ink)' }}>{n.mensaje}</p>
                    <p className="text-xs mt-0.5 font-mono" style={{ color: 'var(--muted)' }}>{formatFecha(n.fecha)}</p>
                  </span>
                </button>
              ))
            )}
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------- Pantalla de acceso ---------- */

function LoginScreen({ onLogin, onRegistrar }) {
  const [modo, setModo] = useState('login');
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [confirmar, setConfirmar] = useState('');
  const [nombre, setNombre] = useState('');
  const [area, setArea] = useState('');
  const [error, setError] = useState('');

  function submitLogin(e) {
    e.preventDefault();
    const res = onLogin(email, password);
    setError(res.ok ? '' : res.error);
  }

  function submitRegistro(e) {
    e.preventDefault();
    const chequeoPw = validarPassword(password);
    if (!chequeoPw.ok) { setError(chequeoPw.error); return; }
    if (password !== confirmar) { setError('Las contraseñas no coinciden.'); return; }
    const res = onRegistrar({ email, password, nombre, area });
    setError(res.ok ? '' : res.error);
  }

  function cambiarModo(m) {
    setModo(m); setError(''); setPassword(''); setConfirmar('');
  }

  return (
    <div className="flex items-center justify-center px-4" style={{ minHeight: '100vh' }}>
      <div className="w-full max-w-sm">
        <div className="flex flex-col items-center mb-6">
          <div className="flex items-center justify-center rounded-lg mb-3" style={{ width: 44, height: 44, background: 'var(--ink)', color: 'var(--bg)' }}>
            <Wrench size={20} />
          </div>
          <h1 className="font-display font-semibold" style={{ fontSize: 21, color: 'var(--ink)' }}>Central de Tickets</h1>
          <p className="text-xs" style={{ color: 'var(--muted)', marginTop: 2 }}>Acceso exclusivo para cuentas @genesisig.com</p>
        </div>

        <div className="rounded-lg p-5" style={{ background: 'var(--card)', border: '1px solid var(--line)' }}>
          <div className="flex gap-1 p-1 rounded-lg mb-4" style={{ background: 'var(--bg)' }}>
            <button onClick={() => cambiarModo('login')} className="flex-1 text-sm font-medium py-1.5 rounded-md transition"
              style={{ background: modo === 'login' ? 'var(--card)' : 'transparent', color: modo === 'login' ? 'var(--ink)' : 'var(--muted)', boxShadow: modo === 'login' ? '0 1px 3px rgba(27,36,48,0.12)' : 'none' }}>
              Iniciar sesión
            </button>
            <button onClick={() => cambiarModo('registro')} className="flex-1 text-sm font-medium py-1.5 rounded-md transition"
              style={{ background: modo === 'registro' ? 'var(--card)' : 'transparent', color: modo === 'registro' ? 'var(--ink)' : 'var(--muted)', boxShadow: modo === 'registro' ? '0 1px 3px rgba(27,36,48,0.12)' : 'none' }}>
              Crear cuenta
            </button>
          </div>

          {modo === 'login' ? (
            <form onSubmit={submitLogin} className="flex flex-col gap-3">
              <div>
                <label className="text-xs font-semibold uppercase" style={{ color: 'var(--muted)', letterSpacing: '0.05em' }}>Correo corporativo</label>
                <div className="relative mt-1.5">
                  <Mail size={14} style={{ position: 'absolute', left: 10, top: 12, color: 'var(--muted)' }} />
                  <input value={email} onChange={e => setEmail(e.target.value)} placeholder="nombre@genesisig.com" className="input-field" style={{ paddingLeft: 30 }} />
                </div>
              </div>
              <div>
                <label className="text-xs font-semibold uppercase" style={{ color: 'var(--muted)', letterSpacing: '0.05em' }}>Contraseña</label>
                <div className="relative mt-1.5">
                  <Lock size={14} style={{ position: 'absolute', left: 10, top: 12, color: 'var(--muted)' }} />
                  <input type="password" value={password} onChange={e => setPassword(e.target.value)} placeholder="••••••••" className="input-field" style={{ paddingLeft: 30 }} />
                </div>
              </div>
              {error && <p className="text-xs" style={{ color: PRIORIDADES.urgente.color }}>{error}</p>}
              <button type="submit" className="btn-primary justify-center mt-1">Entrar</button>
            </form>
          ) : (
            <form onSubmit={submitRegistro} className="flex flex-col gap-3">
              <div>
                <label className="text-xs font-semibold uppercase" style={{ color: 'var(--muted)', letterSpacing: '0.05em' }}>Nombre completo</label>
                <input value={nombre} onChange={e => setNombre(e.target.value)} placeholder="Ej. Juan Pérez" className="input-field mt-1.5" />
              </div>
              <div>
                <label className="text-xs font-semibold uppercase" style={{ color: 'var(--muted)', letterSpacing: '0.05em' }}>Área / departamento</label>
                <input value={area} onChange={e => setArea(e.target.value)} placeholder="Ej. Ventas" className="input-field mt-1.5" />
              </div>
              <div>
                <label className="text-xs font-semibold uppercase" style={{ color: 'var(--muted)', letterSpacing: '0.05em' }}>Correo corporativo</label>
                <div className="relative mt-1.5">
                  <Mail size={14} style={{ position: 'absolute', left: 10, top: 12, color: 'var(--muted)' }} />
                  <input value={email} onChange={e => setEmail(e.target.value)} placeholder="nombre@genesisig.com" className="input-field" style={{ paddingLeft: 30 }} />
                </div>
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="text-xs font-semibold uppercase" style={{ color: 'var(--muted)', letterSpacing: '0.05em' }}>Contraseña</label>
                  <input type="password" value={password} onChange={e => setPassword(e.target.value)} placeholder="6-12, may., número, especial" className="input-field mt-1.5" />
                </div>
                <div>
                  <label className="text-xs font-semibold uppercase" style={{ color: 'var(--muted)', letterSpacing: '0.05em' }}>Confirmar</label>
                  <input type="password" value={confirmar} onChange={e => setConfirmar(e.target.value)} placeholder="Repite" className="input-field mt-1.5" />
                </div>
              </div>
              <MedidorFuerzaPassword password={password} />
              {error && <p className="text-xs" style={{ color: PRIORIDADES.urgente.color }}>{error}</p>}
              <button type="submit" className="btn-primary justify-center mt-1">Crear cuenta</button>
              <p className="text-xs text-center" style={{ color: 'var(--muted)' }}>Las cuentas nuevas ingresan con perfil de usuario.</p>
            </form>
          )}
        </div>

        <p className="text-xs text-center mt-4" style={{ color: 'var(--muted)' }}>
          Cuenta de prueba admin: <span className="font-mono">admin@genesisig.com</span> / <span className="font-mono">Genesis2026</span>
        </p>
      </div>
    </div>
  );
}

/* ---------- Cabecera / navegación ---------- */

function Header({ view, setView, total, user, onLogout, esAdmin, notificaciones, onAbrirNotif, onMarcarTodasNotif }) {
  const NAV = [
    { key: 'dashboard', label: 'Panel', Icon: LayoutDashboard },
    { key: 'nuevo', label: 'Nuevo ticket', Icon: Plus },
    { key: 'tickets', label: 'Tickets', Icon: ListChecks },
    ...(esAdmin ? [{ key: 'usuarios', label: 'Usuarios', Icon: Users }] : []),
  ];
  return (
    <header style={{ borderBottom: '1px solid var(--line)', background: 'var(--card)' }}>
      <div className="max-w-6xl mx-auto px-4 md:px-6 py-4 flex items-center justify-between gap-4 flex-wrap">
        <div className="flex items-center gap-3">
          <div className="flex items-center justify-center rounded-lg" style={{ width: 38, height: 38, background: 'var(--ink)', color: 'var(--bg)' }}>
            <Wrench size={18} />
          </div>
          <div>
            <h1 className="font-display font-semibold" style={{ fontSize: 19, lineHeight: 1, color: 'var(--ink)' }}>Central de Tickets</h1>
            <p className="text-xs" style={{ color: 'var(--muted)', marginTop: 3 }}>
              {esAdmin ? `${total} ticket${total === 1 ? '' : 's'} en total` : `${total} ticket${total === 1 ? '' : 's'} tuyo${total === 1 ? '' : 's'}`}
            </p>
          </div>
        </div>
        <nav className="flex items-center gap-1 p-1 rounded-lg" style={{ background: 'var(--bg)' }}>
          {NAV.map(n => {
            const active = view === n.key;
            return (
              <button key={n.key} onClick={() => setView(n.key)} className="flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium transition"
                style={{ background: active ? 'var(--card)' : 'transparent', color: active ? 'var(--ink)' : 'var(--muted)', boxShadow: active ? '0 1px 3px rgba(27,36,48,0.12)' : 'none' }}>
                <n.Icon size={15} />
                <span className="hidden sm:inline">{n.label}</span>
              </button>
            );
          })}
        </nav>
        <div className="flex items-center gap-2">
          <NotificationBell notificaciones={notificaciones} onAbrir={onAbrirNotif} onMarcarTodas={onMarcarTodasNotif} />
          <div className="text-right hidden sm:block">
            <p className="text-sm font-medium" style={{ color: 'var(--ink)' }}>{user.nombre}</p>
            <Pill label={esAdmin ? 'Admin' : 'Usuario'} color={esAdmin ? '#6B4E9E' : '#7C93A6'} Icon={esAdmin ? ShieldCheck : User} uppercase />
          </div>
          <button onClick={onLogout} className="icon-btn" title="Cerrar sesión"><LogOut size={16} /></button>
        </div>
      </div>
    </header>
  );
}

/* ---------- Tarjeta de estadística ---------- */

function StatCard({ label, value, color }) {
  return (
    <div className="stub-flat p-4 rounded-lg" style={{ background: 'var(--card)', border: '1px solid var(--line)' }}>
      <p className="text-xs font-medium uppercase" style={{ color: 'var(--muted)', letterSpacing: '0.05em' }}>{label}</p>
      <p className="font-display font-semibold" style={{ fontSize: 30, color: color || 'var(--ink)', marginTop: 4 }}>{value}</p>
    </div>
  );
}

/* ---------- Fila de ticket estilo "boleto perforado" ---------- */

function TicketRow({ ticket, onOpen }) {
  const tipo = TIPOS[ticket.tipo];
  const estado = ESTADOS[ticket.estado];
  const prioridad = PRIORIDADES[ticket.prioridad];
  const EstadoIcon = estado.Icon;
  return (
    <div className="ticket-row" onClick={() => onOpen(ticket)}>
      <div className="stub-col" style={{ background: tipo.bg }}>
        <tipo.Icon size={18} color={tipo.color} />
        <span className="font-mono" style={{ fontSize: 10.5, color: tipo.color, fontWeight: 600 }}>{ticket.id}</span>
        <span className="notch-top" />
        <span className="notch-bottom" />
      </div>
      <div className="flex-1 min-w-0 flex items-center justify-between gap-3 px-4 py-3 flex-wrap">
        <div className="min-w-0" style={{ flex: '1 1 240px' }}>
          <p className="font-medium truncate" style={{ color: 'var(--ink)', fontSize: 14.5 }}>{ticket.titulo}</p>
          <p className="text-xs truncate" style={{ color: 'var(--muted)', marginTop: 2 }}>{ticket.area} · {ticket.solicitante}</p>
        </div>
        <div className="flex items-center gap-2 flex-wrap justify-end" style={{ flex: '0 0 auto' }}>
          <Pill label={prioridad.label} color={prioridad.color} uppercase />
          <Pill label={estado.label} color={estado.color} Icon={EstadoIcon} />
          <span className="font-mono text-xs hidden md:inline" style={{ color: 'var(--muted)', minWidth: 118, textAlign: 'right' }}>{formatFecha(ticket.creado)}</span>
        </div>
      </div>
    </div>
  );
}

/* ---------- Panel / Dashboard ---------- */

/* ---------- Gráfica: tickets por área y estado ---------- */

function GraficaAreaEstado({ tickets }) {
  const [vistaTabla, setVistaTabla] = useState(false);
  const [activo, setActivo] = useState(null);

  const filas = Object.keys(TIPOS).map(tipoKey => {
    const deArea = tickets.filter(t => t.tipo === tipoKey);
    const porEstado = ESTADO_ORDEN.reduce((acc, e) => { acc[e] = deArea.filter(t => t.estado === e).length; return acc; }, {});
    return { tipo: tipoKey, total: deArea.length, porEstado };
  });
  const maxTotal = Math.max(1, ...filas.map(f => f.total));

  if (tickets.length === 0) return null;

  return (
    <div className="stub-flat rounded-lg p-4 md:p-5" style={{ background: 'var(--card)', border: '1px solid var(--line)' }}>
      <div className="flex items-start justify-between flex-wrap gap-3 mb-1">
        <div>
          <h2 className="font-display font-semibold" style={{ color: 'var(--ink)', fontSize: 15 }}>Tickets por área y estado</h2>
          <p className="text-xs mt-0.5" style={{ color: 'var(--muted)' }}>Cuántos están pendientes, en proceso, resueltos o cerrados en cada área.</p>
        </div>
        <button type="button" onClick={() => setVistaTabla(v => !v)} className="btn-secondary" style={{ padding: '5px 10px', fontSize: 12 }}>
          {vistaTabla ? 'Ver gráfica' : 'Ver tabla'}
        </button>
      </div>

      <div className="flex items-center gap-3 flex-wrap my-3">
        {ESTADO_ORDEN.map(k => (
          <span key={k} className="flex items-center gap-1.5 text-xs" style={{ color: 'var(--muted)' }}>
            <span style={{ width: 10, height: 10, borderRadius: 3, background: ESTADOS[k].color, display: 'inline-block', flexShrink: 0 }} />
            {ESTADOS[k].label}
          </span>
        ))}
      </div>

      {vistaTabla ? (
        <div className="overflow-x-auto">
          <table className="w-full text-sm" style={{ borderCollapse: 'collapse' }}>
            <thead>
              <tr style={{ borderBottom: '1px solid var(--line)' }}>
                <th className="text-left py-2 pr-3 text-xs font-semibold uppercase" style={{ color: 'var(--muted)' }}>Área</th>
                {ESTADO_ORDEN.map(k => (
                  <th key={k} className="text-right py-2 px-3 text-xs font-semibold uppercase" style={{ color: 'var(--muted)' }}>{ESTADOS[k].label}</th>
                ))}
                <th className="text-right py-2 pl-3 text-xs font-semibold uppercase" style={{ color: 'var(--muted)' }}>Total</th>
              </tr>
            </thead>
            <tbody>
              {filas.map(f => (
                <tr key={f.tipo} style={{ borderBottom: '1px solid var(--line)' }}>
                  <td className="py-2 pr-3" style={{ color: 'var(--ink)' }}>{TIPOS[f.tipo].label}</td>
                  {ESTADO_ORDEN.map(k => (
                    <td key={k} className="text-right py-2 px-3 font-mono" style={{ color: 'var(--ink)' }}>{f.porEstado[k]}</td>
                  ))}
                  <td className="text-right py-2 pl-3 font-mono font-semibold" style={{ color: 'var(--ink)' }}>{f.total}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : (
        <div className="flex flex-col gap-5 mt-2">
          {filas.map(f => {
            const t = TIPOS[f.tipo];
            const segmentos = ESTADO_ORDEN.filter(k => f.porEstado[k] > 0);
            return (
              <div key={f.tipo}>
                <div className="flex items-center justify-between mb-1.5">
                  <span className="flex items-center gap-1.5 text-sm font-medium" style={{ color: 'var(--ink)' }}>
                    <t.Icon size={13} color={t.color} /> {t.label}
                  </span>
                  <span className="text-xs font-mono" style={{ color: 'var(--muted)' }}>{f.total} ticket{f.total === 1 ? '' : 's'}</span>
                </div>
                <div className="flex items-stretch" style={{ height: 24, borderRadius: 6, background: 'var(--bg)' }}>
                  {segmentos.map((k, i) => {
                    const cantidad = f.porEstado[k];
                    const pctDelMax = (cantidad / maxTotal) * 100;
                    const key = `${f.tipo}|${k}`;
                    const primero = i === 0;
                    const ultimo = i === segmentos.length - 1;
                    const colorTexto = textoLegibleSobre(ESTADOS[k].color);
                    return (
                      <button
                        type="button"
                        key={k}
                        onMouseEnter={() => setActivo(key)}
                        onMouseLeave={() => setActivo(a => (a === key ? null : a))}
                        onFocus={() => setActivo(key)}
                        onBlur={() => setActivo(a => (a === key ? null : a))}
                        className="relative flex items-center justify-center"
                        style={{
                          width: `${pctDelMax}%`,
                          minWidth: 4,
                          background: ESTADOS[k].color,
                          marginRight: ultimo ? 0 : 2,
                          borderTopLeftRadius: primero ? 5 : 0,
                          borderBottomLeftRadius: primero ? 5 : 0,
                          borderTopRightRadius: ultimo ? 5 : 0,
                          borderBottomRightRadius: ultimo ? 5 : 0,
                          border: 'none',
                          padding: 0,
                          cursor: 'default',
                        }}
                        aria-label={`${t.label} · ${ESTADOS[k].label}: ${cantidad}`}
                      >
                        {pctDelMax >= 8 && (
                          <span className="text-xs font-semibold font-mono" style={{ color: colorTexto }}>{cantidad}</span>
                        )}
                        {activo === key && (
                          <div
                            className="absolute z-10 px-2 py-1 rounded-md text-xs whitespace-nowrap pointer-events-none"
                            style={{ bottom: '130%', left: '50%', transform: 'translateX(-50%)', background: 'var(--ink)', color: '#fff', boxShadow: '0 4px 12px rgba(27,36,48,0.25)' }}
                          >
                            <b>{cantidad}</b> {ESTADOS[k].label.toLowerCase()} · {t.label}
                          </div>
                        )}
                      </button>
                    );
                  })}
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

function Dashboard({ tickets, onOpen, onNav, onFiltrarTipo, esAdmin }) {
  const total = tickets.length;
  const conteoEstado = ESTADO_ORDEN.reduce((acc, k) => { acc[k] = tickets.filter(t => t.estado === k).length; return acc; }, {});
  const conteoTipo = Object.keys(TIPOS).reduce((acc, k) => { acc[k] = tickets.filter(t => t.tipo === k).length; return acc; }, {});
  const recientes = [...tickets].sort((a, b) => new Date(b.creado) - new Date(a.creado)).slice(0, 5);

  if (total === 0) {
    return (
      <div className="stub-flat rounded-lg p-2" style={{ background: 'var(--card)', border: '1px solid var(--line)' }}>
        <EmptyState texto={esAdmin ? 'Aún no hay tickets registrados' : 'Todavía no has creado tickets'} sub="Crea el primero desde “Nuevo ticket”." />
        <div className="flex justify-center pb-6">
          <button onClick={() => onNav('nuevo')} className="btn-primary"><Plus size={15} /> Crear el primer ticket</button>
        </div>
      </div>
    );
  }

  return (
    <div className="flex flex-col gap-6">
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
        <StatCard label={esAdmin ? 'Total' : 'Mis tickets'} value={total} />
        <StatCard label="Pendientes" value={conteoEstado.pendiente} color={ESTADOS.pendiente.color} />
        <StatCard label="En proceso" value={conteoEstado.en_proceso} color={ESTADOS.en_proceso.color} />
        <StatCard label="Resueltos" value={conteoEstado.resuelto} color={ESTADOS.resuelto.color} />
      </div>

      <div className="grid md:grid-cols-3 gap-6">
        <div className="md:col-span-2 flex flex-col gap-3">
          <h2 className="font-display font-semibold" style={{ color: 'var(--ink)', fontSize: 15 }}>Por área de destino</h2>
          <div className="grid sm:grid-cols-3 gap-3">
            {Object.entries(TIPOS).map(([key, t]) => (
              <button key={key} onClick={() => { onFiltrarTipo(key); onNav('tickets'); }} className="text-left p-4 rounded-lg transition" style={{ background: t.bg, border: `1px solid ${t.color}30` }}>
                <t.Icon size={20} color={t.color} />
                <p className="font-display font-semibold" style={{ fontSize: 24, color: t.color, marginTop: 8 }}>{conteoTipo[key]}</p>
                <p className="text-sm font-medium" style={{ color: 'var(--ink)' }}>{t.label}</p>
                <p className="text-xs" style={{ color: 'var(--muted)', marginTop: 2 }}>{t.desc}</p>
              </button>
            ))}
          </div>
        </div>
        <div className="flex flex-col gap-3 min-w-0">
          <h2 className="font-display font-semibold" style={{ color: 'var(--ink)', fontSize: 15 }}>Recientes</h2>
          <div className="flex flex-col gap-2">
            {recientes.map(t => (
              <button key={t.id} onClick={() => onOpen(t)} className="text-left p-3 rounded-lg flex items-center gap-2 min-w-0" style={{ background: 'var(--card)', border: '1px solid var(--line)' }}>
                <TipoBadge tipo={t.tipo} size={12} />
                <span className="text-sm font-medium truncate" style={{ color: 'var(--ink)' }}>{t.titulo}</span>
              </button>
            ))}
          </div>
        </div>
      </div>

      <GraficaAreaEstado tickets={tickets} />
    </div>
  );
}

/* ---------- Formulario nuevo ticket ---------- */

function NuevoTicketForm({ onCreate, user }) {
  const [tipo, setTipo] = useState('soporte');
  const [titulo, setTitulo] = useState('');
  const [descripcion, setDescripcion] = useState('');
  const [prioridad, setPrioridad] = useState('media');
  const [adjuntos, setAdjuntos] = useState([]);
  const [tocado, setTocado] = useState(false);

  const valido = titulo.trim() && descripcion.trim();

  function submit(e) {
    e.preventDefault();
    setTocado(true);
    if (!valido) return;
    onCreate({ tipo, titulo: titulo.trim(), descripcion: descripcion.trim(), prioridad, adjuntos });
    setTitulo(''); setDescripcion(''); setPrioridad('media'); setAdjuntos([]); setTocado(false);
  }

  return (
    <form onSubmit={submit} className="stub-flat rounded-lg p-5 md:p-6 flex flex-col gap-5 max-w-2xl" style={{ background: 'var(--card)', border: '1px solid var(--line)' }}>
      <div className="flex items-center gap-2 p-2.5 rounded-lg text-sm" style={{ background: 'var(--bg)', color: 'var(--muted)' }}>
        <User size={14} /> Enviando como <b style={{ color: 'var(--ink)' }}>{user.nombre}</b> · {user.area}
      </div>

      <div>
        <label className="text-xs font-semibold uppercase" style={{ color: 'var(--muted)', letterSpacing: '0.05em' }}>¿A qué área va dirigido?</label>
        <div className="grid sm:grid-cols-3 gap-2 mt-2">
          {Object.entries(TIPOS).map(([key, t]) => {
            const active = tipo === key;
            return (
              <button type="button" key={key} onClick={() => setTipo(key)} className="text-left p-3 rounded-lg transition" style={{ background: active ? t.bg : 'var(--bg)', border: `1.5px solid ${active ? t.color : 'var(--line)'}` }}>
                <t.Icon size={17} color={active ? t.color : 'var(--muted)'} />
                <p className="text-sm font-semibold" style={{ color: active ? t.color : 'var(--ink)', marginTop: 6 }}>{t.label}</p>
              </button>
            );
          })}
        </div>
      </div>

      <div>
        <label className="text-xs font-semibold uppercase" style={{ color: 'var(--muted)', letterSpacing: '0.05em' }}>Título del problema o solicitud</label>
        <input value={titulo} onChange={e => setTitulo(e.target.value)} placeholder="Ej. La impresora del 2do piso no enciende" className="input-field mt-1.5" />
        {tocado && !titulo.trim() && <p className="text-xs mt-1" style={{ color: PRIORIDADES.urgente.color }}>Escribe un título breve.</p>}
      </div>

      <div>
        <label className="text-xs font-semibold uppercase" style={{ color: 'var(--muted)', letterSpacing: '0.05em' }}>Descripción</label>
        <textarea value={descripcion} onChange={e => setDescripcion(e.target.value)} rows={4} placeholder="Describe el problema con el mayor detalle posible: qué pasó, desde cuándo, a quién afecta…" className="input-field mt-1.5 resize-none" />
        {tocado && !descripcion.trim() && <p className="text-xs mt-1" style={{ color: PRIORIDADES.urgente.color }}>Agrega una descripción.</p>}
      </div>

      <div>
        <label className="text-xs font-semibold uppercase" style={{ color: 'var(--muted)', letterSpacing: '0.05em' }}>Prioridad</label>
        <div className="flex gap-2 mt-1.5 flex-wrap">
          {Object.entries(PRIORIDADES).map(([key, p]) => {
            const active = prioridad === key;
            return (
              <button type="button" key={key} onClick={() => setPrioridad(key)} className="px-3 py-1.5 rounded-full text-xs font-semibold uppercase transition" style={{ background: active ? p.color : 'transparent', color: active ? '#fff' : p.color, border: `1.5px solid ${p.color}`, letterSpacing: '0.03em' }}>
                {p.label}
              </button>
            );
          })}
        </div>
      </div>

      <div>
        <label className="text-xs font-semibold uppercase flex items-center gap-1.5" style={{ color: 'var(--muted)', letterSpacing: '0.05em' }}>
          <Paperclip size={13} /> Evidencia (opcional)
        </label>
        <div className="mt-2">
          <SelectorAdjuntos onSeleccionar={nuevos => setAdjuntos(prev => [...prev, ...nuevos])} restantes={MAX_ADJUNTOS_TICKET - adjuntos.length} />
        </div>
        {adjuntos.length > 0 && (
          <div className="flex flex-col gap-2 mt-2">
            {adjuntos.map((a, i) => (
              <AdjuntoItem key={i} adjunto={a} onQuitar={() => setAdjuntos(prev => prev.filter((_, idx) => idx !== i))} />
            ))}
          </div>
        )}
      </div>

      <div className="flex items-center justify-between pt-1" style={{ borderTop: '1px dashed var(--line)' }}>
        <p className="text-xs" style={{ color: 'var(--muted)', paddingTop: 12 }}>El equipo de {TIPOS[tipo].label} recibirá este ticket como <b>pendiente</b>.</p>
        <button type="submit" className="btn-primary" style={{ marginTop: 12 }}><Plus size={15} /> Crear ticket</button>
      </div>
    </form>
  );
}

/* ---------- Lista + filtros ---------- */

function TicketsList({ tickets, onOpen, filtroTipo, setFiltroTipo }) {
  const [filtroEstado, setFiltroEstado] = useState('todos');
  const [filtroPrioridad, setFiltroPrioridad] = useState('todos');
  const [q, setQ] = useState('');

  const filtrados = tickets
    .filter(t => filtroTipo === 'todos' || t.tipo === filtroTipo)
    .filter(t => filtroEstado === 'todos' || t.estado === filtroEstado)
    .filter(t => filtroPrioridad === 'todos' || t.prioridad === filtroPrioridad)
    .filter(t => {
      if (!q.trim()) return true;
      const s = q.toLowerCase();
      return t.titulo.toLowerCase().includes(s) || t.id.toLowerCase().includes(s) || t.area.toLowerCase().includes(s) || t.solicitante.toLowerCase().includes(s);
    })
    .sort((a, b) => new Date(b.creado) - new Date(a.creado));

  return (
    <div className="flex flex-col gap-4">
      <div className="stub-flat rounded-lg p-3 flex flex-wrap gap-2 items-center" style={{ background: 'var(--card)', border: '1px solid var(--line)' }}>
        <div className="relative" style={{ flex: '1 1 200px' }}>
          <Search size={14} style={{ position: 'absolute', left: 10, top: 9, color: 'var(--muted)' }} />
          <input value={q} onChange={e => setQ(e.target.value)} placeholder="Buscar por título, ID, área o solicitante…" className="input-field" style={{ paddingLeft: 30, height: 34 }} />
        </div>
        <select value={filtroTipo} onChange={e => setFiltroTipo(e.target.value)} className="select-field">
          <option value="todos">Todas las áreas</option>
          {Object.entries(TIPOS).map(([k, t]) => <option key={k} value={k}>{t.label}</option>)}
        </select>
        <select value={filtroEstado} onChange={e => setFiltroEstado(e.target.value)} className="select-field">
          <option value="todos">Todos los estados</option>
          {ESTADO_ORDEN.map(k => <option key={k} value={k}>{ESTADOS[k].label}</option>)}
        </select>
        <select value={filtroPrioridad} onChange={e => setFiltroPrioridad(e.target.value)} className="select-field">
          <option value="todos">Toda prioridad</option>
          {Object.entries(PRIORIDADES).map(([k, p]) => <option key={k} value={k}>{p.label}</option>)}
        </select>
        <button
          onClick={() => exportarTicketsExcel(filtrados)}
          disabled={filtrados.length === 0}
          className="btn-secondary flex items-center gap-1.5"
          style={{ padding: '8px 12px', opacity: filtrados.length === 0 ? 0.5 : 1 }}
          title="Exportar los tickets filtrados a Excel"
        >
          <Download size={14} /> Exportar a Excel
        </button>
      </div>

      {filtrados.length === 0 ? (
        <div className="rounded-lg" style={{ background: 'var(--card)', border: '1px solid var(--line)' }}>
          <EmptyState texto="No hay tickets con estos filtros" sub="Prueba a cambiar los filtros o la búsqueda." />
        </div>
      ) : (
        <div className="flex flex-col gap-2.5">
          {filtrados.map(t => <TicketRow key={t.id} ticket={t} onOpen={onOpen} />)}
        </div>
      )}
    </div>
  );
}

/* ---------- Modal de detalle ---------- */

function TicketModal({ ticket, onClose, onUpdate, onAddNota, onAddAdjuntos, onDelete, esAdmin, sessionUser }) {
  const [nota, setNota] = useState('');
  const [asignado, setAsignado] = useState(ticket.asignadoA || '');
  const tipo = TIPOS[ticket.tipo];
  const estado = ESTADOS[ticket.estado];
  const prioridad = PRIORIDADES[ticket.prioridad];
  const readOnly = !esAdmin;
  const puedeComentar = esAdmin || ticket.creadoPor === sessionUser.email;

  function enviarNota() {
    if (!nota.trim()) return;
    onAddNota(ticket.id, nota.trim());
    setNota('');
  }

  function confirmarBorrado() {
    if (window.confirm(`¿Eliminar el ticket ${ticket.id}? Esta acción no se puede deshacer.`)) onDelete(ticket.id);
  }

  return (
    <div className="fixed inset-0 z-40 flex items-center justify-center p-4" style={{ background: 'rgba(27,36,48,0.45)' }} onClick={onClose}>
      <div className="w-full max-w-2xl rounded-xl flex flex-col" style={{ background: 'var(--card)', maxHeight: '90vh' }} onClick={e => e.stopPropagation()}>
        <div className="flex items-center justify-between px-5 py-4" style={{ borderBottom: '1px solid var(--line)' }}>
          <div className="flex items-center gap-2">
            <TipoBadge tipo={ticket.tipo} />
            <span className="font-mono text-xs" style={{ color: 'var(--muted)' }}>{ticket.id}</span>
            {readOnly && <Pill label="Vista de usuario" color="#7C93A6" Icon={Eye} />}
          </div>
          <button onClick={onClose} className="icon-btn"><X size={17} /></button>
        </div>

        <div className="px-5 py-4 overflow-y-auto flex flex-col gap-5">
          <div>
            <h2 className="font-display font-semibold" style={{ fontSize: 19, color: 'var(--ink)' }}>{ticket.titulo}</h2>
            <p className="text-sm mt-2" style={{ color: 'var(--ink)', opacity: 0.85, whiteSpace: 'pre-wrap' }}>{ticket.descripcion}</p>
          </div>

          <div className="grid sm:grid-cols-2 gap-3 text-sm">
            <div className="flex items-center gap-2"><Building2 size={14} style={{ color: 'var(--muted)' }} /><span style={{ color: 'var(--muted)' }}>Área:</span><span style={{ color: 'var(--ink)' }}>{ticket.area}</span></div>
            <div className="flex items-center gap-2"><User size={14} style={{ color: 'var(--muted)' }} /><span style={{ color: 'var(--muted)' }}>Solicitante:</span><span style={{ color: 'var(--ink)' }}>{ticket.solicitante}</span></div>
            <div className="text-xs" style={{ color: 'var(--muted)' }}>Creado: {formatFecha(ticket.creado)}</div>
            <div className="text-xs" style={{ color: 'var(--muted)' }}>Actualizado: {formatFecha(ticket.actualizado)}</div>
          </div>

          <div>
            <label className="text-xs font-semibold uppercase" style={{ color: 'var(--muted)', letterSpacing: '0.05em' }}>Asignado a</label>
            {readOnly ? (
              <p className="text-sm mt-1.5" style={{ color: ticket.asignadoA ? 'var(--ink)' : 'var(--muted)' }}>{ticket.asignadoA || 'Sin asignar'}</p>
            ) : (
              <input value={asignado} onChange={e => setAsignado(e.target.value)} onBlur={() => onUpdate(ticket.id, { asignadoA: asignado.trim() })} placeholder="Sin asignar" className="input-field mt-1.5" />
            )}
          </div>

          <div>
            <label className="text-xs font-semibold uppercase" style={{ color: 'var(--muted)', letterSpacing: '0.05em' }}>Prioridad</label>
            {readOnly ? (
              <div className="mt-1.5"><Pill label={prioridad.label} color={prioridad.color} uppercase /></div>
            ) : (
              <div className="flex gap-2 mt-1.5 flex-wrap">
                {Object.entries(PRIORIDADES).map(([key, p]) => {
                  const active = ticket.prioridad === key;
                  return (
                    <button key={key} onClick={() => onUpdate(ticket.id, { prioridad: key })} className="px-3 py-1.5 rounded-full text-xs font-semibold uppercase transition" style={{ background: active ? p.color : 'transparent', color: active ? '#fff' : p.color, border: `1.5px solid ${p.color}` }}>
                      {p.label}
                    </button>
                  );
                })}
              </div>
            )}
          </div>

          <div>
            <label className="text-xs font-semibold uppercase" style={{ color: 'var(--muted)', letterSpacing: '0.05em' }}>Estado</label>
            {readOnly ? (
              <div className="mt-1.5"><Pill label={estado.label} color={estado.color} Icon={estado.Icon} /></div>
            ) : (
              <div className="flex gap-2 mt-1.5 flex-wrap">
                {ESTADO_ORDEN.map(key => {
                  const e = ESTADOS[key];
                  const active = ticket.estado === key;
                  return (
                    <button key={key} onClick={() => onUpdate(ticket.id, { estado: key })} className="flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-semibold transition" style={{ background: active ? e.color : 'transparent', color: active ? '#fff' : e.color, border: `1.5px solid ${e.color}` }}>
                      <e.Icon size={13} /> {e.label}
                    </button>
                  );
                })}
              </div>
            )}
          </div>

          <div>
            <label className="text-xs font-semibold uppercase flex items-center gap-1.5" style={{ color: 'var(--muted)', letterSpacing: '0.05em' }}>
              <MessageSquare size={13} /> Notas de seguimiento
            </label>
            <div className="flex flex-col gap-2 mt-2 max-h-40 overflow-y-auto pr-1">
              {ticket.notas.length === 0 && <p className="text-xs" style={{ color: 'var(--muted)' }}>Sin notas todavía.</p>}
              {ticket.notas.map((n, i) => (
                <div key={i} className="text-sm p-2.5 rounded-md" style={{ background: 'var(--bg)' }}>
                  <p style={{ color: 'var(--ink)' }}>{n.texto}</p>
                  <p className="text-xs mt-1" style={{ color: 'var(--muted)' }}>
                    {n.autorNombre ? `${n.autorNombre} · ` : ''}{formatFecha(n.fecha)}
                  </p>
                </div>
              ))}
            </div>
            {puedeComentar && (
              <div className="flex gap-2 mt-2">
                <input value={nota} onChange={e => setNota(e.target.value)} onKeyDown={e => { if (e.key === 'Enter') enviarNota(); }} placeholder="Escribe una actualización o comentario…" className="input-field" style={{ flex: 1 }} />
                <button onClick={enviarNota} className="icon-btn" style={{ background: tipo.color, color: '#fff' }}><Send size={15} /></button>
              </div>
            )}
          </div>

          <div>
            <label className="text-xs font-semibold uppercase flex items-center gap-1.5" style={{ color: 'var(--muted)', letterSpacing: '0.05em' }}>
              <Paperclip size={13} /> Evidencia adjunta
            </label>
            <div className="flex flex-col gap-2 mt-2">
              {ticket.adjuntos.length === 0 && <p className="text-xs" style={{ color: 'var(--muted)' }}>Sin archivos adjuntos.</p>}
              {ticket.adjuntos.map(a => <AdjuntoItem key={a.id} adjunto={a} />)}
            </div>
            {puedeComentar && (
              <div className="mt-2">
                <SelectorAdjuntos
                  onSeleccionar={nuevos => onAddAdjuntos(ticket.id, nuevos)}
                  restantes={MAX_ADJUNTOS_TICKET - ticket.adjuntos.length}
                />
              </div>
            )}
          </div>
        </div>

        <div className="flex items-center justify-between px-5 py-3" style={{ borderTop: '1px solid var(--line)' }}>
          {readOnly ? (
            <p className="text-xs" style={{ color: 'var(--muted)' }}>Puedes comentar; el estado, prioridad y asignación los maneja el equipo de TI y Mantenimiento.</p>
          ) : (
            <button onClick={confirmarBorrado} className="flex items-center gap-1.5 text-xs font-medium" style={{ color: PRIORIDADES.urgente.color }}>
              <Trash2 size={13} /> Eliminar ticket
            </button>
          )}
          <button onClick={onClose} className="btn-secondary">Cerrar</button>
        </div>
      </div>
    </div>
  );
}

/* ---------- Panel de usuarios (solo admin) ---------- */

function UsersPanel({ usuarios, onCambiarRol, onEditar, onEliminar, sesionEmail }) {
  const [editando, setEditando] = useState(null);
  const [nombreForm, setNombreForm] = useState('');
  const [areaForm, setAreaForm] = useState('');
  const totalAdmins = usuarios.filter(u => u.rol === 'admin').length;

  function empezarEdicion(u) {
    setEditando(u.email);
    setNombreForm(u.nombre);
    setAreaForm(u.area);
  }

  function cancelarEdicion() {
    setEditando(null);
  }

  function guardarEdicion(email) {
    if (!nombreForm.trim() || !areaForm.trim()) return;
    onEditar(email, { nombre: nombreForm.trim(), area: areaForm.trim() });
    setEditando(null);
  }

  function confirmarEliminar(u) {
    if (window.confirm(`¿Eliminar la cuenta de ${u.nombre} (${u.email})? Sus tickets no se borrarán, pero ya no podrá iniciar sesión.`)) {
      onEliminar(u.email);
    }
  }

  return (
    <div className="rounded-lg overflow-hidden" style={{ background: 'var(--card)', border: '1px solid var(--line)' }}>
      <div className="px-4 py-3" style={{ borderBottom: '1px solid var(--line)' }}>
        <h2 className="font-display font-semibold" style={{ fontSize: 15, color: 'var(--ink)' }}>Cuentas registradas</h2>
        <p className="text-xs" style={{ color: 'var(--muted)' }}>{usuarios.length} cuenta{usuarios.length === 1 ? '' : 's'} con dominio {DOMINIO}</p>
      </div>
      <div className="flex flex-col">
        {usuarios.map(u => {
          const esAdminFila = u.rol === 'admin';
          const esUnoMismo = u.email === sesionEmail;
          const esUltimoAdmin = esAdminFila && totalAdmins <= 1;
          const enEdicion = editando === u.email;

          if (enEdicion) {
            return (
              <div key={u.email} className="flex flex-col gap-2 px-4 py-3" style={{ borderBottom: '1px solid var(--line)', background: 'var(--bg)' }}>
                <p className="text-xs truncate" style={{ color: 'var(--muted)' }}>{u.email}</p>
                <div className="grid sm:grid-cols-2 gap-2">
                  <div>
                    <label className="text-xs font-semibold uppercase" style={{ color: 'var(--muted)', letterSpacing: '0.05em' }}>Nombre</label>
                    <input value={nombreForm} onChange={e => setNombreForm(e.target.value)} className="input-field mt-1" style={{ height: 34, background: 'var(--card)' }} />
                  </div>
                  <div>
                    <label className="text-xs font-semibold uppercase" style={{ color: 'var(--muted)', letterSpacing: '0.05em' }}>Área</label>
                    <input value={areaForm} onChange={e => setAreaForm(e.target.value)} className="input-field mt-1" style={{ height: 34, background: 'var(--card)' }} />
                  </div>
                </div>
                <div className="flex items-center gap-2 justify-end mt-1">
                  <button onClick={cancelarEdicion} className="btn-secondary" style={{ padding: '6px 10px', fontSize: 12 }}>Cancelar</button>
                  <button onClick={() => guardarEdicion(u.email)} className="btn-primary" style={{ padding: '6px 12px', fontSize: 12 }}>Guardar</button>
                </div>
              </div>
            );
          }

          return (
            <div key={u.email} className="flex items-center justify-between gap-3 px-4 py-3 flex-wrap" style={{ borderBottom: '1px solid var(--line)' }}>
              <div className="min-w-0">
                <p className="text-sm font-medium truncate" style={{ color: 'var(--ink)' }}>{u.nombre} {esUnoMismo && <span style={{ color: 'var(--muted)' }}>(tú)</span>}</p>
                <p className="text-xs truncate" style={{ color: 'var(--muted)' }}>{u.email} · {u.area}</p>
              </div>
              <div className="flex items-center gap-2 flex-wrap">
                <Pill label={esAdminFila ? 'Admin' : 'Usuario'} color={esAdminFila ? '#6B4E9E' : '#7C93A6'} Icon={esAdminFila ? ShieldCheck : User} uppercase />
                <button onClick={() => onCambiarRol(u.email)} className="btn-secondary" style={{ padding: '6px 10px', fontSize: 12 }}>
                  {esAdminFila ? 'Quitar admin' : 'Hacer admin'}
                </button>
                <button onClick={() => empezarEdicion(u)} className="icon-btn" title="Editar cuenta" style={{ width: 30, height: 30 }}>
                  <Pencil size={14} />
                </button>
                <button
                  onClick={() => confirmarEliminar(u)}
                  disabled={esUnoMismo || esUltimoAdmin}
                  className="icon-btn"
                  title={esUnoMismo ? 'No puedes eliminar tu propia cuenta' : esUltimoAdmin ? 'Debe existir al menos un administrador' : 'Eliminar cuenta'}
                  style={{ width: 30, height: 30, opacity: (esUnoMismo || esUltimoAdmin) ? 0.4 : 1, color: (esUnoMismo || esUltimoAdmin) ? 'var(--muted)' : PRIORIDADES.urgente.color }}
                >
                  <Trash2 size={14} />
                </button>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ---------- App raíz ---------- */

export default function App() {
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [tickets, setTickets] = useState([]);
  const [seq, setSeq] = useState(1);
  const [usuarios, setUsuarios] = useState([]);
  const [notificaciones, setNotificaciones] = useState([]);
  const [sessionUser, setSessionUser] = useState(null);
  const [view, setView] = useState('dashboard');
  const [selectedId, setSelectedId] = useState(null);
  const [toast, setToast] = useState(null);
  const [filtroTipo, setFiltroTipo] = useState('todos');

  useEffect(() => {
    (async () => {
      try {
        const [resTickets, resUsers, resNotif] = await Promise.all([
          window.storage.get(STORAGE_TICKETS, true).catch(() => null),
          window.storage.get(STORAGE_USERS, true).catch(() => null),
          window.storage.get(STORAGE_NOTIFICACIONES, true).catch(() => null),
        ]);
        if (resTickets && resTickets.value) {
          const parsed = JSON.parse(resTickets.value);
          setTickets((parsed.tickets || []).map(t => ({ notas: [], adjuntos: [], ...t })));
          setSeq(parsed.seq || 1);
        }
        if (resNotif && resNotif.value) {
          setNotificaciones(JSON.parse(resNotif.value) || []);
        }
        if (resUsers && resUsers.value) {
          const lista = JSON.parse(resUsers.value);
          setUsuarios(lista.length ? lista : [{ ...ADMIN_SEED, creado: new Date().toISOString() }]);
          if (!lista.length) await window.storage.set(STORAGE_USERS, JSON.stringify([{ ...ADMIN_SEED, creado: new Date().toISOString() }]), true);
        } else {
          const seed = [{ ...ADMIN_SEED, creado: new Date().toISOString() }];
          setUsuarios(seed);
          await window.storage.set(STORAGE_USERS, JSON.stringify(seed), true);
        }
      } catch (e) {
        setUsuarios([{ ...ADMIN_SEED, creado: new Date().toISOString() }]);
      } finally {
        setLoading(false);
      }
    })();
  }, []);

  async function persistTickets(newTickets, newSeq) {
    try {
      const ok = await window.storage.set(STORAGE_TICKETS, JSON.stringify({ tickets: newTickets, seq: newSeq }), true);
      setError(ok ? null : 'No se pudo guardar el cambio. Intenta de nuevo.');
    } catch (e) { setError('No se pudo guardar el cambio. Intenta de nuevo.'); }
  }

  async function persistUsuarios(lista) {
    try { await window.storage.set(STORAGE_USERS, JSON.stringify(lista), true); }
    catch (e) { setError('No se pudo guardar el cambio de usuarios.'); }
  }

  async function persistNotificaciones(lista) {
    try { await window.storage.set(STORAGE_NOTIFICACIONES, JSON.stringify(lista), true); }
    catch (e) { setError('No se pudo guardar la notificación.'); }
  }

  function showToast(msg) { setToast(msg); setTimeout(() => setToast(null), 2400); }

  function crearNotificacion({ para, ticketId, tipo, mensaje }) {
    if (!para) return;
    const notif = {
      id: `N-${nuevoId(seq)}-${para}-${Math.random().toString(36).slice(2, 7)}`,
      para, ticketId, tipo, mensaje, fecha: new Date().toISOString(), leida: false,
    };
    setNotificaciones(prev => {
      const nuevas = [notif, ...prev];
      persistNotificaciones(nuevas);
      return nuevas;
    });
    enviarNotificacionEmail(para, `Central de Tickets · ${ticketId}`, mensaje);
  }

  function notificarCambioEstado(ticketAnterior, nuevoEstado) {
    if (!ticketAnterior || ticketAnterior.creadoPor === sessionUser.email) return;
    crearNotificacion({
      para: ticketAnterior.creadoPor,
      ticketId: ticketAnterior.id,
      tipo: 'estado',
      mensaje: `Tu ticket ${ticketAnterior.id} (“${ticketAnterior.titulo}”) cambió a “${ESTADOS[nuevoEstado].label}”.`,
    });
  }

  function notificarNuevaNota(ticket, autorEmail, autorNombre) {
    const autorEsAdmin = usuarios.find(u => u.email === autorEmail)?.rol === 'admin';
    if (autorEsAdmin) {
      if (ticket.creadoPor !== autorEmail) {
        crearNotificacion({
          para: ticket.creadoPor,
          ticketId: ticket.id,
          tipo: 'comentario',
          mensaje: `Nuevo comentario de ${autorNombre} en tu ticket ${ticket.id}.`,
        });
      }
    } else {
      usuarios.filter(u => u.rol === 'admin' && u.email !== autorEmail).forEach(admin => {
        crearNotificacion({
          para: admin.email,
          ticketId: ticket.id,
          tipo: 'comentario',
          mensaje: `${autorNombre} comentó en el ticket ${ticket.id} (“${ticket.titulo}”).`,
        });
      });
    }
  }

  function marcarNotifLeida(id) {
    setNotificaciones(prev => {
      const nuevas = prev.map(n => n.id === id ? { ...n, leida: true } : n);
      persistNotificaciones(nuevas);
      return nuevas;
    });
  }

  function marcarTodasNotifLeidas() {
    setNotificaciones(prev => {
      const nuevas = prev.map(n => n.para === sessionUser.email ? { ...n, leida: true } : n);
      persistNotificaciones(nuevas);
      return nuevas;
    });
  }

  function abrirNotificacion(n) {
    marcarNotifLeida(n.id);
    setSelectedId(n.ticketId);
  }

  function intentarLogin(email, password) {
    const e = (email || '').trim().toLowerCase();
    if (!validarCorreo(e)) return { ok: false, error: 'Usa tu correo corporativo (@genesisig.com).' };
    const u = usuarios.find(x => x.email.toLowerCase() === e);
    if (!u) return { ok: false, error: 'No existe una cuenta con ese correo.' };
    if (u.password !== password) return { ok: false, error: 'Contraseña incorrecta.' };
    setSessionUser({ email: u.email, nombre: u.nombre, area: u.area, rol: u.rol });
    setView('dashboard');
    return { ok: true };
  }

  function registrar({ email, password, nombre, area }) {
    const e = (email || '').trim().toLowerCase();
    if (!validarCorreo(e)) return { ok: false, error: 'Usa tu correo corporativo (@genesisig.com).' };
    if (!nombre.trim() || !area.trim() || !password) return { ok: false, error: 'Completa todos los campos.' };
    const chequeoPw = validarPassword(password);
    if (!chequeoPw.ok) return { ok: false, error: chequeoPw.error };
    if (usuarios.some(u => u.email.toLowerCase() === e)) return { ok: false, error: 'Ya existe una cuenta con ese correo.' };
    const nuevo = { email: e, nombre: nombre.trim(), area: area.trim(), password, rol: 'usuario', creado: new Date().toISOString() };
    const nuevos = [...usuarios, nuevo];
    setUsuarios(nuevos);
    persistUsuarios(nuevos);
    setSessionUser({ email: nuevo.email, nombre: nuevo.nombre, area: nuevo.area, rol: nuevo.rol });
    setView('dashboard');
    return { ok: true };
  }

  function cerrarSesion() { setSessionUser(null); setSelectedId(null); setView('dashboard'); }

  function cambiarRol(email) {
    const objetivo = usuarios.find(u => u.email === email);
    if (!objetivo) return;
    const nuevoRol = objetivo.rol === 'admin' ? 'usuario' : 'admin';
    if (objetivo.rol === 'admin' && nuevoRol === 'usuario' && usuarios.filter(u => u.rol === 'admin').length <= 1) {
      showToast('Debe existir al menos un administrador.');
      return;
    }
    const nuevos = usuarios.map(u => u.email === email ? { ...u, rol: nuevoRol } : u);
    setUsuarios(nuevos);
    persistUsuarios(nuevos);
    showToast(`${objetivo.nombre} ahora es ${nuevoRol === 'admin' ? 'admin' : 'usuario'}`);
    if (sessionUser && sessionUser.email === email) setSessionUser({ ...sessionUser, rol: nuevoRol });
  }

  function editarUsuario(email, patch) {
    const objetivo = usuarios.find(u => u.email === email);
    if (!objetivo) return;
    const nuevos = usuarios.map(u => u.email === email ? { ...u, ...patch } : u);
    setUsuarios(nuevos);
    persistUsuarios(nuevos);
    showToast('Cuenta actualizada');
    if (sessionUser && sessionUser.email === email) setSessionUser({ ...sessionUser, ...patch });
  }

  function eliminarUsuario(email) {
    const objetivo = usuarios.find(u => u.email === email);
    if (!objetivo) return;
    if (objetivo.rol === 'admin' && usuarios.filter(u => u.rol === 'admin').length <= 1) {
      showToast('Debe existir al menos un administrador.');
      return;
    }
    if (sessionUser && sessionUser.email === email) {
      showToast('No puedes eliminar tu propia cuenta mientras tienes sesión iniciada.');
      return;
    }
    const nuevos = usuarios.filter(u => u.email !== email);
    setUsuarios(nuevos);
    persistUsuarios(nuevos);
    showToast(`Cuenta de ${objetivo.nombre} eliminada`);
  }

  function crearTicket(data) {
    const id = nuevoId(seq);
    const ahora = new Date().toISOString();
    const adjuntosIniciales = (data.adjuntos || []).map((a, i) => ({
      ...a, id: `${id}-A${i}`, subidoPor: sessionUser.email, subidoNombre: sessionUser.nombre, fecha: ahora,
    }));
    const ticket = {
      id, tipo: data.tipo, titulo: data.titulo, descripcion: data.descripcion, prioridad: data.prioridad,
      area: sessionUser.area, solicitante: sessionUser.nombre, creadoPor: sessionUser.email,
      estado: 'pendiente', asignadoA: '', creado: ahora, actualizado: ahora, notas: [], adjuntos: adjuntosIniciales,
    };
    const nuevos = [ticket, ...tickets];
    const nuevoSeq = seq + 1;
    setTickets(nuevos); setSeq(nuevoSeq);
    persistTickets(nuevos, nuevoSeq);
    showToast(`Ticket ${id} creado`);
    setSelectedId(id);
    setView('tickets');
  }

  function actualizarTicket(id, patch) {
    const anterior = tickets.find(t => t.id === id);
    const nuevos = tickets.map(t => t.id === id ? { ...t, ...patch, actualizado: new Date().toISOString() } : t);
    setTickets(nuevos);
    persistTickets(nuevos, seq);
    if (anterior && patch.estado && patch.estado !== anterior.estado) {
      notificarCambioEstado(anterior, patch.estado);
    }
  }

  function agregarNota(id, texto) {
    const ticket = tickets.find(t => t.id === id);
    const nuevos = tickets.map(t => t.id === id ? {
      ...t,
      notas: [...t.notas, { texto, fecha: new Date().toISOString(), autor: sessionUser.email, autorNombre: sessionUser.nombre }],
      actualizado: new Date().toISOString(),
    } : t);
    setTickets(nuevos);
    persistTickets(nuevos, seq);
    if (ticket) notificarNuevaNota(ticket, sessionUser.email, sessionUser.nombre);
  }

  function agregarAdjuntos(id, nuevosArchivos) {
    const ahora = new Date().toISOString();
    const nuevos = tickets.map(t => t.id === id ? {
      ...t,
      adjuntos: [
        ...t.adjuntos,
        ...nuevosArchivos.map((a, i) => ({
          ...a, id: `${id}-A${t.adjuntos.length + i}-${Date.now()}`, subidoPor: sessionUser.email, subidoNombre: sessionUser.nombre, fecha: ahora,
        })),
      ],
      actualizado: ahora,
    } : t);
    setTickets(nuevos);
    persistTickets(nuevos, seq);
  }

  function eliminarTicket(id) {
    const nuevos = tickets.filter(t => t.id !== id);
    setTickets(nuevos);
    persistTickets(nuevos, seq);
    showToast('Ticket eliminado');
  }

  const esAdmin = sessionUser && sessionUser.rol === 'admin';
  const ticketsVisibles = !sessionUser ? [] : esAdmin ? tickets : tickets.filter(t => t.creadoPor === sessionUser.email);
  const seleccionado = tickets.find(t => t.id === selectedId) || null;
  const notifsPropias = !sessionUser ? [] : notificaciones.filter(n => n.para === sessionUser.email).sort((a, b) => new Date(b.fecha) - new Date(a.fecha));

  const estilosGlobales = (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Oswald:wght@500;600;700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap');
      :root { --bg: #EDEAE2; --card: #FFFFFF; --ink: #1B2430; --muted: #7A7E86; --line: #DAD6CC; }
      .font-display { font-family: 'Oswald', sans-serif; }
      .font-mono { font-family: 'IBM Plex Mono', monospace; }
      .stub-flat { box-shadow: 0 1px 2px rgba(27,36,48,0.05); }
      .input-field, .select-field { width: 100%; box-sizing: border-box; height: 38px; padding: 0 10px; border: 1px solid var(--line); border-radius: 8px; background: var(--bg); color: var(--ink); font-size: 13.5px; font-family: 'Inter', sans-serif; outline: none; }
      textarea.input-field { height: auto; padding: 8px 10px; }
      .input-field:focus, .select-field:focus { border-color: var(--ink); }
      .select-field { width: auto; cursor: pointer; }
      .btn-primary { display: inline-flex; align-items: center; gap: 6px; background: var(--ink); color: #fff; font-size: 13.5px; font-weight: 600; padding: 9px 16px; border-radius: 8px; transition: opacity .15s; }
      .btn-primary:hover { opacity: 0.88; }
      .btn-secondary { font-size: 13.5px; font-weight: 600; padding: 8px 14px; border-radius: 8px; border: 1px solid var(--line); color: var(--ink); background: transparent; }
      .icon-btn { display: inline-flex; align-items: center; justify-content: center; width: 32px; height: 32px; border-radius: 8px; color: var(--muted); background: var(--bg); flex-shrink: 0; }
      .icon-btn:hover { opacity: 0.85; }
      .ticket-row { display: flex; position: relative; background: var(--card); border-radius: 10px; border: 1px solid var(--line); cursor: pointer; transition: transform .12s, box-shadow .12s; }
      .ticket-row:hover { transform: translateY(-1px); box-shadow: 0 4px 14px rgba(27,36,48,0.08); }
      .stub-col { width: 78px; flex-shrink: 0; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 4px; padding: 10px 4px; position: relative; border-right: 2px dashed rgba(27,36,48,0.16); border-radius: 10px 0 0 10px; }
      .notch-top, .notch-bottom { position: absolute; right: -7px; width: 14px; height: 14px; border-radius: 50%; background: var(--bg); z-index: 2; }
      .notch-top { top: -1px; }
      .notch-bottom { bottom: -1px; }
      @media (prefers-reduced-motion: reduce) { .ticket-row { transition: none; } }
    `}</style>
  );

  if (loading) {
    return (
      <div style={{ minHeight: '100vh', background: 'var(--bg)', fontFamily: "'Inter', sans-serif" }}>
        {estilosGlobales}
        <div className="flex items-center justify-center py-24 text-sm" style={{ color: 'var(--muted)' }}>Cargando…</div>
      </div>
    );
  }

  if (!sessionUser) {
    return (
      <div style={{ minHeight: '100vh', background: 'var(--bg)', fontFamily: "'Inter', sans-serif" }}>
        {estilosGlobales}
        <LoginScreen onLogin={intentarLogin} onRegistrar={registrar} />
      </div>
    );
  }

  return (
    <div style={{ minHeight: '100vh', background: 'var(--bg)', fontFamily: "'Inter', sans-serif" }}>
      {estilosGlobales}
      <Header
        view={view} setView={setView} total={ticketsVisibles.length} user={sessionUser} onLogout={cerrarSesion} esAdmin={esAdmin}
        notificaciones={notifsPropias} onAbrirNotif={abrirNotificacion} onMarcarTodasNotif={marcarTodasNotifLeidas}
      />

      <main className="max-w-6xl mx-auto px-4 md:px-6 py-6">
        {error && (
          <div className="flex items-center gap-2 mb-4 p-3 rounded-lg text-sm" style={{ background: '#FBEAEA', color: PRIORIDADES.urgente.color, border: `1px solid ${PRIORIDADES.urgente.color}40` }}>
            <AlertTriangle size={15} /> {error}
          </div>
        )}

        {view === 'dashboard' ? (
          <Dashboard tickets={ticketsVisibles} onOpen={t => setSelectedId(t.id)} onNav={setView} onFiltrarTipo={setFiltroTipo} esAdmin={esAdmin} />
        ) : view === 'nuevo' ? (
          <NuevoTicketForm onCreate={crearTicket} user={sessionUser} />
        ) : view === 'usuarios' && esAdmin ? (
          <UsersPanel usuarios={usuarios} onCambiarRol={cambiarRol} onEditar={editarUsuario} onEliminar={eliminarUsuario} sesionEmail={sessionUser.email} />
        ) : (
          <TicketsList tickets={ticketsVisibles} onOpen={t => setSelectedId(t.id)} filtroTipo={filtroTipo} setFiltroTipo={setFiltroTipo} />
        )}
      </main>

      {view !== 'nuevo' && (
        <button onClick={() => setView('nuevo')} className="fixed flex items-center justify-center rounded-full shadow-xl md:hidden" style={{ bottom: 20, right: 20, width: 52, height: 52, background: 'var(--ink)', color: '#fff' }}>
          <Plus size={22} />
        </button>
      )}

      {seleccionado && (
        <TicketModal
          ticket={seleccionado}
          onClose={() => setSelectedId(null)}
          onUpdate={actualizarTicket}
          onAddNota={agregarNota}
          onAddAdjuntos={agregarAdjuntos}
          onDelete={eliminarTicket}
          esAdmin={esAdmin}
          sessionUser={sessionUser}
        />
      )}

      {toast && <Toast message={toast} />}
    </div>
  );
}
