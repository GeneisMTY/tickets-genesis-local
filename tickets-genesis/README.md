# Central de Tickets · Genesis (proyecto local)

Versión de escritorio del sistema de tickets, para correr con Node.js en tu propia PC.

## Requisitos
- Tener instalado **Node.js** (18 o más reciente). Descárgalo de https://nodejs.org si no lo tienes.

## Cómo arrancarlo

1. Descomprime esta carpeta donde quieras (ej. `Escritorio\Tickers Genesis`).
2. Abre una terminal **dentro de esa carpeta** (la que sí contiene `package.json`).
3. Instala las dependencias (solo la primera vez):
   ```
   npm install
   ```
4. Levanta la app:
   ```
   npm run dev
   ```
5. Abre en tu navegador la dirección que te muestre la terminal (normalmente `http://localhost:5173`).

Cuenta de prueba para entrar: `admin@genesisig.com` / `Genesis2026`.

## ⚠️ Importante: esto guarda los datos solo en este navegador

Este proyecto usa `localStorage` del navegador para guardar tickets y cuentas — es decir,
**todo se queda en esta computadora y en este navegador**. Si otra persona abre la app desde
otra PC, va a ver una lista vacía y sus propias cuentas: no hay sincronización entre equipos.

Esto es perfecto para probarlo tú mismo o hacer una demo local, pero **no sirve todavía
como sistema real para que Ventas, RRHH, etc. manden tickets y que TI los vea desde otra
computadora**. Para eso hay dos caminos:

- **Más simple:** seguir usando la versión que corre dentro de Claude.ai (la que ya tenías) —
  ahí el guardado sí es compartido entre cualquiera que abra el enlace.
- **Versión real para la empresa:** montar un backend (por ejemplo Node + una base de datos
  como PostgreSQL o SQLite) y un servidor donde todos los empleados puedan entrar por
  navegador — puedo ayudarte a construir eso cuando quieras dar ese paso.

## Estructura del proyecto
```
src/
  App.jsx               → toda la lógica y las pantallas del sistema
  storage-polyfill.js    → guarda los datos en localStorage (ver aviso arriba)
  main.jsx               → punto de entrada de React
  index.css
index.html
package.json
vite.config.js
```

## Nota sobre esta entrega
No pude ejecutar `npm install` desde este chat para probarlo en vivo — el entorno donde yo
corro comandos tuvo un bloqueo de red hacia el registro de npm en este momento. El código y
la configuración siguen el patrón estándar de un proyecto Vite + React, así que debería
instalar sin problema en tu equipo; si al correr `npm install` o `npm run dev` te sale algún
error, pégamelo y lo resolvemos.
