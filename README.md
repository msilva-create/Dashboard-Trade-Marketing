# Tracker de Inversiones por Cliente

App web para gestionar inversiones, presupuestos y pendientes por cliente.

## Funcionalidades

- **Dashboard** — KPIs globales, gráfica por mes, donut por canal, tabla por cliente
- **Inversiones** — registra cada inversión con filtros por cliente y mes
- **Presupuesto** — define presupuesto mensual por cliente con comparativo automático
- **Pendientes** — tareas con prioridad, estado y responsable. Click en el círculo para avanzar estado.

Los datos se guardan en el navegador (localStorage) — no necesitas base de datos.

---

## Cómo subir a Vercel (5 minutos)

### Opción A — Desde GitHub (recomendada)

1. Sube esta carpeta a un repositorio en GitHub
2. Ve a [vercel.com](https://vercel.com) → "Add New Project"
3. Importa tu repositorio
4. Vercel detecta Vite automáticamente → clic en **Deploy**
5. ¡Listo! Tienes una URL pública

### Opción B — Desde la terminal

```bash
# Instala Vercel CLI (solo una vez)
npm i -g vercel

# Dentro de esta carpeta
npm install
vercel

# Sigue las instrucciones en pantalla
```

### Opción C — Drag & Drop

```bash
npm install
npm run build
```

Luego arrastra la carpeta `dist/` a [vercel.com/new](https://vercel.com/new) directo en el navegador.

---

## Desarrollo local

```bash
npm install
npm run dev
```

Abre http://localhost:5173
