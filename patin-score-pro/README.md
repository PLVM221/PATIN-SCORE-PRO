# Patín Score Pro

Prototipo local para competencias de patín artístico.

## Ejecutar

```bash
npm install
npm run dev
```

En Windows PowerShell puede hacer falta:

```bash
npm.cmd run dev
```

## Ejecutar con SQL

En una terminal:

```bash
npm.cmd run server
```

En otra terminal:

```bash
npm.cmd run dev
```

La base queda en `server/data/patin-score-pro.sqlite`. Si el servidor SQL no está abierto, la app sigue usando `localStorage`.

Para producción local:

```bash
npm.cmd run build
npm.cmd run server
```

## Incluye

- Torneos, categorías, clubes, técnicas, patinadoras y jueces demo.
- Pantalla operador con música, silbato, receso, navegación y doble confirmación.
- Pantalla jueces con puntajes de varios jueces, modo consensuado, bonus y observaciones privadas.
- Pantalla pública/LED.
- Rankings por categoría y clubes.
- Actas imprimibles con firma de jueces y premiación final.
- Registros básicos.
- Persistencia SQL con SQLite y fallback en localStorage.
