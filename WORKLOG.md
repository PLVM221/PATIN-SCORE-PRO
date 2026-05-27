# WORKLOG.md

### 2026-05-26

| Hora | Cambio |
|------|--------|
| 10:20 | feat(app): scaffold React/Vite Patin Score Pro con flujo completo demo |
| 10:22 | test: npm.cmd install OK, npm.cmd run build OK |
| 10:24 | test: dev server OK en http://127.0.0.1:5173/ |
| 10:32 | feat(app): prototipo completo implementado en subproyecto Vite `patin-score-pro/` |
| 10:33 | test: npm.cmd run lint OK, npm.cmd run build OK, dev server OK en http://localhost:5175/ |
| 10:38 | style(ui): tema visual verde/negro aplicado, lint y build OK |
| 10:48 | feat(flow): jueces solo puntuan, operador ausente/postergar, admin importa Excel y carga canciones por ID |
| 10:49 | test: npm.cmd run lint OK, npm.cmd run build OK, dev server HTTP 200 |
| 10:53 | feat(admin): import Excel/CSV para clubes y profesoras, lint/build OK |
| 10:56 | refactor(data): dorsal eliminado de vistas, import y rankings |
| 10:58 | feat(admin): import Excel/CSV para categorias, lint/build OK |
| 11:02 | fix(copy): textos visibles normalizados a castellano con acentos y ñ, lint/build OK |
| 11:08 | feat(display): pública muestra puntaje 10s y luego tanteador por categoría; ranking clubes separado |
| 11:10 | fix(scoreboard): tanteador solo muestra patinadoras con puntaje confirmado |
| 11:17 | feat(scoring): conceptos de puntaje configurables, consensuado sin selector de juez, total técnico preliminar |
| 11:22 | feat(scoring): publicación de puntaje pasa al juez; operador solo ve estado |
| 11:25 | fix(copy): panel import de patinadoras renombrado y columna Música simplificada |
| 11:30 | feat(admin): editores de clubes/profesoras/patinadoras con OK, carga manual y textos de logo/foto |
| 11:39 | feat(judges): Enter avanza campo, lista próximas, fin de categoría con siguiente/receso y receso automático en pública |
| 11:34 | feat(judges): Enter guarda sin publicar, publicación solo clic y muestra próxima patinadora |
| 11:52 | feat(eventos): configuración de evento, turnos por día, orden de salida, reportes imprimibles, hora online/fallback y clima en operador |
| 11:53 | test: npm.cmd run lint OK, npm.cmd run build OK |
| 12:02 | feat(datos): carga manual simplificada y orden de categorías por día/turno vía Excel o formulario |
| 12:03 | test: npm.cmd run lint OK, npm.cmd run build OK |
| 12:10 | feat(tv): tanteador público en filas compactas con parciales, total y paginación automática |
| 12:11 | test: npm.cmd run lint OK, npm.cmd run build OK |
| 12:18 | feat(datos): confirmación OK en cargas manuales y configuración de datos |
| 12:19 | test: npm.cmd run lint OK, npm.cmd run build OK |
| 12:23 | feat(reportes): orden fijo por día, turno, categoría y orden de ingreso |
| 12:24 | test: npm.cmd run lint OK, npm.cmd run build OK |
| 12:29 | feat(receso): demora configurable antes de pantalla de receso y rediseño TV con datos completos |
| 12:30 | test: npm.cmd run lint OK, npm.cmd run build OK |
| 12:35 | feat(branding): escudo del club organizador en cabecera y pantalla de receso con torneo/liga/club |
| 12:36 | test: npm.cmd run lint OK, npm.cmd run build OK |
| 12:48 | feat(sql): servidor Node con SQLite, API de estado, fallback localStorage y proxy Vite |
| 12:49 | test: npm.cmd run lint OK, npm.cmd run build OK, /api/health OK |
| 09:48 | feat(app): mejoras de operador, jueces, prueba de pista, buffet web, usuarios y reportes agrupados |
| 09:49 | test: npm.cmd run lint OK, npm.cmd run build OK |
