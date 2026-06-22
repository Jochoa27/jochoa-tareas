# ROADMAP · MEJORAS PLANIFICADAS
**Proyecto:** Centro de Comando — Juan Ochoa  
**Última actualización:** 2026-06-21  
**Convención de estados:** 💡 Idea · 📋 Definido · 🔨 En desarrollo · ✅ Completado · ❌ Descartado

---

## PRIORIDAD ALTA — Próximo sprint

### [F-001] Módulo de Reuniones 📋
**Descripción:** Crear el módulo de interfaz para la hoja REUNIONES (ya existe en el Excel).  
**Entregables:**
- Vista de reuniones pasadas y próximas
- Editor de compromisos con responsable y fecha
- Semáforo de compromisos vencidos
- Integración con módulo Seguimiento de Terceros (cruzar participantes)

**Dependencias:** Ninguna — la hoja ya existe.  
**Esfuerzo estimado:** Medio (1 sesión de trabajo)  
**Archivo afectado:** `tareas_app.py` — agregar bloque `elif mod == "Reuniones":`

---

### [F-002] Búsqueda y filtros en vistas Kanban 📋
**Descripción:** Agregar barra de búsqueda y filtro por proyecto en las vistas Kanban y Calendario del Centro de Comando.  
**Por qué:** Con más de 20 tareas activas el Kanban se vuelve difícil de leer.  
**Entregables:**
- Input de búsqueda sobre el Kanban (filtra por nombre de tarea)
- Selector de proyecto encima del Kanban
- Badge con conteo de tareas filtradas vs totales

**Esfuerzo estimado:** Bajo (< 1 sesión)

---

### [F-003] Persistencia del tema entre sesiones 📋
**Descripción:** Guardar la preferencia de tema en `st.query_params` para que persista al recargar.  
**Por qué:** Actualmente el tema se resetea a "Void" al recargar la página.  
**Esfuerzo estimado:** Muy bajo (< 30 min)

---

## PRIORIDAD MEDIA — Próximo mes

### [F-004] Vista de Gantt / Línea de tiempo ✅ COMPLETADO 2026-06-21
**Descripción:** Visualización de tareas en línea de tiempo horizontal (estilo Gantt simplificado).  
**Por qué:** El Calendario semanal solo muestra 7 días. No hay visibilidad de tareas a 30–90 días.  
**Alcance:**
- Eje X: días del mes actual + siguiente
- Eje Y: proyectos o tareas
- Barra coloreada por prioridad
- Clic para abrir detalle

**Esfuerzo estimado:** Alto — requiere ECharts custom chart o componente nuevo.

---

### [F-005] Notificaciones / Resumen diario 💡
**Descripción:** Sección de "Alerta del día" al abrir la app con las tareas más urgentes.  
**Entregables:**
- Pop-up o banner al cargar con: tareas vencidas hoy, tareas Críticas sin fecha, compromisos con terceros > 5 días sin respuesta
- Opción de descartar el banner por sesión

**Esfuerzo estimado:** Medio

---

### [F-006] Exportar a PDF / reporte semanal 💡
**Descripción:** Generar un PDF de resumen semanal con KPIs y lista de tareas activas.  
**Casos de uso:** Presentar en reuniones de coordinación, enviar a jefatura.  
**Dependencias:** Requiere librería (`weasyprint` o `reportlab`) — evaluar compatibilidad con Streamlit Cloud.  
**Esfuerzo estimado:** Alto

---

### [F-007] Historial de cambios por tarea 💡
**Descripción:** Mostrar timeline de cambios dentro del panel de detalle de una tarea.  
**Por qué:** COMENTARIOS almacena historial en texto plano — estructurarlo visualmente.  
**Entregables:**
- Timeline vertical en el panel de detalle
- Cada entrada muestra: fecha, campo cambiado, valor anterior → valor nuevo
- Requiere cambiar formato de COMENTARIOS de texto plano a JSON

**Esfuerzo estimado:** Medio-Alto (requiere migración de datos)

---

### [F-008] Campo RESPONSABLE en TAREAS 💡
**Descripción:** Agregar columna RESPONSABLE para tareas delegadas (además del campo TERCERO que es para externos).  
**Por qué:** En algunos proyectos Juan delega a colaboradores internos.  
**Entregables:**
- Agregar columna al Excel
- Mostrar en kanban y detalle
- Filtro por responsable en Bandeja Operacional

**Esfuerzo estimado:** Bajo

---

## PRIORIDAD BAJA — Backlog

### [F-009] Integración con iConstruye 💡
**Descripción:** Leer estados de facturas u órdenes de compra de iConstruye y mostrarlos como tareas o alertas.  
**Dependencias:** Requiere API o exportación de iConstruye.

---

### [F-010] App móvil / PWA 💡
**Descripción:** Optimización completa para móvil (actualmente hay responsive básico con media queries).  
**Alcance:**
- Sidebar colapsable en móvil
- Tarjetas Kanban adaptadas a pantalla vertical
- Botón flotante para agregar tarea rápida

---

### [F-011] Integración con calendario externo (Google Calendar / Outlook) 💡
**Descripción:** Sincronizar FECHA_COMPROMISO con eventos de calendario.  
**Dependencias:** OAuth2 con Google o Microsoft — complejidad alta.

---

### [F-012] Multi-usuario / equipos 💡
**Descripción:** Permitir que otros miembros del equipo tengan su propia vista filtrada.  
**Dependencias:** Requiere autenticación (actualmente sin login). Cambio arquitectural mayor.

---

## MEJORAS TÉCNICAS (deuda técnica)

| ID | Descripción | Complejidad | Prioridad |
|----|-------------|-------------|-----------|
| TEC-001 | Usar SHA de commit como clave de cache (en lugar de mtime) | Baja | Media |
| TEC-002 | Separar CSS en archivo externo o función centralizada | Baja | Baja |
| TEC-003 | Tests automatizados para funciones de lógica de negocio | Media | Baja |
| TEC-004 | Migrar COMENTARIOS de texto plano a JSON estructurado | Media | Media |
| TEC-005 | Confirmación antes de completar/cancelar tareas desde Kanban | Baja | Media |

---

## HISTORIAL DE ENTREGAS

| Fecha | ID | Descripción | Estado |
|-------|----|-------------|--------|
| 2026-06-21 | — | Creación de archivos CONTEXTO, MEMORIA, ROADMAP | ✅ |
| 2026-06-21 | DEC-010 | Eliminar hoja REUNIONES del Excel y código | ✅ |
| 2026-06-21 | DEC-011 | ESFUERZO_HRS badge en tarjetas Kanban, Calendario y Agenda | ✅ |
| 2026-06-21 | DEC-012 | Cache SHA-based (TEC-001) | ✅ |
| 2026-06-21 | F-004 | Módulo Gantt con ECharts custom series | ✅ |

---

## CÓMO AGREGAR UNA MEJORA

1. Asignar ID correlativo (`F-XXX`)
2. Definir descripción, entregables concretos, dependencias y esfuerzo
3. Actualizar estado cuando avance
4. Al completar: registrar en columna **Historial de entregas** con fecha
5. Registrar decisiones de diseño en `MEMORIA.md`
