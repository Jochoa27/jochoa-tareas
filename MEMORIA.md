# MEMORIA · DECISIONES Y EVOLUCIÓN DEL PROYECTO
**Proyecto:** Centro de Comando — Juan Ochoa  
**Formato:** Registro cronológico de decisiones de diseño, arquitectura y UX  
**Actualizar con cada cambio significativo.**

---

## CÓMO USAR ESTE ARCHIVO

Cada entrada debe responder tres preguntas:
1. **¿Qué se decidió o cambió?**
2. **¿Por qué?** (contexto, problema que resolvía)
3. **¿Qué alternativas se descartaron y por qué?**

---

## REGISTRO DE DECISIONES

---

### [DEC-001] Streamlit como framework de la aplicación
**Fecha:** 2026 (construcción inicial)  
**Estado:** ✅ Activo

**Decisión:** Usar Streamlit como único framework (frontend + backend en Python).

**Por qué:**  
- Velocidad de desarrollo para uso personal/interno
- Sin necesidad de servidor dedicado (Streamlit Cloud gratuito)
- Juan Ochoa maneja Python; no requiere aprender JS/React

**Alternativas descartadas:**  
- Dash/Plotly: más control pero mayor complejidad de deploy
- Aplicación web custom (React + FastAPI): excesivo para el alcance personal
- Notion / ClickUp: no integrables con flujo de trabajo propio de P&C

---

### [DEC-002] Excel como fuente de datos (en lugar de base de datos)
**Fecha:** 2026 (construcción inicial)  
**Estado:** ✅ Activo

**Decisión:** Usar `TAREAS_JOCHOA.xlsx` con 3 hojas (TAREAS, TERCEROS, REUNIONES) como única fuente de datos.

**Por qué:**  
- El Excel también sirve como respaldo editable sin la app
- Facilidad de edición directa cuando Streamlit Cloud no está disponible
- Compatible con el flujo de trabajo existente en Constructora Londres
- Sin necesidad de base de datos relacional para ~100 registros

**Alternativas descartadas:**  
- SQLite: requiere acceso al servidor, no funciona bien en Streamlit Cloud sin volumen persistente
- Google Sheets: depende de API de Google con credenciales adicionales; latencia mayor
- Supabase/PostgreSQL: sobre-ingeniería para el volumen actual

**Riesgo conocido:** Concurrencia — si dos sesiones guardan simultáneamente puede haber conflicto. Aceptado dado uso personal (usuario único).

---

### [DEC-003] Persistencia vía GitHub API (no escritura directa de archivo)
**Fecha:** 2026 (construcción inicial)  
**Estado:** ✅ Activo

**Decisión:** Los cambios desde la app se guardan haciendo PUT al archivo en el repo GitHub vía REST API, usando `GITHUB_TOKEN` como secreto de Streamlit Cloud.

**Por qué:**  
- Streamlit Cloud no tiene sistema de archivos persistente entre reinicios
- GitHub actúa como base de datos con historial de versiones (cada guardado = commit)
- Permite sincronización bidireccional: edición local → git push | edición en app → PUT API

**Consecuencia:** Sin `GITHUB_TOKEN` configurado, la app funciona en modo lectura. Se muestra aviso en sidebar.

---

### [DEC-004] Apache ECharts para gráficos (en lugar de Plotly/Altair)
**Fecha:** 2026 (construcción inicial)  
**Estado:** ✅ Activo

**Decisión:** Usar ECharts 5.4.3 via CDN dentro de `st.components.v1.html()` para todos los gráficos.

**Por qué:**  
- Estética más premium y personalizable que Plotly nativo de Streamlit
- Soporte SVG renderer (mejor en pantallas HiDPI)
- Tooltips y animaciones más fluidos
- Funciones JS inline para lógica dinámica (símbolo por tamaño de esfuerzo)

**Trade-off aceptado:** No hay interactividad bidireccional (clic en gráfico → Python). Los gráficos son de solo visualización.

---

### [DEC-005] SortableJS para drag & drop en Kanban y Calendario
**Fecha:** 2026 (construcción inicial)  
**Estado:** ✅ Activo

**Decisión:** Implementar drag & drop usando SortableJS 1.15.2 via CDN en componentes HTML embebidos. La comunicación JS→Python se hace mediante inputs ocultos (`aria-label="TASK_ACTION"` y `DRAG_ORDER`) que se modifican con native setter + eventos sintéticos.

**Por qué:**  
- Streamlit no tiene drag & drop nativo
- SortableJS es estable, ligero y sin dependencias
- El truco del input oculto permite comunicación desde iframe a Streamlit

**Limitación conocida:** Requiere doble clic para abrir detalle de tarea (para distinguir del drag). El primer clic inicia temporizador de 380ms.

---

### [DEC-006] Cache con TTL=30s invalidado por mtime del archivo
**Fecha:** 2026 (construcción inicial)  
**Estado:** ✅ Activo

**Decisión:** `@st.cache_data(ttl=30)` con argumento `mtime` del archivo como clave de invalidación.

**Por qué:**  
- Si el archivo no cambió, no re-lee el disco (performance)
- Si el archivo cambió (por git push o guardado desde app), el mtime cambia → cache invalida automáticamente

**Riesgo conocido:** En Streamlit Cloud, el archivo se descarga de GitHub al iniciar. El mtime local puede no reflejar el tiempo real del cambio remoto. Mejora pendiente: usar el SHA del commit como clave de cache.

---

### [DEC-007] Hoja REUNIONES sin módulo de interfaz (deuda técnica)
**Fecha:** 2026  
**Estado:** ⏳ Pendiente — ver ROADMAP.md

**Situación:** La hoja REUNIONES existe en el Excel (4 registros, estructura completa con compromisos y responsables) pero no hay módulo de interfaz en la app para visualizarla ni editarla.

**Por qué no se implementó aún:** La hoja fue creada para uso futuro durante el diseño del modelo de datos.

---

### [DEC-008] Sistema de 3 temas de color (Void / Zinc / Pearl)
**Fecha:** 2026 (v2.0)  
**Estado:** ✅ Activo

**Decisión:** Implementar selector de tema en sidebar con 3 opciones: Void (ultra dark), Zinc (dark gray), Pearl (light).

**Por qué:**  
- El modo oscuro es estándar para dashboards de operaciones
- Pearl permite usar en salas con mucha luz sin forzar la vista
- Los 3 temas comparten la misma estructura de variables (`_TC`) para no duplicar CSS

**Implementación:** El tema se guarda en `st.session_state["theme"]` y persiste durante la sesión (no entre sesiones — mejora posible con `st.query_params` o cookie).

---

### [DEC-009] Módulo "Diagnóstico de Riesgo" con cuadrantes Urgencia × Impacto
**Fecha:** 2026 (v2.0)  
**Estado:** ✅ Activo

**Decisión:** Visualizar tareas como scatter plot con ejes Urgencia (X) e Impacto (Y), tamaño proporcional al Esfuerzo estimado, dividido en 4 cuadrantes.

| Cuadrante | Zona | Acción recomendada |
|-----------|------|-------------------|
| Alta urgencia + Alto impacto | Rojo | ACTUAR AHORA |
| Baja urgencia + Alto impacto | Amarillo | PLANIFICAR |
| Alta urgencia + Bajo impacto | Azul | DELEGAR |
| Baja urgencia + Bajo impacto | Gris | MONITOREAR |

**Modelo:** Matriz de Eisenhower adaptada con escala 1–5.

---

## EVOLUCIÓN DE VERSIONES

| Versión | Descripción | Cambios clave |
|---------|-------------|---------------|
| v1.0 | Versión inicial | Tabla básica de tareas, sin persistencia cloud |
| v2.0 | Rediseño completo | Kanban con drag & drop, Calendario, ECharts, temas, GitHub API, módulos separados |

---

---

### [DEC-010] Eliminación de hoja REUNIONES del Excel y la app
**Fecha:** 2026-06-21  
**Estado:** ✅ Aplicado

**Decisión:** Se eliminó la hoja REUNIONES del archivo `TAREAS_JOCHOA.xlsx` y todas sus referencias en `tareas_app.py` (carga en `cargar()`, variable `df_reu`, preservación en `guardar_github()`).

**Por qué:** La hoja no tenía módulo de interfaz y generaba carga innecesaria. Las reuniones se gestionan como tareas de tipo "Reunión" en el Kanban, que sí tiene UI completa.

---

### [DEC-011] ESFUERZO_HRS visible en todas las tarjetas
**Fecha:** 2026-06-21  
**Estado:** ✅ Aplicado

**Decisión:** Se agregó badge de horas estimadas (`kc-hrs` en cian) en las tarjetas del Calendario, Kanban y Agenda del Día. Solo visible si `ESFUERZO_HRS > 0`. Formato inteligente: "2h" para enteros, "2.5h" para decimales.

---

### [DEC-012] Cache basado en SHA de GitHub (en lugar de mtime local)
**Fecha:** 2026-06-21  
**Estado:** ✅ Aplicado

**Decisión:** Se reemplazó `mtime` local como clave de cache por el SHA del archivo en GitHub (obtenido via API, cacheado 60 s con `@st.cache_data(ttl=60)`). Fallback a `mtime` si no hay token.

**Por qué:** En Streamlit Cloud el `mtime` del archivo local no cambia cuando otro usuario actualiza el Excel vía la API de GitHub en otra sesión. El SHA sí cambia con cada push/PUT. Resuelve DT-002.

**Efecto adicional:** Al guardar (`guardar_github`), ahora se llama `_get_file_cache_key.clear()` además de `cargar.clear()` para forzar re-fetch del SHA en la próxima carga.

---

### [DEC-013] Módulo Gantt (nivel MS Project) con ECharts custom series
**Fecha:** 2026-06-21  
**Estado:** ✅ Aplicado

**Decisión:** Se agregó el módulo "📊 Gantt" al sidebar. Usa ECharts `custom series` con `renderItem` para barras Gantt con bordes redondeados, texto inline en barras largas, y marcador "HOY" en rojo.

**Características:**
- Filas agrupadas por proyecto con cabeceras coloreadas
- Barras coloreadas por Prioridad / Estado / Proyecto (selector)
- Marcador "HOY" vertical en rojo
- Slider de navegación temporal (zoom horizontal)
- Tooltip rico: nombre, proyecto, prioridad, estado, horas, tercero, fechas, duración
- DataZoom doble (slider + inside scroll/pinch)
- Para tareas sin FECHA_COMPROMISO: duración = max(3d, ESFUERZO_HRS)
- Para tareas Completadas: barra hasta FECHA_CIERRE real (si existe)
- KPIs de contexto bajo el Gantt: tareas en vista, completadas, vencidas, HH estimadas, % completado
- Leyenda de proyectos por color cuando modo = "Proyecto"

---

## DEUDAS TÉCNICAS CONOCIDAS

| ID | Descripción | Impacto | Prioridad |
|----|-------------|---------|-----------|
| DT-001 | Hoja REUNIONES sin módulo UI | Funcionalidad perdida | Media |
| DT-002 | Cache usa mtime local (puede desfasar con GitHub) | Datos obsoletos en cloud | Baja |
| DT-003 | Tema no persiste entre sesiones | UX menor | Baja |
| DT-004 | Sin confirmación antes de marcar completada | Riesgo de accidente | Baja |
| DT-005 | Sin filtro de búsqueda en vistas Kanban | UX con muchas tareas | Media |
| DT-006 | Kanban de Categoría genera 9 columnas — difícil en móvil | UX móvil | Baja |
