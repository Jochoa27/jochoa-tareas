# CONTEXTO · CENTRO DE COMANDO PERSONAL
**Proyecto:** Centro de Comando — Juan Ochoa  
**Organización:** Constructora Londres · Planificación y Control  
**Versión actual:** 2.0  
**Fecha de referencia:** 2026-06-21  

---

## 1. PROPÓSITO

Sistema personal de gestión de tareas y seguimiento operacional para el área de Planificación y Control de Constructora Londres. Permite crear, editar, priorizar y visualizar tareas, compromisos con terceros y reuniones, con persistencia en la nube vía GitHub.

---

## 2. STACK TECNOLÓGICO

| Capa | Tecnología | Versión mínima |
|------|-----------|----------------|
| Frontend / App | Streamlit | ≥ 1.36.0 |
| Procesamiento de datos | Pandas | ≥ 2.1.0 |
| Lectura/escritura Excel | openpyxl | ≥ 3.1.0 |
| Gráficos interactivos | Apache ECharts 5.4.3 | via CDN |
| Drag & drop | SortableJS 1.15.2 | via CDN |
| Persistencia en nube | GitHub REST API v3 | — |
| Despliegue | Streamlit Cloud | — |
| Lenguaje | Python 3.x | — |

---

## 3. ARCHIVOS DEL PROYECTO

```
TAREAS_APP/
├── tareas_app.py          # Aplicación principal (2,254 líneas)
├── TAREAS_JOCHOA.xlsx     # Fuente de datos local + repo GitHub
├── crear_excel.py         # Script v1 de inicialización del Excel
├── crear_excel_v2.py      # Script v2 de inicialización del Excel
├── requirements.txt       # Dependencias pip
├── CONTEXTO.md            # Este archivo
├── MEMORIA.md             # Log de decisiones y evolución
└── ROADMAP.md             # Mejoras pendientes y planificadas
```

---

## 4. MODELO DE DATOS

### 4.1 Hoja: TAREAS (fuente principal — 44 registros a jun-2026)

| Columna | Tipo | Valores posibles | Obligatorio |
|---------|------|-----------------|-------------|
| ID | Entero | Auto-incremental | Sí |
| TAREA | Texto | Nombre de la tarea | Sí |
| DESCRIPCION | Texto libre | — | No |
| TIPO | Categoría | Tarea / Seguimiento / Compromiso / Reunión | Sí |
| PROYECTO | Texto | CL038, SALUD, ICONSTRUYE, CORPORATIVO, PERSONAL, DIRECTORIO, AGENTE_IA, CL027, GENERAL | Sí |
| AREA | Categoría | Trabajo / Personal | Sí |
| CATEGORIA | Categoría | Planificación / Contratos / Compras / Reportes / IA y Automatización / Gestión Corporativa / Reuniones / Salud / Personal | Sí |
| PRIORIDAD | Categoría | Crítica / Alta / Media / Baja | Sí |
| ESTADO | Categoría | Pendiente / En Proceso / Esperando Terceros / Completada / Cancelada | Sí |
| IMPACTO | Entero 1–5 | Escala Likert | No |
| URGENCIA | Entero 1–5 | Calculado automáticamente si = 0 desde FECHA_COMPROMISO | No |
| ESFUERZO_HRS | Decimal | Horas estimadas | No |
| TERCERO | Texto | Nombre del tercero responsable | No |
| FECHA_CREACION | Fecha | Asignada al crear | Auto |
| FECHA_COMPROMISO | Fecha | Deadline de la tarea | No |
| FECHA_CIERRE | Fecha | Asignada al completar | Auto |
| NOTAS | Texto libre | — | No |
| COMENTARIOS | Texto libre | Historial append con timestamp | No |

**Columna calculada (en runtime):**  
- `PRIO_ORD`: orden numérico de prioridad (Crítica=0, Alta=1, Media=2, Baja=3) — se elimina antes de guardar

**Lógica de URGENCIA automática** (cuando el valor es 0 o NaN):

| Días hasta vencimiento | Urgencia asignada |
|------------------------|-------------------|
| Vencida (< 0) | 5 |
| 0–3 días | 5 |
| 4–7 días | 4 |
| 8–14 días | 3 |
| 15–30 días | 2 |
| > 30 días | 1 |

---

### 4.2 Hoja: TERCEROS (6 registros a jun-2026)

| Columna | Descripción |
|---------|-------------|
| ID | Auto-incremental |
| NOMBRE | Nombre del tercero |
| ORGANIZACION | Empresa u organización |
| ROL | Rol en la relación |
| TEMA_PENDIENTE | Descripción del pendiente |
| FECHA_INICIO_SEG | Inicio del seguimiento |
| FECHA_ULTIMO_SEG | Último contacto |
| ESTADO | Pendiente / Resuelto |
| PRIORIDAD | Alta / Media / Baja |
| NOTAS | Notas libres |

**Columna calculada (en runtime):**  
- `DIAS_SIN_RESP`: días desde FECHA_ULTIMO_SEG hasta hoy — se elimina antes de guardar

---

### 4.3 Hoja: REUNIONES (4 registros a jun-2026)

| Columna | Descripción |
|---------|-------------|
| ID | Auto-incremental |
| TITULO | Nombre de la reunión |
| FECHA | Fecha de la reunión |
| TIPO | Tipo de reunión |
| PARTICIPANTES | Lista de participantes |
| ESTADO | Pendiente / Completada |
| COMPROMISO | Acuerdo o tarea generada |
| RESPONSABLE_COMP | Responsable del compromiso |
| FECHA_COMP | Fecha límite del compromiso |
| ESTADO_COMP | Estado del compromiso |
| NOTAS | Notas libres |

> **Nota:** La hoja REUNIONES existe en el Excel pero no tiene módulo de interfaz en la app v2.0. Es una mejora pendiente.

---

## 5. MÓDULOS DE LA APLICACIÓN

| # | Módulo | Ícono | Función principal |
|---|--------|-------|-------------------|
| 1 | Centro de Comando | 🎯 | Dashboard con KPIs, Calendario semanal (7 días), vistas Kanban |
| 2 | Diagnóstico de Riesgo | ⚠️ | Matriz Urgencia × Impacto con cuadrantes, tabla de riesgo ordenada |
| 3 | Bandeja Operacional | 📋 | Editor completo con filtros, búsqueda, ordenamiento y guardado |
| 4 | Seguimiento de Terceros | 👥 | Estado de terceros, días sin respuesta, tareas vinculadas |
| 5 | Productividad | 📈 | Completadas por semana, por proyecto, por categoría, tasa semanal |
| 6 | Consumo de Tiempo | ⏱ | Horas estimadas por categoría, proyecto y balance trabajo/personal |

### Centro de Comando — vistas disponibles
| Vista | Descripción |
|-------|-------------|
| 📅 Cal. | Calendario semanal de 7 días con drag & drop entre días |
| 🔷 Estado | Kanban por estado (Pendiente / En Proceso / Esperando Terceros / Completada) |
| 🎯 Prio. | Kanban por prioridad (Crítica / Alta / Media / Baja) |
| 📁 Área | Kanban por área (Trabajo / Personal) |
| 🗂 Cat. | Kanban por categoría (9 columnas) |
| 🚨 Vencidas | Kanban de tareas vencidas agrupadas por prioridad |

---

## 6. PALETA DE COLORES

| Nombre | Hex | Uso |
|--------|-----|-----|
| C_BG | #060B15 | Fondo principal (tema Void) |
| C_CIAN | #38BDF8 | Acento principal, tareas "En Proceso" |
| C_CRITICO | #FF4757 | Prioridad Crítica, vencidas, alertas |
| C_ALERTA | #FFB300 | Prioridad Alta, alertas moderadas |
| C_OK | #23D160 | Completadas, estado controlado |
| C_GRIS | #64748B | Prioridad Baja, sin fecha |
| C_MORADO | #A855F7 | Tipo "Compromiso", Consumo de Tiempo |
| C_VERDE2 | #10B981 | Tipo "Reunión", Compras |
| C_INDIGO | #818CF8 | Categoría Planificación |

### Temas disponibles
| Tema | Fondo | Sidebar | Uso |
|------|-------|---------|-----|
| Void (default) | #050507 | #030305 | Ultra dark, premium |
| Zinc | #18181B | #101012 | Dark gray, profesional |
| Pearl | #FAFAFA | #18181B | Modo claro con sidebar oscuro |

---

## 7. DESPLIEGUE Y PERSISTENCIA

### Flujo de datos
```
[Excel local]  ──(git push)──►  [GitHub repo: Jochoa27/jochoa-tareas]
                                        │
                                        ▼
                              [Streamlit Cloud]  ◄──(API)──  [Usuario en browser]
                                        │
                                  (GITHUB_TOKEN)
                                        │
                                        ▼
                              [Edición en app]  ──(PUT API)──►  [GitHub repo]
                                                                      │
                                                              (auto-sync Streamlit)
```

### Configuración de secretos (Streamlit Cloud)
```
GITHUB_TOKEN = "ghp_..."   # Token con permisos repo write
```

> Sin `GITHUB_TOKEN`, la app funciona en modo lectura: visualiza datos pero no puede guardar cambios desde la plataforma.

### URL del dashboard
`https://jochoa-tareas-xun7r85qv6xntjuq9qjgw2.streamlit.app/`

### Repositorio GitHub
`https://github.com/Jochoa27/jochoa-tareas`  
Rama principal: `main`  
Archivo de datos: `TAREAS_JOCHOA.xlsx`

---

## 8. LÓGICA DE NEGOCIO CLAVE

### Estado general del sistema
| Condición | Estado | Color |
|-----------|--------|-------|
| Hay tareas Críticas vencidas O > 5 vencidas | CRÍTICO | Rojo |
| Hay vencidas O ≥ 2 tareas Críticas activas | BAJO PRESIÓN | Amarillo |
| Sin vencidas ni críticas excesivas | CONTROLADO | Verde |

### Tasa de cumplimiento semanal
```
Completadas esta semana / (Completadas esta semana + Vencidas de los últimos 7 días)
```

### Semáforo de fechas
| Días restantes | Color | Texto |
|---------------|-------|-------|
| Negativo | Rojo | "Vencida Xd" |
| 0 | Rojo | "Vence HOY" |
| 1–3 | Amarillo | "Vence en Xd" |
| 4–7 | Cian | "Vence en Xd" |
| > 7 | Gris | "Vence en Xd" |

### Puntuación de riesgo
```
RIESGO = URGENCIA × IMPACTO   (rango: 1–25)
≥ 20 → CRÍTICO | ≥ 12 → ALERTA | < 12 → NORMAL
```
