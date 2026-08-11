# Estado de Migración — INFOREST

> Última actualización: 2026-08-11
>
> Estado general: **Fase 2 — Análisis Técnico, Arquitectura y Setup (EN CURSO)**

---

## Resumen Ejecutivo

| Indicador | Valor |
|---|---|
| Fase actual | 2 — Análisis Técnico, Arquitectura y Setup |
| Código .NET 8 existente | IN_PROGRESS — Solución base creada en `modern-net8/` (Domain, Application, Infrastructure, Desktop, Tests) |
| Módulos migrados | 0 / 7 |
| Documentación Legacy | IN_PROGRESS — inventarios y matrices de Fase 2 en construcción/actualización |
| Arquitectura Target definida | IN_PROGRESS — ADR-001 a ADR-008 aceptados |

---

## Estado por Módulo

| Módulo | Ejecutable Legacy | Fase | Estado | Notas |
|---|---|---|---|---|
| Punto de Venta | `InfoRest.exe` | Análisis | NOT_STARTED | Baseline funcional identificado; depende de Infrastructure + Configuración + Turno |
| Caja Rápida | `CajaRapida.exe` | Análisis | NOT_STARTED | Comparte gran parte del stack POS/caja |
| Adición | `Adicion.exe` | Análisis | NOT_STARTED | Flujo auxiliar sobre pedidos en curso |
| Administración | `Administracion.exe` | Análisis | NOT_STARTED | Maestros, parámetros y catálogos críticos |
| Consultas | `Consulta.exe` | Análisis | NOT_STARTED | Alta concentración de reportes `spRep_*` + Crystal |
| Despachador | `Despachador.exe` | Análisis | NOT_STARTED | Delivery, central pedidos e integraciones externas |
| Motorizados | `Motorizado.exe` | Análisis | NOT_STARTED | Depende de delivery/despacho y tarifas |

---

## Estado por Capa

| Capa | Legacy | .NET 8 | Estado |
|---|---|---|---|
| Presentación (UI) | 400 formularios VB6 detectados | 0 | NOT_STARTED |
| Lógica de negocio | 32 módulos + 10 clases | 0 | NOT_STARTED |
| Acceso a datos | ADO + `clsComando` + 150 SP | 0 | NOT_STARTED |
| Base de datos | 126T + 116V + 150SP | 0 | IN_PROGRESS |
| Reportes | 206 Crystal Reports | 0 | NOT_STARTED |
| Integraciones | COM/DLL/OCX + hardware POS | 0 | NOT_STARTED |
| Seguridad/Auth | `INFSEGURIDAD` + cifrado débil Legacy | 0 | IN_PROGRESS |
| Configuración | INI + `TPARAMETRO` + `TCAJA` | 0 | IN_PROGRESS |

---

## Cobertura de Migración

> No disponible todavía para código ejecutable — no existe implementación .NET 8 funcional.

| Componente | Inventario Legacy | Migrado | Cobertura |
|---|---|---|---|
| Formularios | 400 | 0 | NOT_AVAILABLE |
| Módulos BAS | 32 | 0 | NOT_AVAILABLE |
| Clases | 10 | 0 | NOT_AVAILABLE |
| Stored Procedures | 150 | 0 | NOT_AVAILABLE |
| Tablas | 126 | 0 | NOT_AVAILABLE |
| Vistas | 116 | 0 | NOT_AVAILABLE |
| Reportes | 206 | 0 | NOT_AVAILABLE |
| Integraciones externas | 10+ | 0 | NOT_AVAILABLE |

---

## Milestones

| Milestone | Descripción | Estado |
|---|---|---|
| M1 | Documentación Legacy completa | COMPLETED |
| M2 | Arquitectura Target definida | IN_PROGRESS |
| M3 | Proyecto .NET 8 base creado | IN_PROGRESS |
| M4 | Primer módulo migrado (Maestros) | NOT_STARTED |
| M5 | Punto de Venta migrado | NOT_STARTED |
| M6 | Caja y Pagos migrados | NOT_STARTED |
| M7 | Todos los módulos migrados | NOT_STARTED |
| M8 | Validación completa | NOT_STARTED |
| M9 | Go-live en producción | NOT_STARTED |

---

## Blockers Actuales

| Blocker | Descripción | Impacto |
|---|---|---|
| ADR-001 resuelto | Cliente POS objetivo definido: WinForms .NET 8 | Desbloqueado |
| ADR-002 resuelto | Motor de base y estrategia de compatibilidad definidos: SQL Server | Desbloqueado |
| ADR-003 resuelto | Patrón arquitectónico definido: Clean Architecture + CQRS | Desbloqueado |
| ADR-004 resuelto | Estrategia de migración definida: Strangler Fig | Desbloqueado |
| ADR-005 resuelto | Configuración y secretos definidos | Desbloqueado |
| ADR-006 resuelto | Línea base de autenticación/autorización definida | Desbloqueado |
| ADR-007 resuelto | Estrategia de reportes definida: FastReport .NET | Desbloqueado |
| ADR-008 resuelto | Estrategia multi-país/multi-local definida | Desbloqueado |

---

## Próximos Pasos

1. Crear la solución base .NET 8 en `modern-net8/` conforme a ADR-001..ADR-003.
2. Modelar Infrastructure + Database adapters para SQL Server/SPs.
3. Migrar Seguridad y Configuración (`TPARAMETRO` / `TCAJA`).
4. Iniciar Maestros (Productos / Grupos / Clientes).
5. Preparar la primera vertical operativa: Turno → Pedido simple.

---

## Referencias

- [Estrategia de migración](migration-strategy.md)
- [Inventario Legacy](legacy-inventory.md)
- [Inventario SQL](database/sql-inventory.md)
- [Matriz de lógica de negocio](traceability/business-logic-matrix.md)
- [Decisiones arquitectónicas](../architecture/architecture-decisions.md)
