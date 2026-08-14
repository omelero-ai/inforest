# Documentación — INFOREST

> Índice central de toda la documentación técnica y arquitectónica del proyecto de migración INFOREST VB6 → .NET 8.

---

## Navegación

### Arquitectura

| Documento | Descripción |
|---|---|
| [architecture/README.md](architecture/README.md) | Índice de documentación arquitectónica |
| [architecture/system-context.md](architecture/system-context.md) | Contexto del sistema (diagrama C4 nivel 1) |
| [architecture/legacy-architecture.md](architecture/legacy-architecture.md) | Arquitectura del sistema Legacy VB6 |
| [architecture/target-architecture.md](architecture/target-architecture.md) | Arquitectura objetivo .NET 8 |
| [architecture/architecture-overview.md](architecture/architecture-overview.md) | Vista general comparativa |
| [architecture/architecture-decisions.md](architecture/architecture-decisions.md) | ADRs — Decisiones arquitectónicas |

### Migración

| Documento | Descripción |
|---|---|
| [migration/README.md](migration/README.md) | Índice de documentación de migración |
| [migration/migration-strategy.md](migration/migration-strategy.md) | Estrategia de migración |
| [migration/migration-status.md](migration/migration-status.md) | Estado actual de migración |
| [migration/traceability-matrix.md](migration/traceability-matrix.md) | Matriz Legacy → .NET 8 |
| [migration/known-gaps.md](migration/known-gaps.md) | Brechas y pendientes |
| [migration/migration-rules.md](migration/migration-rules.md) | Reglas del proceso de migración |
| [migration/business-rules.md](migration/business-rules.md) | Reglas de negocio extraídas del Legacy |

### Módulos

| Documento | Descripción |
|---|---|
| [modules/README.md](modules/README.md) | Índice de módulos |
| [modules/restaurante/README.md](modules/restaurante/README.md) | Módulo Restaurante (principal) |

### Base de Datos

| Documento | Descripción |
|---|---|
| [database/README.md](database/README.md) | Índice de base de datos |
| [database/legacy-database.md](database/legacy-database.md) | Base de datos Legacy SQL Server |
| [database/target-database.md](database/target-database.md) | Base de datos Target (objetivo) |
| [database/mapping.md](database/mapping.md) | Mapeo Legacy → Target |
| [database/stored-procedures.md](database/stored-procedures.md) | Inventario de Stored Procedures |

### Estándares

| Documento | Descripción |
|---|---|
| [standards/README.md](standards/README.md) | Índice de estándares |
| [standards/coding-standards.md](standards/coding-standards.md) | Estándares de codificación .NET 8 |
| [standards/naming-conventions.md](standards/naming-conventions.md) | Convenciones de nombres |
| [standards/documentation-standards.md](standards/documentation-standards.md) | Estándares de documentación |

### Referencia

| Documento | Descripción |
|---|---|
| [glossary.md](glossary.md) | Glosario de términos |
| [../legacy-restaurant/README.md](../legacy-restaurant/README.md) | Análisis técnico del Legacy (fuente primaria) |
| [../.github/copilot-instructions.md](../.github/copilot-instructions.md) | Instrucciones para Copilot/IA |

---

## Estado Global de Migración

| Área | Estado |
|---|---|
| Análisis Legacy | PARTIAL — legacy-restaurant/README.md detallado, documentación docs/ en construcción |
| Arquitectura .NET 8 definida | NOT_STARTED |
| Código .NET 8 | NOT_STARTED |
| Base de datos Target | NOT_STARTED |
| Pruebas automatizadas | NOT_STARTED |

---

## Convenciones de Estado

- `CONFIRMED` — Verificado con evidencia del repositorio
- `PARTIAL` — Información incompleta
- `UNKNOWN` — No determinable
- `NOT_STARTED` — No iniciado
- `IN_PROGRESS` — En curso
- `COMPLETED` — Finalizado y validado
- `BLOCKED` — Bloqueado por dependencia

---

*Toda la documentación se basa exclusivamente en evidencia del repositorio.*
