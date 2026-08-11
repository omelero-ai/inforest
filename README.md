# INFOREST — Sistema de Gestión de Restaurante

> **Repositorio de migración:** Visual Basic 6 → .NET 8

---

## ¿Qué es INFOREST?

**INFOREST** es un sistema empresarial de gestión de restaurante desarrollado originalmente en **Visual Basic 6 (VB6)** con base de datos **SQL Server**. Cubre la operación completa de establecimientos de restauración: punto de venta en salón, caja, delivery, despacho, cocina (KDS), inventario, reportería y facturación electrónica multi-país.

| Campo | Valor |
|---|---|
| **Nombre** | INFOREST |
| **Tipo** | Sistema de Punto de Venta y Gestión de Restaurante (ERP Gastronómico) |
| **Usuarios principales** | Mozos, cajeros, administradores, jefes de cocina, motorizados, gerentes |
| **Contexto empresarial** | Establecimientos de restauración en múltiples países de Latinoamérica y España |
| **Multilocal** | Sí — soporta administración centralizada para múltiples locales |
| **Multi-país** | Sí — Perú, Chile, Bolivia, Ecuador, Argentina, España |

---

## Estado Tecnológico

```
Legacy (actual)          Target (objetivo)
───────────────          ─────────────────
Visual Basic 6     →     .NET 8 (C#)
SQL Server         →     Por definir
Crystal Reports    →     Por definir
COM/ActiveX        →     API REST / Web
INI files          →     Configuration (appsettings.json)
Windows Desktop    →     Por definir (Web / Desktop)
```

### Tecnologías detectadas

| Tecnología | Capa | Estado |
|---|---|---|
| Visual Basic 6.0 | UI + Lógica + Datos | CONFIRMED |
| SQL Server (INFOREST, INFSEGURIDAD, ALMACEN, CD, FE) | Persistencia | CONFIRMED |
| Crystal Reports 6/9 | Reportería | CONFIRMED — 206 reportes |
| ADO (ADODB) | Acceso a datos | CONFIRMED |
| COM/ActiveX (OCX) | UI Components | CONFIRMED |
| KDS (Kitchen Display System) | Cocina | CONFIRMED |
| BlueVision/TVS | Display cliente | CONFIRMED |
| Epson Impresora Fiscal | Argentina FE | CONFIRMED |
| PinPad DLL3500 | Pagos tarjeta | CONFIRMED |
| CashDro | Cajón automático | CONFIRMED |
| Biometría (SecuGen FpLibX) | Seguridad | CONFIRMED |
| .NET 8 | Target | NOT_STARTED |

---

## Estado de Migración

```
Legacy VB6
    ↓
[AQUÍ] Análisis y Documentación
    ↓
Diseño Arquitectura .NET 8
    ↓
Migración por Módulos
    ↓
Pruebas y Validación
    ↓
.NET 8 en Producción
```

---

## Módulos del Sistema

| Módulo | Ejecutable Legacy | .NET 8 | Estado | Documentación |
|---|---|---|---|---|
| Punto de Venta (Salón) | `InfoRest.exe` | — | NOT_STARTED | [docs/modules/restaurante/](docs/modules/restaurante/README.md) |
| Caja Rápida | `CajaRapida.exe` | — | NOT_STARTED | [docs/modules/restaurante/](docs/modules/restaurante/README.md) |
| Adición | `Adicion.exe` | — | NOT_STARTED | [docs/modules/restaurante/](docs/modules/restaurante/README.md) |
| Administración | `Administracion.exe` | — | NOT_STARTED | [docs/modules/restaurante/](docs/modules/restaurante/README.md) |
| Consultas / Reportes | `Consulta.exe` | — | NOT_STARTED | [docs/modules/restaurante/](docs/modules/restaurante/README.md) |
| Despachador / Delivery | `Despachador.exe` | — | NOT_STARTED | [docs/modules/restaurante/](docs/modules/restaurante/README.md) |
| Motorizados | `Motorizado.exe` | — | NOT_STARTED | [docs/modules/restaurante/](docs/modules/restaurante/README.md) |

---

## Estructura del Repositorio

```
/
├── README.md                          ← Este archivo
├── legacy-restaurant/                 ← Código fuente Legacy VB6 + BD SQL Server
│   ├── README.md                      ← Análisis técnico detallado del Legacy (1820 líneas)
│   ├── restaurant-vb6/                ← 7 proyectos VB6, 401 forms, 32 módulos, 10 clases, 206 reportes
│   └── database-sql-server/           ← Scripts SQL: 126 tablas, 105 vistas, 105+ SPs
├── modern-net8/                       ← Implementación .NET 8 (NO INICIADA)
├── docs/                              ← Documentación técnica y arquitectónica
│   ├── architecture/                  ← Arquitectura Legacy y Target
│   ├── migration/                     ← Estrategia, estado, trazabilidad, gaps
│   ├── modules/restaurante/           ← Documentación del módulo Restaurante
│   ├── database/                      ← Base de datos, SPs, mapeo
│   ├── standards/                     ← Estándares y convenciones
│   └── glossary.md
└── documento de diseño/               ← Guía de diseño UI/UX (PDF)
```

---

## Inventario Rápido

| Artefacto | Cantidad | Ubicación |
|---|---|---|
| Proyectos VB6 (.vbp) | 7 | `legacy-restaurant/restaurant-vb6/` |
| Formularios (.frm) | 401 | `legacy-restaurant/restaurant-vb6/Formularios/` |
| Módulos (.bas) | 32 | `legacy-restaurant/restaurant-vb6/Modulos/` |
| Clases (.cls) | 10 | `legacy-restaurant/restaurant-vb6/Clases/` |
| Reportes Crystal Reports (.dsr) | 206 | `legacy-restaurant/restaurant-vb6/Reportes/` |
| Tablas SQL Server | 126 | `legacy-restaurant/database-sql-server/1. Estructura.sql` |
| Vistas SQL Server | 105 | `legacy-restaurant/database-sql-server/4. Vistas.sql` |
| Stored Procedures | 105+ | `legacy-restaurant/database-sql-server/5. SP.sql` |
| Scripts por país | 6 | `legacy-restaurant/database-sql-server/opcionales/` |

---

## Documentación Principal

| Documento | Descripción |
|---|---|
| [legacy-restaurant/README.md](legacy-restaurant/README.md) | Análisis técnico completo del Legacy |
| [docs/README.md](docs/README.md) | Índice de toda la documentación |
| [docs/architecture/legacy-architecture.md](docs/architecture/legacy-architecture.md) | Arquitectura Legacy |
| [docs/architecture/target-architecture.md](docs/architecture/target-architecture.md) | Arquitectura objetivo .NET 8 |
| [docs/architecture/architecture-decisions.md](docs/architecture/architecture-decisions.md) | ADRs — Decisiones arquitectónicas |
| [docs/migration/migration-status.md](docs/migration/migration-status.md) | Estado de migración |
| [docs/migration/traceability-matrix.md](docs/migration/traceability-matrix.md) | Trazabilidad Legacy → .NET 8 |
| [docs/migration/known-gaps.md](docs/migration/known-gaps.md) | Brechas identificadas |
| [docs/modules/restaurante/README.md](docs/modules/restaurante/README.md) | Módulo Restaurante |
| [docs/database/legacy-database.md](docs/database/legacy-database.md) | Base de datos Legacy |
| [docs/database/stored-procedures.md](docs/database/stored-procedures.md) | Stored Procedures |
| [docs/glossary.md](docs/glossary.md) | Glosario |
| [.github/copilot-instructions.md](.github/copilot-instructions.md) | Instrucciones para GitHub Copilot |

---

## Convenciones de Estado

| Indicador | Significado |
|---|---|
| `CONFIRMED` | Verificado con evidencia del repositorio |
| `PARTIAL` | Información incompleta |
| `UNKNOWN` | No determinable |
| `NOT_STARTED` | No iniciado |
| `IN_PROGRESS` | En curso |
| `MIGRATED` | Migrado, pendiente validación |
| `COMPLETED` | Migrado y validado |
| `BLOCKED` | Bloqueado |

---

## Regla Fundamental

> **El código Legacy VB6 es la fuente de verdad funcional** hasta que una decisión de negocio o arquitectura documentada establezca explícitamente un comportamiento diferente.

Antes de cualquier trabajo de migración, leer:
1. [legacy-restaurant/README.md](legacy-restaurant/README.md)
2. [docs/modules/restaurante/README.md](docs/modules/restaurante/README.md)
3. [docs/migration/traceability-matrix.md](docs/migration/traceability-matrix.md)
4. [.github/copilot-instructions.md](.github/copilot-instructions.md)

---

*Toda afirmación está basada en evidencia del repositorio. Lo no determinable se indica como `UNKNOWN`.*

