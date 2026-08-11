# Instrucciones para GitHub Copilot — INFOREST Migration

## Contexto del Proyecto

Este repositorio contiene la migración de **INFOREST**, un sistema ERP gastronómico desarrollado en **Visual Basic 6 (VB6)** hacia **.NET 8 (C#)**.

---

## Definiciones Fundamentales

```
LEGACY   = Código VB6 en legacy-restaurant/restaurant-vb6/
           Fuente de comportamiento funcional del sistema

TARGET   = Nueva implementación en modern-net8/
           Implementación moderna objetivo

DATABASE = Base de datos SQL Server documentada en legacy-restaurant/database-sql-server/
           Contratos de datos del sistema

DOCUMENTATION = Documentación en docs/ y legacy-restaurant/README.md
                Fuente de contexto arquitectónico y reglas de negocio

TRACEABILITY = docs/migration/traceability-matrix.md
               Evidencia de equivalencia entre Legacy y .NET 8
```

---

## Regla Crítica

> **El código Legacy VB6 es la fuente de verdad funcional hasta que una decisión de negocio o arquitectura documentada establezca explícitamente un comportamiento diferente.**

---

## Flujo Obligatorio Antes de Migrar

```
1.  Leer docs/README.md
2.  Leer legacy-restaurant/README.md
3.  Leer la documentación del módulo: docs/modules/<modulo>/README.md
4.  Leer el código Legacy correspondiente en legacy-restaurant/restaurant-vb6/
5.  Identificar todas las reglas de negocio (ver docs/migration/business-rules.md)
6.  Identificar todas las dependencias del módulo
7.  Identificar tablas y stored procedures usados
8.  Buscar implementación existente en modern-net8/ (si existe)
9.  Comparar Legacy vs .NET existente — identificar diferencias
10. Revisar docs/migration/traceability-matrix.md
11. Revisar docs/migration/known-gaps.md
12. SOLO ENTONCES proceder con la migración
```

---

## Flujo de Migración

```
1.  Seleccionar módulo
2.  Leer documentación completa del módulo
3.  Analizar código Legacy (formularios, módulos, clases)
4.  Identificar funcionalidades a migrar
5.  Identificar reglas de negocio
6.  Identificar dependencias (BD, SPs, hardware, integraciones)
7.  Identificar base de datos (tablas, vistas, SPs)
8.  Buscar implementación existente en .NET
9.  Comparar Legacy vs .NET actual
10. Identificar diferencias y gaps
11. Implementar migración
12. Escribir pruebas unitarias e integración
13. Validar comportamiento contra Legacy
14. Actualizar docs/migration/traceability-matrix.md
15. Actualizar docs/modules/<modulo>/migration-status.md
16. Actualizar docs/migration/migration-status.md
```

---

## Lo que NO debes hacer

- ❌ NO modificar código VB6 existente
- ❌ NO modificar scripts SQL funcionales
- ❌ NO inventar comportamiento no evidenciado en el Legacy
- ❌ NO eliminar funcionalidades Legacy sin autorización explícita
- ❌ NO cambiar nombres de tablas del Legacy sin documentar la decisión (ADR)
- ❌ NO asumir que código moderno = migración completa
- ❌ NO comenzar a codificar sin leer la documentación primero
- ❌ NO omitir la actualización de trazabilidad después de migrar

---

## Estructura de Archivos Relevante

```
legacy-restaurant/
  README.md                    ← Análisis técnico completo del Legacy (LEER PRIMERO)
  restaurant-vb6/
    Formularios/               ← 401 formularios VB6 (.frm)
    Modulos/                   ← 32 módulos de negocio (.bas)
    Clases/                    ← 10 clases (.cls)
    Reportes/                  ← 206 reportes Crystal Reports
  database-sql-server/
    1. Estructura.sql          ← 126 tablas
    4. Vistas.sql              ← 105 vistas
    5. SP.sql                  ← 105+ stored procedures
    opcionales/                ← Scripts por país

modern-net8/                   ← Target .NET 8 (NO INICIADO)

docs/
  architecture/
    legacy-architecture.md     ← Arquitectura Legacy
    target-architecture.md     ← Arquitectura .NET 8
    architecture-decisions.md  ← ADRs
  migration/
    traceability-matrix.md     ← Mapa Legacy → .NET 8
    migration-status.md        ← Estado actual
    known-gaps.md              ← Brechas identificadas
    business-rules.md          ← Reglas de negocio
  modules/
    restaurante/README.md      ← Módulo principal
  database/
    legacy-database.md         ← Inventario BD Legacy
    stored-procedures.md       ← Inventario SPs
```

---

## Módulos del Sistema

| Módulo | Ejecutable VB6 | Formularios | Módulos BAS | Estado |
|---|---|---|---|---|
| Punto de Venta | `InfoRest.exe` | ~120 | 18 | NOT_STARTED |
| Caja Rápida | `CajaRapida.exe` | ~100 | 16 | NOT_STARTED |
| Adición | `Adicion.exe` | 37 | 12 | NOT_STARTED |
| Administración | `Administracion.exe` | 151 | 13 | NOT_STARTED |
| Consultas | `Consulta.exe` | 124 | 15 | NOT_STARTED |
| Despachador | `Despachador.exe` | 25 | 10 | NOT_STARTED |
| Motorizados | `Motorizado.exe` | 2 | 7 | NOT_STARTED |

---

## Bases de Datos del Sistema Legacy

| Base de Datos | Propósito |
|---|---|
| `INFOREST` | Base principal de operación del restaurante |
| `INFSEGURIDAD` | Seguridad, usuarios, auditoría |
| `ALMACEN` | Gestión de inventario y almacén |
| `CENTRALDELIVERY` | Operaciones de delivery centralizado |
| `FACTURACION` | Facturación electrónica |

---

## Reglas para Manejo de Trazabilidad

Cada componente migrado debe tener entrada en `docs/migration/traceability-matrix.md` con:

| Legacy | Tipo | .NET 8 | Tipo | Estado | Evidencia |
|---|---|---|---|---|---|
| `frmVenta.frm` | Form | `VentaController` | Controller | IN_PROGRESS | src/... |

Estados válidos: `NOT_STARTED` | `ANALYSIS` | `IN_PROGRESS` | `MIGRATED` | `VALIDATING` | `COMPLETED` | `BLOCKED` | `NOT_APPLICABLE` | `UNKNOWN`

---

## Reglas para Manejo de Reglas de Negocio

Formato obligatorio para documentar reglas de negocio:

```
BR-XXX
Nombre:
Origen: Legacy/<archivo>
Archivo: legacy-restaurant/restaurant-vb6/<ruta>
Procedimiento/Función:
Descripción:
Condición:
Resultado:
Excepciones:
Destino .NET: [componente target]
Estado: NOT_STARTED | IN_PROGRESS | MIGRATED | COMPLETED
```

---

## Convención de ADR (Architecture Decision Records)

Cuando debas tomar una decisión arquitectónica, documentarla como:

```
ADR-XXX

Título:
Estado: Proposed | Accepted | Deprecated | Superseded
Contexto:
Problema:
Decisión:
Alternativas consideradas:
Consecuencias:
Fecha:
Evidencia:
```

Archivo: `docs/architecture/architecture-decisions.md`

---

## Criterio de Migración Completa

Una migración se considera COMPLETA únicamente cuando:

```
✓ Funcionalidad Legacy identificada
✓ Reglas de negocio documentadas
✓ Dependencias identificadas
✓ Comportamiento de base de datos verificado
✓ Implementación .NET 8 existente
✓ Pruebas unitarias escritas
✓ Pruebas de integración escritas
✓ Comportamiento validado contra Legacy
✓ Trazabilidad actualizada
✓ Documentación del módulo actualizada
```

**Código moderno ≠ migración completa.**

---

## Contacto con el Legacy

Al trabajar con código VB6:
- Los formularios `.frm` contienen UI y lógica mezclada
- Los módulos `.bas` contienen lógica de negocio compartida
- Las clases `.cls` encapsulan objetos de dominio
- `modDeclaracion.bas` contiene 543 variables globales — revisar antes de asumir ausencia de estado
- `modProcedimiento.bas` contiene utilidades núcleo del sistema
- `TPARAMETRO` es la tabla maestra de configuración del sistema — muchos comportamientos dependen de sus flags
- `TCAJA` controla el comportamiento específico de cada caja con 30+ flags

---

*Estas instrucciones son mandatorias para cualquier trabajo de migración en este repositorio.*
