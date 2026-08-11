# INFOREST — Sistema Moderno .NET 8

> Migración de INFOREST VB6 → .NET 8 (C# 12)
>
> Estado: **Fase 2 — Proyecto base creado. Migración funcional en preparación.**
>
> Última actualización: 2026-08-11

---

## Estructura de la Solución

```
Inforest.sln
├── src/
│   ├── Inforest.Domain/           ← Entidades, reglas de negocio, interfaces (cero dependencias externas)
│   ├── Inforest.Application/      ← Casos de uso, Commands, Queries, interfaces de servicios
│   ├── Inforest.Infrastructure/   ← Dapper/SQL Server, BCrypt, Serilog, hardware adapters
│   └── Inforest.Desktop/          ← WinForms .NET 8 (net8.0-windows)
└── tests/
    ├── Inforest.Domain.Tests/      ← 8 tests pasando
    ├── Inforest.Application.Tests/
    └── Inforest.Infrastructure.Tests/
```

## Requisitos

- .NET 8 SDK
- SQL Server (base de datos INFOREST)
- Windows OS (para Inforest.Desktop — hardware POS intensivo)

## Build y Tests

```bash
cd modern-net8
dotnet build
dotnet test
```

## Configuración

Editar `src/Inforest.Desktop/appsettings.json`:

```json
{
  "ConnectionStrings": {
    "Inforest": "Server=localhost;Database=INFOREST;Trusted_Connection=true;TrustServerCertificate=true;"
  }
}
```

Para desarrollo, usar User Secrets:

```bash
cd src/Inforest.Desktop
dotnet user-secrets set "ConnectionStrings:Inforest" "Server=...;"
```

## Antes de implementar cualquier funcionalidad

**Leer obligatoriamente:**

1. [legacy-restaurant/README.md](../legacy-restaurant/README.md) — análisis técnico del Legacy
2. [docs/migration/legacy-inventory.md](../docs/migration/legacy-inventory.md) — inventario completo Legacy
3. [docs/migration/database/sql-inventory.md](../docs/migration/database/sql-inventory.md) — inventario SQL
4. [docs/migration/business-rules.md](../docs/migration/business-rules.md) — reglas de negocio
5. [docs/migration/traceability/business-logic-matrix.md](../docs/migration/traceability/business-logic-matrix.md)
6. [docs/migration/migration-order.md](../docs/migration/migration-order.md) — orden de migración
7. [docs/standards/dotnet.md](../docs/standards/dotnet.md) — estándares de código
8. [.github/copilot-instructions.md](../.github/copilot-instructions.md) — instrucciones Copilot

## Decisiones Tecnológicas (Aceptadas)

| Decisión | Tecnología | ADR |
|---|---|---|
| UI | WinForms .NET 8 | ADR-001 |
| Base de datos | SQL Server + Dapper | ADR-002 |
| Arquitectura | Clean Architecture + CQRS | ADR-003 |
| Estrategia migración | Strangler Fig Pattern | ADR-004 |
| Credenciales | appsettings.json + env vars | ADR-005 |
| Seguridad | BCrypt + RBAC | ADR-006 |
| Reportes | FastReport .NET | ADR-007 |
| Multi-país/local | Feature flags + tenant_id | ADR-008 |

Ver todos: [docs/architecture/architecture-decisions.md](../docs/architecture/architecture-decisions.md)

## Estado de Migración

| Componente | Estado |
|---|---|
| Solución .NET 8 | ✅ CREADA |
| Domain (entidades base) | ✅ IN_PROGRESS |
| Application (interfaces base) | ✅ IN_PROGRESS |
| Infrastructure (DI, logging) | ✅ IN_PROGRESS |
| Desktop (WinForms, DI setup) | ✅ IN_PROGRESS |
| Tests base | ✅ 8 tests pasando |
| Módulos funcionales migrados | 0 / 7 |

Ver detalle: [docs/migration/migration-status.md](../docs/migration/migration-status.md)
