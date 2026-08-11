# Sistema INFOREST — Implementación .NET 8

> **Status: NOT_STARTED**
>
> Esta carpeta contendrá la nueva implementación .NET 8 del sistema INFOREST.
> Actualmente está vacía — no existe código .NET 8 en el repositorio.

---

## Antes de comenzar a implementar

**Leer obligatoriamente:**

1. [README.md del repositorio](../README.md)
2. [legacy-restaurant/README.md](../legacy-restaurant/README.md) — análisis técnico del Legacy
3. [docs/architecture/target-architecture.md](../docs/architecture/target-architecture.md) — arquitectura objetivo
4. [docs/architecture/architecture-decisions.md](../docs/architecture/architecture-decisions.md) — decisiones ADRs
5. [docs/migration/migration-rules.md](../docs/migration/migration-rules.md) — reglas del proceso
6. [.github/copilot-instructions.md](../.github/copilot-instructions.md) — instrucciones para Copilot

---

## Estructura Propuesta (Pendiente de Decisión)

```
modern-net8/
├── src/
│   ├── Inforest.Domain/
│   ├── Inforest.Application/
│   ├── Inforest.Infrastructure/
│   └── Inforest.Presentation/
└── tests/
    ├── Inforest.Domain.Tests/
    ├── Inforest.Application.Tests/
    └── Inforest.Integration.Tests/
```

> La estructura definitiva se define en [ADR-003](../docs/architecture/architecture-decisions.md#adr-003).

---

## Estado

| Componente | Estado |
|---|---|
| Proyecto .NET 8 | NOT_STARTED |
| Base de datos Target | NOT_STARTED |
| Módulos migrados | 0 / 7 |

