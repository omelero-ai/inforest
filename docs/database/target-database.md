# Base de Datos Target — INFOREST .NET 8

> Status: NOT_STARTED — no existe definición de base de datos Target.
>
> Este documento se completa cuando se tomen las decisiones arquitectónicas correspondientes (ver [ADR-002](../architecture/architecture-decisions.md#adr-002)).

---

## Estado

No existe código .NET 8 ni definición de base de datos Target en el repositorio.

La definición de la base de datos Target requiere previamente:

1. Decisión sobre motor de BD ([ADR-002](../architecture/architecture-decisions.md#adr-002))
2. Decisión sobre patrón arquitectónico ([ADR-003](../architecture/architecture-decisions.md#adr-003))
3. Estrategia de migración definida ([migration-strategy.md](../migration/migration-strategy.md))

---

## Consideraciones para el Diseño

### Renombramiento de Tablas

Si se decide renombrar las tablas Legacy (que usan prefijos M/D/T/A), se debe:
- Documentar el mapeo en [mapping.md](mapping.md)
- Actualizar la [matriz de trazabilidad](../migration/traceability-matrix.md)
- Crear un ADR para la decisión

### Multi-Local

La BD Target debe soportar `local_id` para mantener unicidad en arquitectura multi-local, dado que el Legacy maneja múltiples establecimientos.

### Multi-País

Deben mantenerse las configuraciones por país (impuestos, datos fiscales, FE) que actualmente se manejan con scripts SQL opcionales.

---

## Placeholder

Este documento se completará cuando se inicie el diseño de la BD Target.

---

## Referencias

- [Legacy Database](legacy-database.md)
- [Mapeo](mapping.md)
- [ADR-002](../architecture/architecture-decisions.md#adr-002)
