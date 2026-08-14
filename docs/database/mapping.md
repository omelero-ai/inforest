# Mapeo de Base de Datos — Legacy → Target

> Status: NOT_STARTED — no existe definición de BD Target.

---

## Estado

El mapeo Legacy → Target no puede completarse hasta que se defina la base de datos Target (ver [ADR-002](../architecture/architecture-decisions.md#adr-002)).

---

## Estructura del Mapeo (Plantilla)

Cuando se defina la BD Target, este documento debe incluir:

| Legacy Table | Legacy Type | Target Table | Target Type | Cambios | Estado |
|---|---|---|---|---|---|
| `MPEDIDO` | SQL Server | `orders` (ej.) | Por definir | Renombre, campos | NOT_STARTED |
| `DPEDIDO` | SQL Server | `order_items` (ej.) | Por definir | Renombre, campos | NOT_STARTED |
| ... | | | | | |

---

## Consideraciones de Mapeo

### Campos que cambiarán con certeza

| Campo Legacy | Razón del cambio | Campo Target propuesto |
|---|---|---|
| Nombre de tabla (ej. `MPEDIDO`) | Convención de nombres moderna | UNKNOWN |
| Tipos `float` para precios | Usar `decimal` en lugar de `float` | `decimal(18,4)` |
| `smalldatetime` | Usar `datetime2` | `datetime2` |
| Collation `Modern_Spanish_CI_AS` | Revisar para multi-idioma | UNKNOWN |

### Campos sensibles a preservar

| Campo Legacy | Razón |
|---|---|
| `TPARAMETRO` — 100+ flags | Comportamiento del sistema depende de estos flags |
| `TCAJA` — 30+ flags | Comportamiento por caja |
| `TPRODUCTO.lImpuesto1..15` | Modelo tributario multi-país |
| `TPRODUCTO.nPrecio*` (5 canales) | Precios multi-canal |

### Multi-Local

En la BD Target, todas las tablas maestras deben incluir `local_id` para unicidad en contexto multi-local.

Ejemplo:
```sql
-- Legacy: PK solo por código
PRIMARY KEY (tCodigoProducto)

-- Target: PK incluyendo local_id
PRIMARY KEY (local_id, product_code)
```

---

## Referencias

- [Legacy Database](legacy-database.md)
- [Target Database](target-database.md)
- [ADR-002](../architecture/architecture-decisions.md#adr-002)
- [ADR-008](../architecture/architecture-decisions.md#adr-008) (Multi-país/local)
