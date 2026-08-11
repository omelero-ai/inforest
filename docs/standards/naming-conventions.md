# Convenciones de Nombres — INFOREST

> Convenciones de nombres para la nueva implementación .NET 8 y su relación con los nombres del Legacy.

---

## .NET 8 — Nombres de Proyecto

| Proyecto | Descripción |
|---|---|
| `Inforest.Domain` | Capa de dominio |
| `Inforest.Application` | Capa de aplicación |
| `Inforest.Infrastructure` | Capa de infraestructura |
| `Inforest.Api` | API (si aplica) |
| `Inforest.Desktop` | Aplicación desktop (si aplica) |

---

## Correspondencia Legacy → .NET

### Módulos VB6 → Servicios/Handlers

| VB6 | .NET 8 | Ejemplo |
|---|---|---|
| `mod<Nombre>.bas` | `<Nombre>Service.cs` | `modProducto` → `ProductoService` |
| `frm<Nombre>.frm` | `<Nombre>Controller.cs` | `frmVenta` → `VentaController` |
| `cls<Nombre>.cls` | `<Nombre>.cs` | `clsAlmacen` → `Almacen.cs` |

### Tablas SQL → Entidades

| SQL Legacy | .NET 8 | Notas |
|---|---|---|
| `MPEDIDO` | `Pedido` / `Order` | Cabecera → entidad raíz |
| `DPEDIDO` | `DetallePedido` / `OrderItem` | Detalle → entidad hijo |
| `TPRODUCTO` | `Producto` / `Product` | Maestro → entidad |
| `TCAJA` | `Caja` / `CashRegister` | Config → entidad config |
| `MTURNO` | `Turno` / `Shift` | Turno → entidad |

> El idioma (español/inglés) para los nombres en .NET debe definirse como ADR antes de comenzar.

---

## Convenciones de BD Target

| Tipo | Legacy | Target (propuesto) |
|---|---|---|
| Tablas | `UPPER_CASE` | `snake_case` o `PascalCase` |
| Columnas | `tNombre`, `nCantidad`, `lActivo` | `nombre`, `cantidad`, `activo` |
| PKs | `tCodigo...` | `id` (UUID o int) |
| FKs | Implícitas | Explícitas con nombre estándar |
| Índices | Sin naming standard | `ix_<tabla>_<columna>` |

---

## Reglas Generales

1. Inglés o Español consistente — no mezclar por módulo
2. PascalCase para clases, métodos, propiedades
3. camelCase para variables locales y parámetros
4. _camelCase para campos privados
5. Prefijo `I` para interfaces
6. Sufijo `Async` para métodos async
7. Sufijo `Service`, `Repository`, `Controller`, `Handler` según rol

---

## Referencias

- [Coding Standards](coding-standards.md)
- [ADR-003](../architecture/architecture-decisions.md#adr-003)
