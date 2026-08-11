# Base de Datos — INFOREST

> Documentación de las bases de datos del sistema INFOREST.

---

## Documentos

| Documento | Descripción |
|---|---|
| [legacy-database.md](legacy-database.md) | Base de datos Legacy SQL Server (CONFIRMED) |
| [target-database.md](target-database.md) | Base de datos Target (NOT_STARTED) |
| [mapping.md](mapping.md) | Mapeo Legacy → Target |
| [stored-procedures.md](stored-procedures.md) | Inventario de Stored Procedures |

---

## Resumen

| Aspecto | Legacy | Target |
|---|---|---|
| Motor | SQL Server | UNKNOWN |
| Base principal | INFOREST | UNKNOWN |
| Tablas | 126 | UNKNOWN |
| Vistas | 105 | UNKNOWN |
| Stored Procedures | 105+ | UNKNOWN |
| Scripts país | 6 | UNKNOWN |

---

## Bases de Datos Legacy

| Base | Propósito | Acceso desde |
|---|---|---|
| `INFOREST` | Operación principal del restaurante | Todos los ejecutables |
| `INFSEGURIDAD` | Usuarios, grupos, auditoría de acceso | Todos los ejecutables |
| `ALMACEN` | Inventario, kardex, insumos | InfoRest, Administracion |
| `CENTRALDELIVERY` | Delivery centralizado multi-local | Despachador |
| `FACTURACION` | Facturación electrónica | InfoRest, CajaRapida |
