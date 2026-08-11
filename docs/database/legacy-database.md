# Base de Datos Legacy — INFOREST SQL Server

> Status: CONFIRMED — basado en análisis directo de `legacy-restaurant/database-sql-server/`

---

## Motor y Versión

| Campo | Valor |
|---|---|
| Motor | Microsoft SQL Server |
| Versión | UNKNOWN (no especificada en scripts) |
| Collation | Modern_Spanish_CI_AS |
| Provider ADO | SQLOLEDB |

---

## Bases de Datos

| Base | Propósito | Script |
|---|---|---|
| `INFOREST` | Base principal de operación | `1. Estructura.sql`, `5. SP.sql`, etc. |
| `INFSEGURIDAD` | Seguridad, usuarios, auditoría | `opcionales/Seguridad.sql` |
| `ALMACEN` | Inventario y almacén | (BD separada, referenciada desde código) |
| `CENTRALDELIVERY` | Delivery centralizado | (BD separada, referenciada desde código) |
| `FACTURACION` | Facturación electrónica | `8. InfoFact.sql` |

---

## Scripts SQL

| Script | Propósito | Objetos creados |
|---|---|---|
| `1. Estructura.sql` | Tablas de la BD INFOREST | 126 tablas |
| `2. Columns.sql` | Alteraciones de columnas | Modificaciones |
| `3. PK.sql` | Primary keys y constraints | Constraints |
| `4. Vistas.sql` | Vistas de lectura | 105 vistas |
| `5. SP.sql` | Stored procedures y funciones | 105+ SPs, 2 funciones |
| `6. Actualiza.sql` | Actualizaciones de datos | Scripts puntales |
| `8. InfoFact.sql` | Vistas para FE InfoFact | Vistas FE |
| `opcionales/Seguridad.sql` | BD INFSEGURIDAD | Tablas seguridad |
| `opcionales/scriptPeruAlIniciar.sql` | Configuración Perú | Scripts Perú |
| `opcionales/scriptChileAlIniciar.sql` | Configuración Chile | Scripts Chile |
| `opcionales/scriptBoliviaAlIniciar.sql` | Configuración Bolivia | Scripts Bolivia |
| `opcionales/scriptEcuadorAlIniciar.sql` | Configuración Ecuador | Scripts Ecuador |
| `opcionales/Ejecuta Columnas_Ecuador.sql` | Columnas Ecuador | Scripts Ecuador |
| `opcionales/scriptEspanaAlIniciar.sql` | Configuración España | Scripts España |
| `opcionales/scriptArgentinaAlIniciar.sql` | Configuración Argentina | Scripts Argentina |

---

## Inventario de Tablas por Dominio

> Total: 126 tablas. Status: CONFIRMED.

### Pedidos y Ventas

| Tabla | Descripción |
|---|---|
| `MPEDIDO` | Cabecera de pedido (mesa, mozo, motorizado, estado, canal, delivery, hotel, reserva) |
| `DPEDIDO` | Detalle de pedido (producto, precios×3impuestos, cantidad, área, estado ítem, comanda) |
| `CPEDIDO` | Detalle de combos dentro de un pedido |
| `APEDIDO` | Ítems anulados/eliminados del pedido (auditoría) |
| `MDOCUMENTO` | Cabecera de documento de venta (boleta/factura/ticket) |
| `DDOCUMENTO` | Detalle de documento de venta |
| `MNOTACREDITO` | Cabecera de nota de crédito |
| `DNOTACREDITO` | Detalle de nota de crédito |
| `LOG_PEDIDO_DOCUMENTO` | Log relación pedido→documento |
| `TPEDIDO` | Catálogo/referencia de pedidos |
| `TPEDIDOMESA` | Relación pedido↔mesa (multi-mesa) |

### Caja, Turno y Pagos

| Tabla | Descripción |
|---|---|
| `MTURNO` | Turno de caja (efectivo N/E, cheque, pagaré, 8 tarjetas con propinas) |
| `TCAJA` | Configuración de caja (30+ flags de comportamiento) |
| `TCAJACANALVENTA` | Canales de venta habilitados por caja |
| `TCAJAORIGEN_BLOQUEO` | Bloqueos de origen por caja |
| `TCAJATERMINAL` | Terminales de pago por caja |
| `DPAGODOCUMENTO` | Pagos registrados por documento |
| `DPAGODOCUMENTO_VC` | Pagos en vale de consumo |
| `DPAGOTARJETA` | Datos extendidos de pago con tarjeta |
| `DPREPAGO` | Prepagos contra pedido |
| `MEGRESO` | Egresos de caja |
| `MINGRESO` | Ingresos de caja (anticipos) |
| `DINGRESO` | Detalle de ingresos |
| `MCIERRE` | Cierre contable por período |
| `TDIACONTABLE` | Día contable activo |
| `INTEGRACION_CASHDRO` | Transacciones CashDro |
| `PEDIDO_PAGOEFECTIVO` | Relación pedido↔Pago Efectivo |

### Maestros Comerciales

| Tabla | Descripción |
|---|---|
| `TPRODUCTO` | Maestro de productos (15 flags impuesto, 5 precios canal, imagen, combo, área cocina) |
| `TGRUPO` | Grupos de productos |
| `TSUBGRUPO` | Subgrupos de productos |
| `TPROPIEDAD` | Propiedades/modificadores de producto |
| `TPRODUCTOPROPIEDAD` | Relación producto↔propiedades |
| `TCOMBOPROPIEDAD` | Propiedades aplicadas a combos |
| `TCOMBO` | Definición de combos |
| `TPRODUCTOXPRODUCTO` | Sustitutos/equivalencias entre productos |
| `TPRODUCTOAREA` | Áreas de impresión de cocina por producto |
| `TINSUMO` | Insumos/ingredientes para descargo de stock |
| `TOFERTA` | Ofertas comerciales |
| `TPROGRAMAPRECIOS_CAB` | Cabecera de programación de precios |

### Clientes y Delivery

| Tabla | Descripción |
|---|---|
| `TCLIENTE` | Maestro de clientes |
| `TDELIVERY` | Datos de delivery por cliente |
| `TDELIVERYCLIENTE` | Relación delivery↔cliente |
| `TDELIVERYINVITADO` | Datos de cliente invitado delivery |
| `TPARIENTE` | Contactos/parientes del cliente |
| `TCLIENTEPRODUCTO` | Preferencias de cliente |
| `TMOTORIZADODATOS` | Datos de motorizados |

### Configuración del Negocio

| Tabla | Descripción |
|---|---|
| `TPARAMETRO` | Parámetros del sistema (100+ flags de configuración) |
| `TCOMPANIA` | Datos de la empresa |
| `TLOCAL` | Locales del negocio (multi-local) |
| `TTIENDA` | Tiendas/establecimientos |
| `TCANALVENTA` | Canales de venta configurados |
| `TTIPOPEDIDODETALLE` | Tipos de pedido (local, delivery, llevar) |
| `TTIPODOCUMENTO` | Tipos de documentos de venta |
| `TTIPODOCUMENTOIMPRESORA` | Config de documentos por impresora/caja |
| `TTIPOCAMBIO` | Tipos de cambio de moneda |
| `TMOTIVODESCUENTO` | Motivos de descuento |
| `TMESA` | Mesas del restaurante |
| `TAREA` | Áreas del restaurante (salón, delivery, etc.) |
| `TAREAIMPRESORA` | Áreas de impresión configuradas |
| `TAREAPANTALLA` | Áreas para pantalla/KDS |
| `TAREAPANTALLA1` | Extensión de áreas pantalla |
| `TAREAPANTALLADESPACHO` | Áreas de despacho |
| `TAREACHEF` | Áreas de chef control |
| `TAREASUBGRUPO` | Subgrupos por área |
| `TIMPRESORA` | Maestro de impresoras |
| `TMENSAJE` | Mensajes del sistema |
| `TMENSAJECOCINA` | Mensajes de cocina |
| `TMENSAJEUSUARIO` | Mensajes a usuarios |

### Seguridad y Auditoría

| Tabla | Descripción |
|---|---|
| `TUSUARIO` | Usuarios del sistema |
| `TGRUPOUSUARIO` | Grupos de usuarios |
| `TGRUPOACCESO` | Grupos de acceso a módulos |
| `TMODULO` | Módulos del sistema |
| `TACCESO` | Accesos por usuario/grupo |
| `TACCESOENVIA` | Control de envíos por acceso |
| `TOPERADOR` | Operadores |
| `TLOG` | Log general del sistema |
| `TLOG_IMPRESION` | Log de impresiones |
| `THARDKEY` | Control de licencias/hardkey |

### KDS y Cocina

| Tabla | Descripción |
|---|---|
| `DPEDIDOKDS` | Detalle de pedidos KDS |
| `TORIGENCODIGOCONTROL` | Origen de código de control |

### Reservas y Hotel

| Tabla | Descripción |
|---|---|
| `TRESERVA` | Reservas del restaurante |
| `TCOMPANIA` | Compañías (hotel) |

### Inventario y Almacén

| Tabla | Descripción |
|---|---|
| `TINSUMO` | Insumos configurados |

### Otras Tablas

| Tabla | Descripción |
|---|---|
| `MENVIO` | Envíos |
| `MPROPINA` | Propinas |
| `MGUIATRANSPORTE` | Cabecera guías de transporte |
| `DGUIATRANSPORTE` | Detalle guías de transporte |
| `TIMPORTACION` | Importaciones |
| `TIMPORTACIONLOG` | Log de importaciones |
| `TTRAMITE` | Trámites |
| `TSOLICITUD` | Solicitudes |
| `TSOLICITUDDETALLE` | Detalle de solicitudes |
| `TDIACONTABLE` | Día contable |
| `TTTABLA` | Tabla genérica de parámetros |

---

## Convenciones de Nomenclatura de Tablas

| Prefijo | Tipo | Descripción |
|---|---|---|
| `M` | Master/Movement | Tablas de movimiento/transacción (cabeceras) |
| `D` | Detail | Tablas de detalle de transacciones |
| `T` | Table/Type | Tablas maestras (catálogos, configuración) |
| `A` | Audit | Tablas de auditoría/anulados |
| `LOG_` | Log | Tablas de log |
| `INTEGRACION_` | Integration | Tablas de integraciones externas |

---

## Convenciones de Columnas

| Prefijo | Tipo VB6/SQL | Descripción |
|---|---|---|
| `t` | varchar/nvarchar | Texto/String |
| `n` | float/decimal/int | Numérico |
| `l` | bit | Lógico/Boolean |
| `f` | smalldatetime/datetime | Fecha/Fecha-Hora |
| `c` | char | Carácter fijo |

---

## Restricciones de Integridad Conocidas

- Primary Keys definidas en `3. PK.sql`
- Foreign Keys: PARTIAL — algunas relaciones son implícitas en código, no siempre FK explícitas
- Índices: UNKNOWN — no hay script específico de índices

---

## Referencias

- `legacy-restaurant/database-sql-server/` — scripts fuente
- [Stored Procedures](stored-procedures.md)
- [Mapeo Legacy→Target](mapping.md)
