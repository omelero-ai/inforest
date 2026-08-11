# Base de Datos — Sistema Legacy Restaurant (SQL Server)

## Información General

| Campo | Valor |
|---|---|
| **Motor** | Microsoft SQL Server (SQLOLEDB) |
| **Bases de datos principales** | `INFOREST`, `INFSEGURIDAD`, `ALMACEN`, `BDEFACT`, `INFORESTCD` (Central Delivery), `WEBAPP` |
| **Collation** | `Modern_Spanish_CI_AS` |
| **Tablas en INFOREST** | 126 (verificado en `1. Estructura.sql`) |
| **Vistas en INFOREST** | 105 (verificado en `4. Vistas.sql`) |
| **Procedimientos almacenados** | 105 (verificado en `5. SP.sql`) |
| **Funciones escalares** | 2 (`fn_cortesia_calculo`, `CreaTabla`) |
| **Triggers en scripts** | 0 detectados |
| **Integridad referencial** | Sin FK explícitas detectadas en scripts (mantenida por aplicación VB6 y SPs) |

---

## Scripts de despliegue

| Nº | Archivo | Descripción | Orden ejecución |
|---|---|---|---|
| 1 | `1. Estructura.sql` | CREATE TABLE de las 126 tablas principales de INFOREST | 1° |
| 2 | `2. Columns.sql` | ALTER TABLE / ADD COLUMN — ajustes y extensiones de columnas | 2° |
| 3 | `3. PK.sql` | CREATE PRIMARY KEY y constraints | 3° |
| 4 | `4. Vistas.sql` | CREATE VIEW — 105 vistas de lectura operacional | 4° |
| 5 | `5. SP.sql` | CREATE PROCEDURE / CREATE FUNCTION — 105 SPs y 2 funciones | 5° |
| 6 | `6. Actualiza.sql` | Scripts de actualización puntual de datos (registros semilla) | 6° |
| 7 | `8. InfoFact.sql` | Vistas adicionales para integración de facturación electrónica InfoFact | 7° |
| — | `opcionales/Seguridad.sql` | Crea base de datos `INFSEGURIDAD` completa con tablas de auditoría y licencias | Independiente |
| — | `opcionales/scriptPeruAlIniciar.sql` | Configuración inicial específica para Perú (SUNAT/IGV) | Por país |
| — | `opcionales/scriptChileAlIniciar.sql` | Configuración para Chile (IVA) | Por país |
| — | `opcionales/scriptBoliviaAlIniciar.sql` | Configuración para Bolivia (SIN, Código de Control) | Por país |
| — | `opcionales/scriptEcuadorAlIniciar.sql` | Configuración para Ecuador (SRI) | Por país |
| — | `opcionales/Ejecuta Columnas_Ecuador.sql` | Columnas adicionales para Ecuador | Por país |
| — | `opcionales/scriptEspanaAlIniciar.sql` | Configuración para España (IVA) | Por país |
| — | `opcionales/scriptArgentinaAlIniciar.sql` | Configuración para Argentina (AFIP, impresora fiscal) | Por país |

---

## Base de datos INFOREST

Base de datos operativa principal del sistema de restaurante.

### Tablas por dominio

Ver `legacy/restaurant/README.md` Sección 4.3 para el inventario completo de 126 tablas con descripción detallada.

#### Tablas transaccionales (alta escritura)

| Tabla | Descripción breve | Clave primaria |
|---|---|---|
| `MPEDIDO` | Cabecera de pedido | `tCodigoPedido` |
| `DPEDIDO` | Detalle de pedido (ítems) | `tCodigoPedido + tItem` |
| `CPEDIDO` | Combos en pedido | `tCodigoPedido + tItem + tItemCombo` |
| `APEDIDO` | Ítems eliminados de pedido (auditoría) | `tCodigoPedido + tItem` |
| `MDOCUMENTO` | Documento de venta emitido | `tDocumento` |
| `DDOCUMENTO` | Detalle de documento | `tDocumento + tItem` |
| `DPAGODOCUMENTO` | Pagos de documento | `tDocumento + tCorrelativo` |
| `MTURNO` | Turno de caja | `tTurno` |
| `MEGRESO` | Egreso de caja | `tRecibo` |
| `MINGRESO` | Ingreso de caja | `tRecibo` |
| `MNOTACREDITO` | Nota de crédito | Definida en PK.sql |
| `DNOTACREDITO` | Detalle nota de crédito | Definida en PK.sql |

#### Tablas maestras (baja escritura)

| Tabla | Descripción breve | Clave primaria |
|---|---|---|
| `TPRODUCTO` | Catálogo de productos | `tCodigoProducto` |
| `TGRUPO` | Grupos de productos | `tCodigoGrupo` |
| `TSUBGRUPO` | Subgrupos de productos | `tCodigoSubGrupo` |
| `TCLIENTE` | Clientes | `tCodigoCliente` |
| `TUSUARIO` | Usuarios del sistema | — |
| `TCAJA` | Configuración de cajas | `tCaja` |
| `TMESA` | Mesas | — |
| `TPARAMETRO` | Parámetros globales (165+ columnas) | Tabla única |
| `TDIACONTABLE` | Día contable activo | — |
| `MCIERRE` | Control de cierre por período | `tPeriodo` |

### Estructura de MPEDIDO (tabla núcleo)

| Columna | Tipo | Descripción |
|---|---|---|
| `tCodigoPedido` | varchar(10) NOT NULL | Código único del pedido (PK) |
| `nCorrelativo` | int | Número correlativo del pedido en el turno |
| `fFecha` | smalldatetime | Fecha/hora de creación del pedido |
| `tEstadoPedido` | varchar(2) | Estado del pedido (abierto, pagado, anulado, etc.) |
| `tTipoPedido` | varchar(2) | Canal: 01=Local, 02=Delivery, 03=Llevar, 04=Canal4, 05=Canal5 |
| `tTipoAtencion` | varchar(2) | Tipo de atención |
| `tMesa` | varchar(3) | Mesa asignada |
| `nMesa` | smallint | Número de mesa |
| `tMozo` | varchar(4) | Código del mozo |
| `tMotorizado` | varchar(4) | Código del motorizado (delivery) |
| `tCaja` | varchar(3) | Caja que procesa el pedido |
| `tSalon` | varchar(2) | Salón |
| `tTurno` | varchar(10) | Turno activo |
| `nAdulto` | smallint | Pax adultos |
| `nNino` | smallint | Pax niños |
| `nDescuento` | float | Descuento aplicado al pedido |
| `tDescuento` | varchar(3) | Motivo de descuento |
| `tMotivoAnulacion` | varchar(3) | Motivo de anulación |
| `tHabitacion` | varchar(6) | Habitación (integración hotel) |
| `tReserva` | varchar(6) | Reserva asociada |
| `tPasajero` | varchar(50) | Nombre del pasajero (hotel) |
| `tCompania` | varchar(5) | Compañía (hotel) |
| `fDiaContable` | smalldatetime | Día contable en que cae el pedido |
| `nPrecuenta` | int | Contador de impresiones de precuenta |
| `nReimpresion` | int | Contador de reimpresiones |
| `lEntregado` | bit | Si el pedido fue entregado (delivery) |
| `fAsignacion` / `fSalida` / `fEntrega` / `fLlegada` | smalldatetime | Tiempos de delivery |

### Estructura de TPRODUCTO (tabla maestra clave)

| Columna | Tipo | Descripción |
|---|---|---|
| `tCodigoProducto` | varchar(7) NOT NULL | Código único (PK) |
| `tGrupo` | varchar(2) | Grupo de clasificación |
| `tSubGrupo` | varchar(4) | Subgrupo de clasificación |
| `tTipoProducto` | varchar(2) | Tipo de producto |
| `tDetallado` | varchar(50) | Nombre completo |
| `tResumido` | varchar(24) | Nombre corto (para botones y comandas) |
| `lImpuesto1..15` | bit × 15 | Flags de aplicación de impuesto por producto |
| `nPrecioVenta` | float | Precio local |
| `nPrecioDelivery` | float | Precio para delivery |
| `nPrecioLlevar` | float | Precio para llevar |
| `nPrecioCanal4` | float | Precio canal 4 |
| `nPrecioCanal5` | float | Precio canal 5 |
| `tArea` | varchar(3) | Área de cocina/producción |
| `lImprimeArea` | bit | Si imprime en área de cocina |
| `lActivo` | bit | Si el producto está activo |
| `lCombinacion` | bit | Si es producto combinado (combo) |
| `nBoton` | int | Posición en panel de botones del POS |
| `nBotonRapido` | int | Posición en caja rápida |
| `tEnlace` | varchar(7) | Producto del menú almacén enlazado |
| `oFoto` | image | Imagen del producto |
| `tDescargo` | varchar(1) | Tipo de descargo de insumos |

### Estructura de MTURNO (cierre de caja)

`MTURNO` consolida los totales de un turno de caja con 40+ columnas de montos por tipo de pago:

| Grupo de columnas | Descripción |
|---|---|
| `nMontoIN/IE` | Montos en efectivo nacional/extranjero |
| `nMontoEN/EE` | Montos en cheques nacional/extranjero |
| `nMontoCN/CE` | Montos en pagaré nacional/extranjero |
| `nMontoPN/PE` | Montos en pago a cuenta |
| `nMontoFN/FE` | Montos en firma nacional/extranjero |
| `nTarjeta1..8` | Montos por cada tipo de tarjeta (8 tipos) |
| `nPropina1..8` | Propinas por cada tipo de tarjeta |
| `nOtroN1..5` | Otros tipos de pago (moneda nacional) |

---

## Base de datos INFSEGURIDAD

Creada por `opcionales/Seguridad.sql`. Contiene el subsistema de licencias y auditoría.

### Tablas

| Tabla | Descripción |
|---|---|
| `THARDKEY` | Licencias de hardware: `id`, `tLicencia` (1000 chars), `fRegistro`, `tCliente` |
| `TPARAMETRO` | Parámetros de seguridad global: RUC, razón social, configuración de módulos, versión |
| `MMODULO` | Módulos del sistema: `tModulo` (PK), `tDetallado`, `tResumido` |
| `MMOVIMIENTO` | Auditoría a nivel de campo: tabla, campo, valor anterior/actual, usuario, fecha, acción |
| `MMOVIMIENTOACCESO` | Auditoría de acceso: login/logout de usuarios con timestamp |

### Modelo de auditoría (MMOVIMIENTO)

```
MMOVIMIENTO:
  PK: nCorrelativo + nCorrelativoCampo
  - tModulo: módulo origen
  - tBaseDatos: BD donde ocurrió el cambio
  - tTabla: tabla modificada
  - tAccion: INSERT/UPDATE/DELETE
  - ffecha: timestamp
  - tUsuario: usuario que hizo el cambio
  - tId_RegistroUnico: PK del registro modificado
  - tId_RegistroDetalle: PK del detalle (si aplica)
  - tCampo: nombre del campo modificado
  - tValorAnterior: valor previo
  - tValorActual: valor nuevo
```

---

## Procedimientos almacenados relevantes (detalle)

### `spIns_MPEDIDO`
Inserta la cabecera de un nuevo pedido en `MPEDIDO`. Recibe todos los campos de cabecera (tipo, mesa, mozo, turno, canal, etc.) y retorna el `tCodigoPedido` generado.

### `spUpd_MPEDIDO`
Actualiza el estado y campos de un pedido existente (estado, mesa, descuento, empacador, motorizado, etc.).

### `usp_Inforest_Impresion`
Obtiene el dataset completo para impresión de un documento, consolidando datos de `MDOCUMENTO`, `DDOCUMENTO`, `TPRODUCTO`, `TCLIENTE`, `TCAJA`, tipo de documento, tipo de cambio, propiedades de ítems, etc.

### `sp_AsignaAnticipo_Pedido`
Vincula un recibo de ingreso marcado como anticipo a un pedido específico; actualiza `MINGRESO` y `MPEDIDO`.

### `fn_cortesia_calculo`
Función escalar que calcula el monto de cortesía a aplicar sobre un documento, según las reglas configuradas en `TMOTIVODESCUENTO`.

### `CreaTabla`
Genera el nombre de una tabla temporal con sufijo de período (`CHAR(3)`, `DATETIME`) para reportes que crean tablas temporales de trabajo.

### `SP_EJECUTA_ACTUALIZA_FE`
Dispara el proceso de actualización de estado de Facturación Electrónica para documentos pendientes. Llama al motor de servicios del proveedor FE configurado.

### `usp_ControlServicioWindows`
Permite iniciar o detener servicios de Windows desde SQL Server (usado por el proceso de descargo automático en `clsAlmacen.EjecutarDescargoAutomatico`). Parámetros: `@NombreServicio`, `@Accion` (START/STOP), `@Resultado` (OUTPUT).

### `BK_INFOREST`
Ejecuta un backup de la base de datos INFOREST. Se invoca desde la interfaz de administración.

### `sp_OptimizarBD`
Reorganiza índices y realiza operaciones de mantenimiento de la BD. Se registra en `LOG_OPTIMIZACION`.

---

## Funciones escalares

### `fn_cortesia_calculo(@tDocumento, ...)`
Calcula el importe de cortesía para un documento. Retorna el monto calculado según reglas de `TMOTIVODESCUENTO` y el total del documento.

### `CreaTabla(@Flag CHAR(3), @Date DATETIME)`
Retorna un nombre de tabla temporal con sufijo de período para uso en reportes dinámicos.

---

## Vistas clave (descripción funcional)

| Vista | Descripción funcional |
|---|---|
| `vDocumento` | Vista principal de documentos de venta; consolida `MDOCUMENTO` + tipo de documento + cliente + caja + turno |
| `vDocumentoImpresora` | Datos de documento preparados para impresión (incluye impresora asignada, series) |
| `vPedido` | Vista de cabecera de pedido con estado, mesa, mozo, tipo |
| `vPedidoDetalle` | Vista de ítems de pedido con nombre de producto, precios, estado |
| `vPreCuenta` | Vista para impresión de precuenta (antes de cobrar) |
| `vPreCuentaDelivery` | Variante de precuenta para pedidos delivery |
| `vLiquidacion` | Datos de liquidación de turno por forma de pago |
| `vAreaImpresora` | Relación área de cocina → impresora → subgrupos de productos |
| `vDelivery` | Vista de clientes y pedidos delivery |
| `vMotorizado` | Datos de motorizados activos |
| `vProducto` | Catálogo de productos activos con precios por canal |
| `vEstadoPedido` | Catálogo de estados de pedido |
| `vEstadoDocumento` | Catálogo de estados de documento |
| `vDocumentoRegistroVentas` | Vista para registro de ventas formato tributario (SUNAT) |
| `vDocumentoRegistroVentas_TransGratuita` | Variante para transferencias gratuitas |
| `vNotaCredito` | Notas de crédito con datos del documento original |
| `vNotaCreditoImpresora` | Datos de NC para impresión |
| `vTipodocumentoImpresora` | Tipos de documento por caja e impresora |
| `vGrupoUsuario` | Grupos de usuarios y sus permisos |
| `vFormulario` | Formularios/módulos accesibles por el sistema |

---

## Consideraciones de migración

| Aspecto | Detalle |
|---|---|
| **Sin FK explícitas** | La integridad referencial está en la aplicación VB6 y los SPs. Al migrar, definir FK reales en PostgreSQL/destino |
| **Triggers eliminados** | El sistema elimina todos los triggers en cada arranque. Si se crean triggers en la BD legada, serán eliminados |
| **Collation** | `Modern_Spanish_CI_AS` — al migrar a PostgreSQL usar `es_ES.utf8` y ajustar comparaciones case-insensitive |
| **Tipos de datos** | Uso extensivo de `float` para montos monetarios — migrar a `DECIMAL(18,4)` en el destino |
| **smalldatetime** | Precisión de 1 minuto — migrar a `TIMESTAMP WITH TIME ZONE` |
| **varchar sin longitud máxima** | Algunas columnas de texto son muy largas (ej. `tEnlace` en TCLIENTE con 400 chars) — revisar longitudes al modelar |
| **Tablas de log** | `MMOVIMIENTO` puede ser muy grande — migrar como eventos inmutables en destino |
| **TPARAMETRO** | 165+ columnas en una sola fila — migrar a modelo de configuración clave-valor o JSON |
| **Imágenes** | `TPRODUCTO.oFoto` usa tipo `image` (deprecated en SQL Server) — migrar a almacenamiento de objetos (S3/GCS) |

---

## Anexo técnico adicional basado en scripts SQL (apéndice)

> Apéndice agregado sin remover contenido previo. Los scripts SQL leídos están codificados mayoritariamente en UTF-16. En esta revisión se detectaron **126 tablas**, **116 vistas**, **253 procedimientos** y **8 funciones** en los scripts fuente; esta cifra es mayor que la ya documentada en el README original, por lo que ambos conteos se preservan y este anexo describe el universo observado directamente en los archivos.

## A. Catálogo completo de tablas detectadas en `1. Estructura.sql`

| Tabla | Propósito inferido | Columnas/clave | Relaciones inferidas |
| --- | --- | --- | --- |
| APEDIDO | Detalle auxiliar de pedido/comanda por ítem. | tCodigoPedido, tItem; cols: tCodigoPedido, tItem, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo, nPrecioNeto | PK=tCodigoPedido,tItem, tCodigoPedido→MPEDIDO, tCodigoProducto→TPRODUCTO, tCodigoGrupo→TGRUPO |
| CPEDIDO | Detalle de combos/propiedades del pedido. | tCodigoPedido, tItem, tItemCombo; cols: tCodigoPedido, tItem, tItemCombo, tProducto, tProductoCombo, nCantidad | PK=tCodigoPedido,tItem,tItemCombo, tCodigoPedido→MPEDIDO, tProducto→TPRODUCTO, tProductoCombo→TPRODUCTO, tCodigoGrupo→TGRUPO |
| DDOCUMENTO | Detalle del documento emitido. | tDocumento, tItem, tCodigoPedido; cols: tDocumento, tItem, tCodigoPedido, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1 | PK=tDocumento,tItem,tCodigoPedido, tDocumento→MDOCUMENTO, tCodigoPedido→MPEDIDO, tCodigoProducto→TPRODUCTO |
| DGUIATRANSPORTE | Tabla detalle/transaccional del proceso indicado por el nombre. | tGuiaTransporte, tItem; cols: tGuiaTransporte, tItem, tCodigoProducto, nPrecioVenta, nCantidad, nVenta | PK=tGuiaTransporte,tItem, tGuiaTransporte→MGUIATRANSPORTE, tCodigoProducto→TPRODUCTO, tDocumento→MDOCUMENTO |
| DICTIONARY_INFOREST | Tabla detalle/transaccional del proceso indicado por el nombre. | idDictionary, clave, Peru; cols: idDictionary, clave, Peru, Bolivia, Ecuador, Argentina | No determinable con certeza desde el código |
| DINGRESO | Tabla detalle/transaccional del proceso indicado por el nombre. | tRecibo, nCorrelativo, monto; cols: tRecibo, nCorrelativo, monto, fRegistro, tUsuario, tCaja | tCaja→TCAJA, tTurno→MTURNO, tDocumento→MDOCUMENTO |
| DNOTACREDITO | Detalle de nota de crédito. | tNotaCredito, tItem; cols: tNotaCredito, tItem, tDocumento, tCodigoProducto, nPrecioNeto, nPrecioImpuesto1 | PK=tNotaCredito,tItem, tNotaCredito→MNOTACREDITO, tDocumento→MDOCUMENTO, tCodigoProducto→TPRODUCTO |
| DPAGODOCUMENTO | Pagos aplicados al documento. | tDocumento, tCorrelativo; cols: tDocumento, tCorrelativo, tTurno, tTipoPago, tOtroTipoPago, tMoneda | PK=tDocumento,tCorrelativo, tDocumento→MDOCUMENTO, tCorrelativo→TLOG, tTurno→MTURNO |
| DPAGODOCUMENTO_VC | Tabla detalle/transaccional del proceso indicado por el nombre. | tDocumento, tCorrelativo, tCod_Vale; cols: tDocumento, tCorrelativo, tCod_Vale, nMonto_Vale | tDocumento→MDOCUMENTO, tCorrelativo→TLOG |
| DPAGOTARJETA | Tabla detalle/transaccional del proceso indicado por el nombre. | tDocumento, tReferencia; cols: tDocumento, tReferencia, tTarjeta, tNumero, tFechaVencimiento, nMonto | PK=tDocumento,tReferencia, tDocumento→MDOCUMENTO |
| DPEDIDO | Detalle de productos del pedido. | tCodigoPedido, tItem; cols: tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo | PK=tCodigoPedido,tItem, tCodigoPedido→MPEDIDO, tCodigoProducto→TPRODUCTO, tCodigoGrupo→TGRUPO, tDocumento→MDOCUMENTO |
| DPEDIDOKDS | Tabla detalle/transaccional del proceso indicado por el nombre. | tCodigoPedido, tItem, fSalida; cols: tCodigoPedido, tItem, fSalida | tCodigoPedido→MPEDIDO |
| DPREPAGO | Prepagos asociados a documento/pedido. | tDocumento, tCorrelativo; cols: tDocumento, tCorrelativo, tTurno, tTipoPago, tOtroTipoPago, tMoneda | PK=tDocumento,tCorrelativo, tDocumento→MDOCUMENTO, tCorrelativo→TLOG, tTurno→MTURNO, tCodigoPedido→MPEDIDO |
| HISTORIAL_NOTICIAS | No determinable con certeza desde el código; el nombre sugiere una estructura de apoyo. | cod_not, tCodigoUsuario, FECHA_REGISTRO; cols: cod_not, tCodigoUsuario, FECHA_REGISTRO | tCodigoUsuario→TUSUARIO |
| INFOVISOR | No determinable con certeza desde el código; el nombre sugiere una estructura de apoyo. | ID, Pedido, tCaja; cols: ID, Pedido, tCaja, TipoDocu, RucFactura, RazonFactura | ID→LOG_INFOREST, tCaja→TCAJA |
| INTEGRACION_CASHDRO | No determinable con certeza desde el código; el nombre sugiere una estructura de apoyo. | Documento, Id, Monto; cols: Documento, Id, Monto, Estado, Json, FechaRegistro | Id→TDESCARGOINSUMO |
| LOG_INFOREST | Bitácora técnica y/o log transversal. | ID; cols: Ttabla, Proceso, CodError, ErrorProcedure, ErrorLine, ErrorMensaje | PK=ID |
| LOG_OPTIMIZACION | No determinable con certeza desde el código; el nombre sugiere una estructura de apoyo. | Usuario, fregistro, BD; cols: Usuario, fregistro, BD | No determinable con certeza desde el código |
| LOG_PEDIDO_DOCUMENTO | No determinable con certeza desde el código; el nombre sugiere una estructura de apoyo. | tCodigoPedido, Campo, ValorAnterior; cols: tCodigoPedido, Campo, ValorAnterior, ValorActual, Usuario, Fecha | tCodigoPedido→MPEDIDO |
| LOG_SESIONES | No determinable con certeza desde el código; el nombre sugiere una estructura de apoyo. | tCaja, fRegistro, lActivo; cols: tCaja, fRegistro, lActivo, tUsuario | tCaja→TCAJA |
| MCIERRE | Tabla maestra/cabecera del proceso indicado por el nombre. | tPeriodo; cols: tPeriodo, lCierre | PK=tPeriodo |
| MDOCUMENTO | Cabecera del comprobante/documento emitido. | tDocumento; cols: tDocumento, tTipoDocumento, tCodigoCliente, tCortesia, nNeto, nPrecioImpuesto1 | PK=tDocumento, tTipoDocumento→TTIPODOCUMENTO, tCodigoCliente→TCLIENTE, tTurno→MTURNO, tCaja→TCAJA |
| MEGRESO | Movimiento de egreso. | tRecibo; cols: tRecibo, tCaja, tTurno, fFecha, tMoneda, nTipoCambio | PK=tRecibo, tCaja→TCAJA, tTurno→MTURNO, fFecha→TTIPOCAMBIO |
| MENVIO | Tabla maestra/cabecera del proceso indicado por el nombre. | fInicio; cols: fInicio, fFinal, tUsuario, nMonto, lCopia, lCierre | PK=fInicio |
| MGUIATRANSPORTE | Tabla maestra/cabecera del proceso indicado por el nombre. | tGuiaTransporte; cols: tGuiaTransporte, fFecha, tCodigoDelivery, tDestinatario, tTienda, tTransportista | PK=tGuiaTransporte, fFecha→TTIPOCAMBIO, tCodigoDelivery→TDELIVERY, tDocumento→MDOCUMENTO, tCaja→TCAJA |
| MINGRESO | Movimiento de ingreso/cobro de caja. | tRecibo; cols: tRecibo, fFecha, tMoneda, tTipoPago, tTarjeta, tReferencia | PK=tRecibo, fFecha→TTIPOCAMBIO, tTurno→MTURNO, tCaja→TCAJA |
| MNOTACREDITO | Cabecera de nota de crédito. | tNotaCredito; cols: tNotaCredito, fFecha, tDocumento, nNeto, nImpuesto1, nImpuesto2 | PK=tNotaCredito, fFecha→TTIPOCAMBIO, tDocumento→MDOCUMENTO, tTurno→MTURNO, tCaja→TCAJA |
| MPEDIDO | Cabecera del pedido (mesa, delivery o llevar). | tCodigoPedido; cols: tCodigoPedido, nCorrelativo, tClienteDelivery, tClienteCtaCte, fFecha, tMoneda | PK=tCodigoPedido, fFecha→TTIPOCAMBIO, tMesa→TMESA, tCaja→TCAJA, tTurno→MTURNO |
| MPROPINA | Tabla maestra/cabecera del proceso indicado por el nombre. | tcodigopedido; cols: tCodigopedido, fRegistro, tMoneda, nMonto, tUsuario, tComanda | PK=tcodigopedido |
| MTURNO | Tabla maestra/cabecera del proceso indicado por el nombre. | tTurno; cols: tTurno, tCaja, tSalon, fInicial, fFinal, tUsuario | PK=tTurno, tCaja→TCAJA |
| NOTICIAS | No determinable con certeza desde el código; el nombre sugiere una estructura de apoyo. | ID_NOT, APLICACION, URI_IMG1; cols: ID_NOT, APLICACION, URI_IMG1, URI_IMG2, URI_IMG3, FECHA_REGISTRO | No determinable con certeza desde el código |
| PEDIDO_PAGOEFECTIVO | No determinable con certeza desde el código; el nombre sugiere una estructura de apoyo. | tCodigoPedido, lEstado; cols: tCodigoPedido, lEstado | tCodigoPedido→MPEDIDO |
| TACCESO | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoAcceso; cols: tCodigoAcceso, tModulo, tDescripcion, tFormulario, tTipoObjeto, tNombreObjeto | PK=tCodigoAcceso |
| TACCESOENVIA | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoAcceso; cols: tCodigoAcceso, lEnvia | PK=tCodigoAcceso, tCodigoAcceso→TACCESO |
| TAREACHEF | Tabla catálogo/configuración del proceso indicado por el nombre. | tCaja, tArea, lArea; cols: tCaja, tArea, lArea, tUsuario, fRegistro | tCaja→TCAJA |
| TAREAIMPRESORA | Mapa área ↔ impresora. | tCaja, tArea; cols: tCaja, tArea, tImpresora, tUsuario, fRegistro | PK=tCaja,tArea, tCaja→TCAJA, tImpresora→TIMPRESORA |
| TAREAPANTALLA | Tabla catálogo/configuración del proceso indicado por el nombre. | tArea, nColumna, lMuestra; cols: tArea, nColumna, lMuestra, nOrden, nAncho, tUsuario | No determinable con certeza desde el código |
| TAREAPANTALLA1 | Tabla catálogo/configuración del proceso indicado por el nombre. | tArea, nColumna, lMuestra; cols: tArea, nColumna, lMuestra, nOrden, nAncho, tUsuario | No determinable con certeza desde el código |
| TAREAPANTALLADESPACHO | Tabla catálogo/configuración del proceso indicado por el nombre. | tArea, nColumna, lMuestra; cols: tArea, nColumna, lMuestra, nOrden, nAncho, tUsuario | No determinable con certeza desde el código |
| TAREASUBGRUPO | Tabla catálogo/configuración del proceso indicado por el nombre. | tCaja, tSubgrupo; cols: tCaja, tSubGrupo, tArea, tUsuario, fRegistro | PK=tCaja,tSubgrupo, tCaja→TCAJA, tSubGrupo→TSUBGRUPO |
| TCAJA | Configuración de caja, serie y flags de emisión. | tCaja; cols: tCaja, tDescripcion, tPrecuenta, lComanda, vComanda, lMotivoEliminaC | PK=tCaja |
| TCAJACANALVENTA | Tabla catálogo/configuración del proceso indicado por el nombre. | tCaja, tCanalventa1, tCanalventa2; cols: tCaja, tCanalventa1, tCanalventa2, tCanalventa3, tCanalventa4, tCanalventa5 | tCaja→TCAJA |
| TCAJAORIGEN_BLOQUEO | Tabla catálogo/configuración del proceso indicado por el nombre. | tCaja, CodigoOrigenVenta; cols: tCaja, CodigoOrigenVenta | PK=tCaja,CodigoOrigenVenta, tCaja→TCAJA |
| TCANALVENTA | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoCanalVenta; cols: tCodigoCanalVenta, tDetallado, lActivaMozo, lActivaMotorizado, lObligaMesa, lObligaPax | PK=tCodigoCanalVenta |
| TCENTROCOSTO | Tabla catálogo/configuración del proceso indicado por el nombre. | TCodigo; cols: TCodigo, CentroCosto, fRegistro, lActivo | PK=TCodigo |
| TCLASESUNAT | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoClase; cols: tCodigoClase, tCodigoFamilia, Descripcion | PK=tCodigoClase, tCodigoFamilia→TFAMILIASUNAT |
| TCLIENTE | Maestro de clientes corporativos/facturación. | tCodigoCliente; cols: tCodigoCliente, tEmpresa, tIdentidad, tDireccion, lActivo, tUsuario | PK=tCodigoCliente |
| TCLIENTEPRODUCTO | Tabla catálogo/configuración del proceso indicado por el nombre. | tcodigoDelivery, tCodigoProducto, nPrecio; cols: tcodigoDelivery, tCodigoProducto, nPrecio, lPermiteDescuentos, tUsuario, fRegistro | tCodigoProducto→TPRODUCTO |
| TCOMBO | Tabla catálogo/configuración del proceso indicado por el nombre. | tCombo, tCodigoProducto; cols: tCombo, tCodigoProducto, nCantidad, lFijo, lUnico, tEtiqueta | PK=tCombo,tCodigoProducto, tCodigoProducto→TPRODUCTO |
| TCOMBOPROPIEDAD | Propiedades de ítems combo. | tCodigoPedido, tItem, tItemCombo, tCodigoPropiedad, tProducto, tEnlace; cols: tCodigoPedido, tItem, tItemCombo, tCodigoPropiedad, tProducto, tEnlace | PK=tCodigoPedido,tItem,tItemCombo,tCodigoPropiedad,tProducto,tEnlace, tCodigoPedido→MPEDIDO, tProducto→TPRODUCTO |
| TCOMPANIA | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoCliente; cols: tCodigoCliente, lEmpresa, tApeCom, tNomSoc, tTipoIdentidad, tIdentidad | PK=tCodigoCliente, tCodigoCliente→TCLIENTE |
| TCONFIGURAPERIFERICO | Tabla catálogo/configuración del proceso indicado por el nombre. | tTabla, tcaja, nDato1; cols: tTabla, tcaja, nDato1, nDato2, nDato3, nDato4 | No determinable con certeza desde el código |
| TDELIVERY | Clientes/condiciones de delivery y crédito. | tCodigoDelivery; cols: tCodigoDelivery, tTipoCliente, tApellido, tNombre, tDireccion, tTelefono | PK=tCodigoDelivery, tCodigoCliente→TCLIENTE, tCodigoTarjeta→TTARJETACREDITO |
| TDELIVERYCLIENTE | Tabla catálogo/configuración del proceso indicado por el nombre. | nCorrelativo; cols: nCorrelativo, tCodigoDelivery, tCodigoCliente | PK=nCorrelativo, tCodigoDelivery→TDELIVERY, tCodigoCliente→TCLIENTE |
| TDELIVERYINVITADO | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoInvitado; cols: tCodigoInvitado, tCodigoDelivery, tNombre, tApellido, fRegistro, tUsuario | PK=tCodigoInvitado, tCodigoDelivery→TDELIVERY |
| TDESCARGOINSUMO | Tabla catálogo/configuración del proceso indicado por el nombre. | Id; cols: Id, tCodigoPedido, Fecha, PlatoVenta, CantidadPlato, Item | PK=Id, tCodigoPedido→MPEDIDO, tDocumento→MDOCUMENTO |
| TDETALLEASISTENCIA | Tabla catálogo/configuración del proceso indicado por el nombre. | idmarca, tCodigo, tTipoUsuario; cols: idmarca, tCodigo, tTipoUsuario, fRegistro, fRegistroContable, tHora | No determinable con certeza desde el código |
| TDIACONTABLE | Tabla catálogo/configuración del proceso indicado por el nombre. | fDiaContable, lApertura, lCierre; cols: fDiaContable, lApertura, lCierre, tUsuario, fregistro, tUsuarioCierre | No determinable con certeza desde el código |
| TDISPENSADOR | Tabla catálogo/configuración del proceso indicado por el nombre. | Id, Descripcion, fRegistro; cols: Id, Descripcion, fRegistro, UsuarioRegistro, FechaModificion, UsuarioModificion | Id→TDESCARGOINSUMO |
| TESTADOBIZLINK | Tabla catálogo/configuración del proceso indicado por el nombre. | tDocumento, fRegistro, EstadoDocumentoInforest; cols: tDocumento, fRegistro, EstadoDocumentoInforest, DocSunat, EstadoBizlink, CodigoEstadoBizlink | tDocumento→MDOCUMENTO |
| TESTADOINFOFACT | Tabla catálogo/configuración del proceso indicado por el nombre. | tDocumento, fRegistro, DocSunat; cols: tDocumento, fRegistro, DocSunat, EstadoInfoFact, NroTicketCdr, ZipCdr | tDocumento→MDOCUMENTO |
| TFAMILIASUNAT | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoFamilia; cols: tCodigoFamilia, tCodigoSegmento, Descripcion | PK=tCodigoFamilia, tCodigoSegmento→TSEGMENTOSUNAT |
| TGRUPO | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoGrupo; cols: tCodigoGrupo, tDetallado, tResumido, nBoton, lActivo, tUsuario | PK=tCodigoGrupo, tCaja→TCAJA |
| TGRUPOACCESO | Tabla catálogo/configuración del proceso indicado por el nombre. | tGrupoUsuario, tCodigoAcceso; cols: tGrupoUsuario, tCodigoAcceso | PK=tGrupoUsuario,tCodigoAcceso, tGrupoUsuario→TGRUPOUSUARIO, tCodigoAcceso→TACCESO |
| TGRUPOUSUARIO | Tabla catálogo/configuración del proceso indicado por el nombre. | tGrupoUsuario; cols: tGrupoUsuario, tDetallado, lModulo01, lModulo02, lModulo03, lOpcion01 | PK=tGrupoUsuario |
| TIMPORTACION | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigo, nOrden; cols: tCodigo, nOrden, tPadre, tAgrupacion, tTablaMostrar, tTablaInterna | PK=tCodigo,nOrden |
| TIMPORTACIONLOG | Tabla catálogo/configuración del proceso indicado por el nombre. | tModulo, fRegistro, tUsuario; cols: tModulo, fRegistro, tUsuario, tEstado, tObservacion | No determinable con certeza desde el código |
| TIMPRESORA | Catálogo de impresoras por caja. | tCaja, tImpresora; cols: tCaja, tImpresora, tDescripcion, tDevice, tRuta, tFont | PK=tCaja,tImpresora, tCaja→TCAJA, tImpresora→TIMPRESORAIMPRESION |
| TIMPRESORAIMPRESION | Tabla catálogo/configuración del proceso indicado por el nombre. | tImpresora; cols: tImpresora, tf1, tf2, tf3, tf4, tf5 | PK=tImpresora, tImpresora→TIMPRESORA |
| TINSUMO | Catálogo de insumos/almacén. | TCODIGO; cols: tcodigo, Descripcion, nStock, tUsuarioReg, fregistro, tcajaReg | PK=TCODIGO |
| TINTEGRACIONES | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigo; cols: tCodigo, Descripcion, tResumido, tValor, lActivo | PK=tCodigo |
| TLISTAESPERA | Tabla catálogo/configuración del proceso indicado por el nombre. | tListaespera; cols: tListaespera, fFecha, fHora, tApellido, tNombre, tTelefono | PK=tListaespera, fFecha→TTIPOCAMBIO, tMesa→TMESA |
| TLOCAL | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoLocal; cols: tCodigoLocal, tDetallado, tResumido, tcodigoSector, tIP, tBaseDatosINF | PK=tCodigoLocal |
| TLOG | Tabla catálogo/configuración del proceso indicado por el nombre. | tCorrelativo; cols: tCorrelativo, nProducto, nPropiedad, nOferta, nMesa, nMozo | PK=tCorrelativo |
| TLOG_IMPRESION | Tabla catálogo/configuración del proceso indicado por el nombre. | tDocumento, tPosicion1, tPosicion2; cols: tDocumento, tPosicion1, tPosicion2, tPosicion3, tPosicion4, tPosicion5 | tDocumento→MDOCUMENTO |
| TLOG_MODPRECIO | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoPedido, tItem, nDescuentoAnt; cols: tCodigoPedido, tItem, nDescuentoAnt, nDescuento, nRecargoAnt, nRecargo | tCodigoPedido→MPEDIDO |
| TMENSAJE | Tabla catálogo/configuración del proceso indicado por el nombre. | tF1, tF2, tF3; cols: tF1, tF2, tF3, tF4, tF5, tF6 | No determinable con certeza desde el código |
| TMENSAJECOCINA | Tabla catálogo/configuración del proceso indicado por el nombre. | codigo; cols: Codigo, tUsuarioReg, Mensaje, fregistro, fFinal, tUsuarioFinal | PK=codigo, tCaja→TCAJA |
| TMENSAJEIMPRESORA | Tabla catálogo/configuración del proceso indicado por el nombre. | teclaFuncion, tCaja, tImpresora; cols: teclaFuncion, tCaja, tImpresora, tUsuarioReg, fRegistro | tCaja→TCAJA, tImpresora→TIMPRESORA |
| TMENSAJEUSUARIO | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoPedido, tUsuario, tMensaje; cols: tCodigoPedido, tUsuario, tMensaje, tCaja, tMozo, tMesa | tCodigoPedido→MPEDIDO, tCaja→TCAJA, tMesa→TMESA |
| TMESA | Catálogo y estado de mesas. | tCodigoMesa; cols: tCodigoMesa, tDetallado, tResumido, tSalon, lFumador, tX | PK=tCodigoMesa |
| TMODULO | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoModulo, tSecuencia; cols: tCodigoModulo, tSecuencia, tFormulario, tMenu1, tMenu2, tMenu3 | PK=tCodigoModulo,tSecuencia |
| TMOTIVODESCUENTO | Tabla catálogo/configuración del proceso indicado por el nombre. | tDescuento; cols: tDescuento, tDetallado, tResumido, nRatio, nTope, lRatio | PK=tDescuento |
| TMOTORIZADODATOS | Catálogo operativo de motorizados. | tCodigo; cols: tCodigo, tDocumentoIdentidad, nTarifaLV, nTarifaSD, nTarifaES | PK=tCodigo |
| TMOVIMIENTOTARJETASRFID | Tabla catálogo/configuración del proceso indicado por el nombre. | Id, CodidoRFID, Tipo; cols: Id, CodidoRFID, Tipo, DocReferencia, CodReferencia, CodDispensador | Id→TDESCARGOINSUMO |
| TOFERTA | Tabla catálogo/configuración del proceso indicado por el nombre. | tOferta, tCodigoProducto; cols: tOferta, tCodigoProducto, tNombre, tResumido, tFrecuencia, fFecha | PK=tOferta,tCodigoProducto, tCodigoProducto→TPRODUCTO, fFecha→TTIPOCAMBIO |
| TOPERADOR | Tabla catálogo/configuración del proceso indicado por el nombre. | tOperador; cols: tOperador, tDetallado, tResumido, lValor, lStockMas, lStockMenos | PK=tOperador |
| TORIGENCODIGOCONTROL | Tabla catálogo/configuración del proceso indicado por el nombre. | tCaja, nCorrelativo; cols: tCaja, nCorrelativo, fInicio, fFin, tAutorizacion, tDosificacion | PK=tCaja,nCorrelativo, tCaja→TCAJA, fInicio→MENVIO |
| TORIGENVENTA | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoOrigenVenta; cols: tCodigoOrigenVenta, tCodigoCanalVenta, tDescripcion, nBoton, lVisiblePos, lDelivery | PK=tCodigoOrigenVenta, tCodigoCanalVenta→TCANALVENTA |
| TPARAMETRO | Configuración global del negocio/país/impuestos/impresión/FE. | tIdentificacionTributaria; cols: tIdentificacionTributaria, tRazonSocial, tRazonComercial, tDireccion, tTelefono, tEmail | PK=tIdentificacionTributaria, tCodigoUbigeo→TUBIGEO |
| TPARIENTE | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoPariente, tCodigoDelivery; cols: tCodigoPariente, tCodigoDelivery, tNombre, tApellido, lConyugue, lHijo | PK=tCodigoPariente,tCodigoDelivery, tCodigoDelivery→TDELIVERY |
| TPEDIDO | Tabla catálogo/configuración del proceso indicado por el nombre. | nCorrelativo; cols: nCorrelativo, tPedidoIni, tItemIni, tPedidoFin, tItemFin, tProducto | PK=nCorrelativo, tProducto→TPRODUCTO, tTurno→MTURNO |
| TPEDIDOMESA | Relación pedido ↔ mesa(s). | tCodigoPedido, tMesa; cols: tCodigoPedido, tMesa | PK=tCodigoPedido,tMesa, tCodigoPedido→MPEDIDO, tMesa→TMESA |
| TPRODUCTO | Catálogo de productos con precios/flags/impuestos por canal. | tCodigoProducto; cols: tCodigoProducto, tGrupo, tSubGrupo, tTipoProducto, tDetallado, tResumido | PK=tCodigoProducto, tSubGrupo→TSUBGRUPO |
| TPRODUCTOAREA | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoProducto, tArea; cols: tCodigoProducto, tArea | PK=tCodigoProducto,tArea, tCodigoProducto→TPRODUCTO |
| TPRODUCTOPROPIEDAD | Propiedades o modificadores válidos por producto. | tCodigoPedido, tItem, tCodigoPropiedad, tProducto, tEnlace; cols: tCodigoPedido, tItem, tCodigoPropiedad, tProducto, tEnlace, nInsumo | PK=tCodigoPedido,tItem,tCodigoPropiedad,tProducto,tEnlace, tCodigoPedido→MPEDIDO, tProducto→TPRODUCTO |
| TPRODUCTOSUNAT | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoProducto; cols: tCodigoProducto, tCodigoClase, Descripcion, DescripcionIngles | PK=tCodigoProducto, tCodigoProducto→TPRODUCTO, tCodigoClase→TCLASESUNAT |
| TPRODUCTOXPRODUCTO | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoProducto, tSubProducto; cols: tCodigoProducto, tSubProducto, nCantidad | PK=tCodigoProducto,tSubProducto, tCodigoProducto→TPRODUCTO |
| TPROGRAMAPRECIOS_CAB | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoProgPrecio, tCodigoProgPlantilla, tNomProgPrecio; cols: tCodigoProgPrecio, tCodigoProgPlantilla, tNomProgPrecio, Fecha, Hora, lActivo | No determinable con certeza desde el código |
| TPROGRAMAPRECIOS_DETA | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoProgPrecio, tCodigoProducto, nPrecioVenta1; cols: tCodigoProgPrecio, tCodigoProducto, nPrecioVenta1, nPrecioVentaN1, nPrecioVenta2, nPrecioVentaN2 | tCodigoProducto→TPRODUCTO |
| TPROPIEDAD | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoPropiedad, tProducto; cols: tCodigoPropiedad, tProducto, tDetallado, tResumido, tOperador, nPrecio | PK=tCodigoPropiedad,tProducto, tProducto→TPRODUCTO, tOperador→TOPERADOR |
| TRESERVA | Tabla catálogo/configuración del proceso indicado por el nombre. | tReserva; cols: tReserva, fFecha, fHora, tApellido, tNombre, tTelefono | PK=tReserva, fFecha→TTIPOCAMBIO, tMesa→TMESA, tCodigoDelivery→TDELIVERY |
| TSEGMENTOSUNAT | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoSegmento; cols: tCodigoSegmento, Descripcion | PK=tCodigoSegmento |
| TSOLICITUD | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoSolicitud; cols: tCodigoSolicitud, tCodigoTramite, tEstadoSolicitud, tNumeroAutorizacion, tNumeroAutorizacionAnterior, fFechaEmision | PK=tCodigoSolicitud, tCodigoTramite→TTRAMITE |
| TSOLICITUDDETALLE | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoSolicitud, nCorrelativo; cols: tCodigoSolicitud, nCorrelativo, tTipoDocumento, tSerieEstablecimiento, tSerieCaja, tCaja | PK=tCodigoSolicitud,nCorrelativo, tCodigoSolicitud→TSOLICITUD, tTipoDocumento→TTIPODOCUMENTO, tCaja→TCAJA |
| TSUBGRUPO | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoGrupo, tCodigoSubgrupo; cols: tCodigoGrupo, tCodigoSubgrupo, tDetallado, tResumido, nBoton, tIcono | PK=tCodigoGrupo,tCodigoSubgrupo, tCodigoGrupo→TGRUPO |
| TTABLA | Tabla catálogo/configuración del proceso indicado por el nombre. | TTABLA, TCODIGO; cols: TTABLA, TCODIGO, nTamano, tDetallado, tResumido, tIcono | PK=TTABLA,TCODIGO, TCODIGO→TINSUMO |
| TTARJETACREDITO | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoTarjeta; cols: tCodigoTarjeta, tDetallado, tResumido, nFactorRetencion, tRepresentante, tTelefono1 | PK=tCodigoTarjeta |
| TTARJETASRFID | Tabla catálogo/configuración del proceso indicado por el nombre. | Id, CodidoRFID, fRegistro; cols: Id, CodidoRFID, fRegistro, UsuarioRegistro, FechaModificion, UsuarioModificion | Id→TDESCARGOINSUMO |
| TTERMINAL | Tabla catálogo/configuración del proceso indicado por el nombre. | Codigo, tIntegracion, tComercio; cols: Codigo, tIntegracion, tComercio, tSucursal, tTerminal, tDescripcion | tCaja→TCAJA |
| TTIENDA | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoDelivery, tCodigoTienda; cols: tCodigoDelivery, tCodigoTienda, tNombre, tDireccion, tTelefono, tEmail | PK=tCodigoDelivery,tCodigoTienda, tCodigoDelivery→TDELIVERY |
| TTIPOCAMBIO | Tipo de cambio por fecha. | fFecha; cols: fFecha, nCompra, nVenta, tUsuario, fRegistro, lReplica | PK=fFecha |
| TTIPODOCUMENTO | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoTipoDocumento; cols: tCodigoTipoDocumento, tDescripcion, tPrefijo, tCodigoSunat, lPideCliente, nMonto | PK=tCodigoTipoDocumento |
| TTIPODOCUMENTOIMPRESORA | Tabla catálogo/configuración del proceso indicado por el nombre. | tCaja, tTipoEmision; cols: tCaja, tImpresora, tTipoEmision, tDescripcion, tFormulario, tSerie | PK=tCaja,tTipoEmision, tCaja→TCAJA, tImpresora→TIMPRESORA, fInicio→MENVIO |
| TTIPOIDENTIDAD | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigo; cols: tCodigo, Descripcion, tResumido, nValor, tValor, tValor2 | PK=tCodigo |
| TTIPOMOVIMIENTO | Tabla catálogo/configuración del proceso indicado por el nombre. | Codido, Description; cols: Codido, Description | No determinable con certeza desde el código |
| TTIPOPEDIDODETALLE | Tabla catálogo/configuración del proceso indicado por el nombre. | tcodigoTipoPedido, lActivaMozo, lActivaMotorizado; cols: tcodigoTipoPedido, lActivaMozo, lActivaMotorizado, lObligaMesa, lObligaPax, lObligaMozo | No determinable con certeza desde el código |
| TTRAMITE | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoTramite; cols: tCodigoTramite, tDescripcion, lSolicitaNAnteriorAutorizacion, lActivo | PK=tCodigoTramite |
| TUBIGEO | Ubigeo/distrito/provincia/departamento. | tCodigo, tDistrito, tProvincia; cols: tCodigo, tDistrito, tProvincia, tDepartamento | No determinable con certeza desde el código |
| TUSUARIO | Tabla catálogo/configuración del proceso indicado por el nombre. | tCodigoUsuario; cols: tCodigoUsuario, tGrupoUsuario, tDetallado, tResumido, tPassword, lActivo | PK=tCodigoUsuario, tGrupoUsuario→TGRUPOUSUARIO |
| TVISIBILIDADTARJETACREDITOXCANAL | Tabla catálogo/configuración del proceso indicado por el nombre. | lVisibilidad, lCanal1, lCanal2; cols: lVisibilidad, lCanal1, lCanal2, lCanal3, lCanal4, lCanal5 | tCodigoTarjeta→TTARJETACREDITO |
| VALE_CONSUMO | No determinable con certeza desde el código; el nombre sugiere una estructura de apoyo. | tTICKET, tCOD_VALE, tDOC_RELACIONADO; cols: tTICKET, tCOD_VALE, tDOC_RELACIONADO, tTIPODOC_CLIENTE, tDOC_CLIENTE, tRAZSOCIAL_CLIENTE | No determinable con certeza desde el código |
| VISIBILIDADPROPIEDADXCANAL | No determinable con certeza desde el código; el nombre sugiere una estructura de apoyo. | visibilidad, canal1, canal2; cols: visibilidad, canal1, canal2, canal3, canal4, canal5 | No determinable con certeza desde el código |
| VISOR_DPEDIDO | No determinable con certeza desde el código; el nombre sugiere una estructura de apoyo. | tCodigoPedido, tItem; cols: tCodigoPedido, tItem, tTipoPedido, tCodigoProducto, tCodigoGrupo, tCodigoSubGrupo | PK=tCodigoPedido,tItem, tCodigoPedido→MPEDIDO, tCodigoProducto→TPRODUCTO, tCodigoGrupo→TGRUPO, tDocumento→MDOCUMENTO |
| ruc_temp | No determinable con certeza desde el código; el nombre sugiere una estructura de apoyo. | xRUC, xNOMBRE, xUrbanizacion; cols: xRUC, xNOMBRE, xUrbanizacion, xDireccion, xCodigoUbigeo | No determinable con certeza desde el código |

## B. Catálogo completo de procedimientos almacenados y funciones detectadas en `5. SP.sql`

| SP | Propósito inferido | Parámetros | Tablas accedidas |
| --- | --- | --- | --- |
| BK_INFOREST | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @RUTA, @BD, @path, @fileName, @fileDate | MPEDIDO |
| CalcularStockOfertas | Gestiona stock, descargo o kardex. | @Producto, @oferta, @cantidad, @tipoPedido, @lPermanente, @lFechaEspecial, @Excluyente, @Fechaini, @Fechafin, @FechaEspecial, @HoraIni, @Horafin, @Lunes, @Martes, @Miercoles, @Jueves, @Viernes, @Sabado | TOFERTA, DPEDIDO |
| DiacontableHora | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @fecha | No determinable con certeza desde el código |
| SP_COVER_INSERTA_PEDIDO_ADICION | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @tCaja, @codigoReserva, @tMozo, @tPedido, @tTurno, @tUsuario, @nCorrela, @tCliente, @tTipoPedido, @lPrioridad, @tTipoAtencion, @tMesa | TRESERVA, MTURNO, TORIGENVENTA, TMESA, TPARAMETRO, MPEDIDO, TPEDIDOMESA |
| SP_DATOS_INFOREST | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @RUC | TPARAMETRO |
| SP_Dictionary_Inforest | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @clave, @pais | DICTIONARY_INFOREST |
| SP_EJECUTA_ACTUALIZA_FE | Actualiza registros del proceso indicado por el nombre. | @cmd, @TB_TEMP | No determinable con certeza desde el código |
| SP_GENERA_ETIQUETA | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @tTipo, @tGrupo, @tSubgrupo, @t | No determinable con certeza desde el código |
| SP_INFOREST_VALPRODASOCIADOS | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @xproducto, @CanalVenta, @fRegistro, @tCodigoDelivery, @tCodigoPedido, @xlRangoHorarios, @xtHoraInicio, @xtHoraFin, @xtCanal1, @xtCanal2, @xtCanal3, @xtCanal4, @xtCanal5, @xlSoloxDia, @fRegistr | TDELIVERY, DPEDIDO, MPEDIDO |
| SP_InsertaDescargoInsumo | Gestiona stock, descargo o kardex. | @Id, @tCodigoPedido, @Fecha, @PlatoVenta, @CantidadPlato, @Item, @tDescargo, @Enlace, @TipoPedido, @CodigoProducto, @CantidadReceta, @SubAreaAlm, @SubareaInf, @RecetaBase, @lProducto, @tCodigoUnicoEtiqueta, @tDocumento, @fDiaContable, @tTipoDescargoDes, @tCodigoRecetaRB, @tCodigoItemRB, @tUsuarioProceso, @Mensaje, @Correlativo | TDESCARGOINSUMO |
| SP_NOTACREDITO_PARCIAL | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @documento, @NotaCredito, @item, @tipooper, @PrecioUnitario, @Cantidad | DDOCUMENTO, MDOCUMENTO, DNOTACREDITO, MNOTACREDITO |
| SP_NOTICIAS_INFOREST | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @opcion, @usuario, @resultado | NOTICIAS, HISTORIAL_NOTICIAS |
| USP_ADD_MARCACION | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @CLAVE, @CODIGOUSUARIO, @cant, @REVISARTIPOMARCA, @FECHACONTABLE, @ACTIVO, @THORADECIERRA, @xlactivo, @RESULT, @xtTipoMarca, @xtipousuario, @xcodigotipousuario | TPARAMETRO, TDETALLEASISTENCIA, TUSUARIO, TTABLA |
| USP_AGREGARINSUMOS | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @CODIGO, @USUARIO, @DESCRIPCION, @NSTOCK, @TCAJA, @ACTIVO, @INSUMO | TINSUMO |
| USP_AGREGARMENSAJE | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @codigo, @usuario, @mensaje, @tcaja, @lactivo, @lactivaralerta | No determinable con certeza desde el código |
| USP_CALCULA_PRECIO | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @tCodigoPedido, @tItem, @tCodigoProducto, @tTipoPedido, @nCantidad, @decide, @nPrecioNeto, @nPrecioImpuesto1, @nPrecioImpuesto2, @nPrecioImpuesto3, @lImpuesto1, @lImpuesto2, @lImpuesto3, @Impuesto1, @Impuesto2 | TPROPIEDAD, DPEDIDO |
| USP_CERRAR_MENSAJES_CIERRETURNO | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @usuario, @tcaja | No determinable con certeza desde el código |
| USP_ELIMINARINSUMOS | Elimina o depura registros del proceso indicado por el nombre. | @CODIGO | TINSUMO |
| USP_ELIMINARRMENSAJES | Elimina o depura registros del proceso indicado por el nombre. | @tCODIGO, @TCODIGO | No determinable con certeza desde el código |
| USP_INF_verificaProductoAreaPantalla | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @tCodigoProducto | TPRODUCTOAREA |
| USP_KDS_GrabarPath | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @tOrderInfo, @tOrderStatus, @tBump, @tOrderInfox, @tOrderStatusx, @tBumpx, @lKDSInforest, @lKDSxArea, @lKDSTodos, @lKDS | TPARAMETRO |
| USP_KDS_GrabarTiempoSalidaDPedido | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @tCodigoPedido, @tItem, @fSalida, @count | No determinable con certeza desde el código |
| USP_KDS_ObtenerArea | Consulta y devuelve datos del proceso indicado por el nombre. | @Codigo | No determinable con certeza desde el código |
| USP_KDS_ObtenerAreaImpresionKDS | Consulta y devuelve datos del proceso indicado por el nombre. | @tCodigoProducto, @tItemCombo, @tCodigoPedido, @tItem | TPRODUCTOAREA |
| USP_KDS_ObtenerCategoria | Consulta y devuelve datos del proceso indicado por el nombre. | @tCodigoGrupo, @tCodigoSubGrupo | No determinable con certeza desde el código |
| USP_KDS_ObtenerDetalleCombo | Consulta y devuelve datos del proceso indicado por el nombre. | @tCodigoPedido, @tItem, @lImprime, @accion | No determinable con certeza desde el código |
| USP_KDS_ObtenerDetallePedido | Consulta y devuelve datos del proceso indicado por el nombre. | @tCodigoPedido, @accion, @MOBILE | MPEDIDO |
| USP_KDS_ObtenerDetallePedidox | Consulta y devuelve datos del proceso indicado por el nombre. | @tCodigoPedido, @accion, @kds, @MOBILE | MPEDIDO |
| USP_KDS_ObtenerDetallePedidoxKDS | Consulta y devuelve datos del proceso indicado por el nombre. | @tCodigoPedido, @ITEM | No determinable con certeza desde el código |
| USP_KDS_ObtenerNombreMesaXCodigo | Consulta y devuelve datos del proceso indicado por el nombre. | @tCodigoMesa | No determinable con certeza desde el código |
| USP_KDS_ObtenerPath | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | TPARAMETRO |
| USP_KDS_ObtenerProductoArea | Consulta y devuelve datos del proceso indicado por el nombre. | @tCodigoProducto, @tArea | TPRODUCTOAREA |
| USP_KDS_ObtenerProductoPedido | Consulta y devuelve datos del proceso indicado por el nombre. | @tCodigoPedido, @tItem, @integracion | TCAJA, MPEDIDO, DPEDIDO, TPRODUCTO, TPRODUCTOAREA, CPEDIDO |
| USP_KDS_ObtenerProductoPedidoDeCombo | Consulta y devuelve datos del proceso indicado por el nombre. | @tCodigoPedido, @tItem, @xItem | CPEDIDO, TPRODUCTO, TPRODUCTOAREA |
| USP_KDS_ObtenerProductoPedidoImpresos | Consulta y devuelve datos del proceso indicado por el nombre. | @tCodigoPedido, @tItem, @accion | DPEDIDO |
| USP_KDS_ObtenerPropiedadesProducto | Consulta y devuelve datos del proceso indicado por el nombre. | @tCodigoPedido, @tItem, @tItemCombo, @tProducto | No determinable con certeza desde el código |
| USP_KDS_ObtenerTipoPedido | Consulta y devuelve datos del proceso indicado por el nombre. | @tCodigo | TTABLA, TCANALVENTA |
| USP_KDS_ResporteTiempoPedido | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @fInicio, @fFinal | DPEDIDOKDS, DPEDIDO, TPRODUCTO |
| USP_KDS_ResporteTiempoProducto | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @fInicio, @fFinal, @tGrupo, @tSubGrupo, @tProducto | DPEDIDOKDS, DPEDIDO, TPRODUCTO, TSUBGRUPO, TGRUPO |
| USP_LISTADOMENSAJES | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | Sin parámetros explícitos detectados | No determinable con certeza desde el código |
| USP_LISTARINSUMOS | Consulta y devuelve datos del proceso indicado por el nombre. | @MODULO | TINSUMO |
| USP_LISTARMENSAJES | Consulta y devuelve datos del proceso indicado por el nombre. | @fechaini, @fechafin, @tcaja | No determinable con certeza desde el código |
| USP_MODIFICARINSUMOS | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @CODIGO, @USUARIO, @DESCRIPCION, @TCAJA, @ACTIVO, @NSTOCK, @MODULO, @LINSUMO | TINSUMO |
| USP_MODIFICARMENSAJE | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @usuario, @codigo, @mensaje, @tcaja, @lactivo, @lactivaralerta | No determinable con certeza desde el código |
| USP_RD_VALIDANDORUC | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | Sin parámetros explícitos detectados | TPARAMETRO |
| USP_RPT_DETA_COMBO | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @fInicio, @fFinal, @ttabla, @tcodigoProducto, @cant, @total_costo, @total_venta, @foodcost | MPEDIDO, TPRODUCTO, CPEDIDO, TCOMBO |
| USP_UpdateReplicado | Actualiza registros del proceso indicado por el nombre. | @WhereReplicado, @TablaReplicado, @Query | No determinable con certeza desde el código |
| USP_actualizaStockInsumo | Actualiza registros del proceso indicado por el nombre. | @vi_detalles, @vi_numdet, @vch_Salida, @codigoInsumo, @cantidad, @vl_contador, @vl_item, @stockactual, @vchMsgError | No determinable con certeza desde el código |
| Usp_ActualizarPreciosTransferenciaAlmacen | Actualiza registros del proceso indicado por el nombre. | @SubGrupo, @BaseDatoAlmacen, @tipooper, @sql, @comilla | TPRODUCTO, TSUBGRUPO |
| Usp_GuiaTransporteXml | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @tGuiaTransporte | TPRODUCTO, DPEDIDO |
| dbo | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | Sin parámetros explícitos detectados | MTURNO |
| spIns_CENTROCOSTO | Inserta registros del proceso indicado por el nombre. | @TCodigo, @tipo, @Definicion, @lactivo, @number, @val | TCENTROCOSTO |
| spIns_DPEDIDO | Inserta registros del proceso indicado por el nombre. | @tPedido, @tTipoPedido, @tItem, @tProducto, @tGrupo, @tSubGrupo, @nPBase, @nRecargo, @nDescuento, @nOficial, @nImpuesto1, @nImpuesto2, @nImpuesto3, @nPVenta, @nCantidad, @tArea, @lImprimeArea, @lCombinacion | DPEDIDO |
| spIns_MPEDIDO | Inserta registros del proceso indicado por el nombre. | @tCliente, @tTipoPedido, @lPrioridad, @tTipoAtencion, @tMesa, @tMozo, @tMotorizado, @tCaja, @tSalon, @tTurno, @tObservacion, @nTiempo, @tUsuario, @nAdulto, @nNino, @nMesa, @tPuntoVenta | TPARAMETRO, TMESA, MPEDIDO, TPEDIDOMESA |
| spIns_MPEDIDO_RESERVA | Inserta registros del proceso indicado por el nombre. | @tCaja, @tUsuario, @tTurno, @fDiaContable, @fProgramacion, @codigoReserva, @tPedido, @nCorrela, @xSalon, @tCliente, @tTipoPedido, @lPrioridad, @tTipoAtencion | TPARAMETRO, TMESA, TRESERVA, MPEDIDO, TPEDIDOMESA |
| spIns_TipoCambio | Inserta registros del proceso indicado por el nombre. | @nTC, @tUsuario, @nTCO, @nTC2, @nTC3 | TTIPOCAMBIO |
| spRep_AnaliticoMotorizado | Genera dataset de reporte/consulta operacional. | @flagProduccion, @flagVenta, @flagCortesia, @flagCuentaCte, @flagPedidosFacturados, @flagCombinacion, @flagCargo, @flagOpcion, @flagTurnoOFecha, @tTipoProducto, @tMotorizado, @tGrupo, @tSubGrupo, @tcodigoProduct | TCANALVENTA, MPEDIDO |
| spRep_AnaliticoMotorizadoIntegrado | Genera dataset de reporte/consulta operacional. | @flagProduccion, @flagVenta, @flagCortesia, @flagCuentaCte, @flagPedidosFacturados, @flagCombinacion, @flagCargo, @sPrecio, @finicio, @ffinal, @sql, @comilla, @sCriterio, @sCriterioP | MPEDIDO |
| spRep_AnaliticoMozo | Genera dataset de reporte/consulta operacional. | @flagFranjaHoraria, @flagProduccion, @flagVenta, @flagCortesia, @flagCuentaCte, @flagPedidosFacturados, @flagCombinacion, @flagCargo, @flagOpcion, @flagTurnoOFecha, @tTipoProducto, @tTipoPedido, @tMozo, @tGrupo | TUSUARIO |
| spRep_AnaliticoMozoIntegrado | Genera dataset de reporte/consulta operacional. | @flagProduccion, @flagVenta, @flagCortesia, @flagCuentaCte, @flagPedidosFacturados, @flagCombinacion, @flagCargo, @finicio, @ffinal, @sPrecio, @sql, @comilla, @sCriterio, @sCriterioM | No determinable con certeza desde el código |
| spRep_Anulacion | Genera dataset de reporte/consulta operacional. | @lFranjaHoraria, @tTurno, @fInicio, @fFinal, @lFlag1, @lFlag2, @lFlag3, @sCriterio, @ExportaExcel, @sql, @comilla | No determinable con certeza desde el código |
| spRep_AnulacionDocumentoIntegrado | Genera dataset de reporte/consulta operacional. | @flagFranja, @fInicio, @fFinal | MDOCUMENTO, DDOCUMENTO, MPEDIDO |
| spRep_AnulacionPedidoIntegrado | Genera dataset de reporte/consulta operacional. | @flagFranja, @fInicio, @fFinal, @flagNoImp, @flagImp, @flagTodos | MPEDIDO, APEDIDO, TMESA, DPEDIDO |
| spRep_Asistencia | Genera dataset de reporte/consulta operacional. | @fecha1, @fecha2, @ttipousuario, @CODIGO, @hi, @hf, @SqlDinamica | TDETALLEASISTENCIA, TUSUARIO |
| spRep_AutorizacionAutoimpresion | Genera dataset de reporte/consulta operacional. | @tTipoDoc, @tTipoTramite, @tEstado, @fInicio, @fFinal | TSOLICITUDDETALLE, TSOLICITUD, TTRAMITE, TTABLA |
| spRep_Cancelacion | Genera dataset de reporte/consulta operacional. | @flagTipo, @tCliente, @tTipoDoc, @sOrden, @finicio, @ffinal, @diaContable, @fpago, @sql, @comilla, @sCriterio, @xCriterio, @xFecha | MDOCUMENTO |
| spRep_CobranzaFecha | Genera dataset de reporte/consulta operacional. | @flagTIPOREPORTE, @flagValor, @sPrefijo, @sAno, @sMes, @sFechaPedido, @sFechaDocumento, @sFecha, @dHour, @sql, @comilla, @sCriterio, @scriterio, @flagT | No determinable con certeza desde el código |
| spRep_Comanda | Genera dataset de reporte/consulta operacional. | @flagTipo, @SOrden, @fInicio, @fFinal, @sCriterio, @sql, @Comilla | No determinable con certeza desde el código |
| spRep_ComprobanteDetallado | Genera dataset de reporte/consulta operacional. | @fInicio, @fFinal, @tCliente, @tTipoDoc, @tEstadoDoc, @tCaja, @flagRegVenta, @DiaContable | No determinable con certeza desde el código |
| spRep_ComprobantesVentas | Genera dataset de reporte/consulta operacional. | @tCliente, @tTipoDoc, @tEstadoDoc, @tAutorizacion, @fInicio, @fFinal | MDOCUMENTO |
| spRep_ControlDocumentos | Genera dataset de reporte/consulta operacional. | @fInicio, @fFinal, @tTipoDoc, @tEstadoDoc, @tUsuario, @tCaja | MDOCUMENTO, DDOCUMENTO |
| spRep_ControlEnviosAutorizados | Genera dataset de reporte/consulta operacional. | @fInicio, @fFinal, @tSalon, @tGrupo, @tSubGrupo, @tCodigoProducto, @tCodigoCliente, @turno, @ttipodato, @tMozo, @tUsuario | DPEDIDO, MPEDIDO, TPRODUCTO, TMESA, TDELIVERY, TUSUARIO |
| spRep_ControlMotorizado | Genera dataset de reporte/consulta operacional. | @tMotorizado, @fInicio, @fFinal | MPEDIDO, DPEDIDO |
| spRep_Cortesia | Genera dataset de reporte/consulta operacional. | @tipo, @fInicio, @fFinal, @Forma, @tipoX, @sCriterio, @sql | CPEDIDO |
| spRep_CtaCteIntegrado | Genera dataset de reporte/consulta operacional. | @flagTipo, @fInicio, @fFinal | No determinable con certeza desde el código |
| spRep_CtaCteN | Genera dataset de reporte/consulta operacional. | @flagTDetalle, @flagTResumido, @flagTConsolidado, @fInicio, @fFinal, @Estado, @Cliente, @TipoCC, @SubTipoCC, @sql, @ley1 | TPARAMETRO |
| spRep_CuentasCobrar | Genera dataset de reporte/consulta operacional. | @tipoListado, @tipo, @cliente, @tipoDoc, @estadodoc, @tipoclifrecuente, @tipoctacte, @subtipoctacte, @fInicio, @fFinal | TDELIVERY, MDOCUMENTO |
| spRep_Descuento | Genera dataset de reporte/consulta operacional. | @flag1, @fInicio, @fFinal, @mdesc, @tipoPedido, @sValor, @ExportaExcel, @IMP, @IMP1, @IMP2 | TPARAMETRO, MPEDIDO, DPEDIDO |
| spRep_Diferencia | Genera dataset de reporte/consulta operacional. | @tTurno, @fInicio, @fFinal | No determinable con certeza desde el código |
| spRep_Entregas | Genera dataset de reporte/consulta operacional. | @fInicio, @fFinal, @tCliente, @tEstadoEntrega, @tEstadoCanc, @tGrupo, @tSubGrupo, @tProducto, @tCanalVenta, @tTipo | DPEDIDO, MPEDIDO, TDELIVERY |
| spRep_FormaPagoIntegrado | Genera dataset de reporte/consulta operacional. | @fRegistroi, @fRegistrof, @flagPromedio | MDOCUMENTO, MPEDIDO |
| spRep_Liquidacion | Genera dataset de reporte/consulta operacional. | @flagTipo, @flagTurnoFecha, @flagDiaContable, @flagDocGenerado, @flagEfectivo, @flagCheque, @flagTarjetaCredito, @flagPagosVarios, @flagPunto, @flagPagoCortesia, @flagPorCobrar, @flagReciboIngreso, @flagReciboAnticipo, @flagReciboEgreso, @flagNoCortesia, @flagCortesia, @flagC | MDOCUMENTO, DPEDIDO, TPARAMETRO, MNOTACREDITO |
| spRep_LiquidacionOrigenVenta | Genera dataset de reporte/consulta operacional. | @flagTurno, @flagDiaContable, @sTurno, @sUsuario, @finicio, @ffinal, @sSectorVenta | TORIGENVENTA, TCANALVENTA |
| spRep_LiquidacionOutPut | Genera dataset de reporte/consulta operacional. | @flagTurno, @flagDiaContable, @sTurno, @sUsuario, @sSectorVenta, @finicio, @ffinal, @xFecha, @cortesia, @Ndolar, @xDOlar, @nOtroDoc, @nNoCobrado, @nNeto, @nImpuesto1, @nImpuesto2 | TTIPOCAMBIO, TPARAMETRO, MDOCUMENTO, MPEDIDO |
| spRep_LiquidacionOutPut_NC | Genera dataset de reporte/consulta operacional. | @flagTurno, @flagDiaContable, @sTurno, @sUsuario, @sSectorVenta, @finicio, @ffinal, @xFecha, @cortesia, @Ndolar, @xDOlar, @nOtroDoc, @nNoCobrado, @nNeto, @nImpuesto1, @nImpuesto2 | TTIPOCAMBIO, MNOTACREDITO, MDOCUMENTO, MPEDIDO |
| spRep_LiquidacionSuma | Genera dataset de reporte/consulta operacional. | @flagTurno, @flagDiaContable, @sTurno, @sUsuario, @finicio, @ffinal, @sSectorVenta | TPARAMETRO, MNOTACREDITO |
| spRep_LiquidacionSuma_NC | Genera dataset de reporte/consulta operacional. | @flagTurno, @flagDiaContable, @sTurno, @sUsuario, @finicio, @ffinal, @sSectorVenta | MNOTACREDITO, DDOCUMENTO |
| spRep_Liquidacion_NC | Genera dataset de reporte/consulta operacional. | @flagTipo, @flagTurnoFecha, @flagDiaContable, @flagDocGenerado, @flagEfectivo, @flagCheque, @flagTarjetaCredito, @flagPagosVarios, @flagPunto, @flagPagoCortesia, @flagPorCobrar, @flagReciboIngreso, @flagReciboAnticipo, @flagReciboEgreso, @flagNoCortesia, @flagCortesia, @flagC | MDOCUMENTO, DPEDIDO, MNOTACREDITO |
| spRep_MensajeUsuario | Genera dataset de reporte/consulta operacional. | @fInicio, @fFinal, @tTipo, @tCaja, @tMozo | TMENSAJEUSUARIO, MPEDIDO |
| spRep_Ocupabilidad | Genera dataset de reporte/consulta operacional. | @fInicio, @fFinal, @nHora, @sCriterio, @Opcion, @nContador, @nOrden, @sHora, @nMesa, @nSilla, @nDia, @nHoraInicial | TMESA |
| spRep_PaloteoComparativo | Genera dataset de reporte/consulta operacional. | @flagTurnoOFecha, @flagTipoValor, @tTipoProducto, @tAreaProduccion, @tGrupo, @tSubGrupo, @tCodigoProducto, @tTurno, @finicio, @ffinal, @flagNFacturado, @sql, @comilla, @sCriterio | No determinable con certeza desde el código |
| spRep_PaloteoInsumo | Genera dataset de reporte/consulta operacional. | @dbAlmacen, @Slocal, @Familia, @SubFamilia, @tCodigoProducto, @Area, @flagPlato, @flagCombo, @flagPropiedad, @flagPCombo, @flagTipo, @flagTipoR, @finicio, @ffinal, @Insumo, @t | TPRODUCTOAREA |
| spRep_PaloteoInsumoIntegrado | Genera dataset de reporte/consulta operacional. | @dbAlmacen, @flagPlato, @flagCombo, @flagPropiedad, @flagPCombo, @flagTipo, @finicio, @ffinal, @tTipoPedido, @CodigoOrigenVenta, @sql, @comilla, @flagtipo | No determinable con certeza desde el código |
| spRep_PaloteoOferta | Genera dataset de reporte/consulta operacional. | @flagTipoReporte, @flagTurnoFecha, @flagAutoManual, @sOferta, @sTipoPedido, @sMozo, @sArea, @sCaja, @sUNegocio, @sGrupo, @sSubGrupo, @sTurno, @finicio, @ffinal, @s | No determinable con certeza desde el código |
| spRep_PaloteoProduccion | Genera dataset de reporte/consulta operacional. | @flagFranjaHoraria, @flagConPropiedades, @flagProduccion, @flagVenta, @flagCortesia, @flagCuentaCte, @flagCombinacion, @flagCargo, @flagPedidosFacturados, @flagPedidosFacturadosCortesia, @flagOpcionCom, @flagTurnoOFecha | CPEDIDO, TPRODUCTO, MPEDIDO, DPEDIDO |
| spRep_PaloteoProduccionPorMes | Genera dataset de reporte/consulta operacional. | @flagProduccion, @flagVenta, @flagCortesia, @flagCuentaCte, @flagCombinacion, @flagCargo, @flagPedidosFacturados, @flagPedidosFacturadosCortesia, @flagPrecioCosto, @tSalon, @tTipoProducto, @tMozo, @tTipoPedido, @tArea, @tCaja | No determinable con certeza desde el código |
| spRep_PaloteoPropiedad | Genera dataset de reporte/consulta operacional. | @flagOrigenProduccion, @flagOrigenVenta, @flagOrigenCortesia, @flagOrigenCuentaCorriente, @flagOrigenCombinacion, @flagOrigenCargos, @flagOrigenPedidosFacturados, @flagTurnoOFecha, @flagOpcion1, @flagOpcion2, @flagOpcion3, @flagOpcion4, @dbAlmacen, @turno, @tOperador, @tUnidad | No determinable con certeza desde el código |
| spRep_PaloteoSubProd | Genera dataset de reporte/consulta operacional. | @flagTurno, @tTurno, @tLocal, @tSalon, @tipoProd, @tmozo, @ttipoPedido, @tarea, @tcaja, @tcodigoproducto, @tcodigocliente, @sBoton2, @sBoton3, @sBoton4, @sBoton5, @fIn | No determinable con certeza desde el código |
| spRep_PaloteoVentaIntegrado | Genera dataset de reporte/consulta operacional. | @flagPropiedades, @flagProduccion, @flagVenta, @flagCortesia, @flagCuentaCte, @flagPedidosFacturados, @flagCombinacion, @flagCargo, @flagOpcion, @flagVVenta, @flagVNeto, @flagVCosto, @tTipoProducto, @tTipoPedido, @tGrupo, @tSubGrupo | No determinable con certeza desde el código |
| spRep_Pedido | Genera dataset de reporte/consulta operacional. | @flagFranjaHoraria, @flagTipo, @flagTurnoFecha, @tturno, @tlocal, @tSalon, @tTipoPedido, @tEstadoPedido, @tcaja, @nAdulto, @nPrecuenta, @nEnvio, @fInicio, @fFinal, @OrigenVenta, @ExportaExcel, @formas | No determinable con certeza desde el código |
| spRep_Pedido_GC | Genera dataset de reporte/consulta operacional. | @flagFranjaHoraria, @flagTipo, @flagTurnoFecha, @tturno, @tlocal, @tSalon, @tTipoPedido, @tEstadoPedido, @tcaja, @nAdulto, @nPrecuenta, @nEnvio, @fInicio, @fFinal, @OrigenVenta, @ExportaExcel, @formas | DPEDIDO |
| spRep_PlanillaMovilidad | Genera dataset de reporte/consulta operacional. | @tMotorizado, @finicio, @ffinal, @nMonto, @tTipo, @tTarifa, @sCriterio | No determinable con certeza desde el código |
| spRep_PlanillaMovilidadGeneral | Genera dataset de reporte/consulta operacional. | @tMotorizado, @finicio, @ffinal, @sCriterio, @tmotorizado | No determinable con certeza desde el código |
| spRep_PrincipalCliente | Genera dataset de reporte/consulta operacional. | @flagTipo, @sMonto, @SCliente, @fInicio, @fFinal, @sql, @Filtro, @comilla, @sCliente, @scliente, @filtro | No determinable con certeza desde el código |
| spRep_ProductosNoEnlazados | Genera dataset de reporte/consulta operacional. | @ttipoprod, @tgrupo, @tsubgrupo | TPRODUCTO |
| spRep_Propina | Genera dataset de reporte/consulta operacional. | @fInicio, @fFinal, @sCondicion, @sql, @comilla | No determinable con certeza desde el código |
| spRep_Ranking | Genera dataset de reporte/consulta operacional. | @flagFranjaHoraria, @flagProduccion, @flagVenta, @flagCortesia, @flagCuentaCte, @flagPedidosFacturados, @flagCombinacion, @flagCargo, @flagOpcion, @flagTurnoOFecha, @flagVVenta, @flagVNeto, @flagVCosto, @tSalon, @tTipoProducto | No determinable con certeza desde el código |
| spRep_RankingIntegrado | Genera dataset de reporte/consulta operacional. | @flagProduccion, @flagVenta, @flagCortesia, @flagCuentaCte, @flagPedidosFacturados, @flagCombinacion, @flagCargo, @flagOpcion, @flagVVenta, @flagVNeto, @flagVCosto, @tTipoProducto, @tTipoPedido, @tGrupo, @tSubGrupo, @sPrecio | No determinable con certeza desde el código |
| spRep_ReciboEgreso | Genera dataset de reporte/consulta operacional. | @usuario, @fincial, @ffinal, @tipooper, @bdalmacen | No determinable con certeza desde el código |
| spRep_RegVenta | Genera dataset de reporte/consulta operacional. | @flagTipo, @flagRegVenta, @flagCorrelativo, @flagEstado, @flagAgrupado, @flagRedondeo, @tCliente, @tTipoDoc, @tEstadoDoc, @tCaja, @sOrden, @finicio, @ffinal, @sAno, @sMES, @sFecha, @sFecha2 | No determinable con certeza desde el código |
| spRep_RegVentaDetallado | Genera dataset de reporte/consulta operacional. | @flagTipo, @flagRegVenta, @flagCorrelativo, @flagEstado, @flagAgrupado, @flagRedondeo, @tCliente, @tTipoDoc, @tEstadoDoc, @tCaja, @sOrden, @finicio, @ffinal, @sAno, @sMES, @sFecha, @sFecha2 | No determinable con certeza desde el código |
| spRep_RegVentaIntegrado | Genera dataset de reporte/consulta operacional. | @flagTipo, @flagRedondeo, @flagRegVenta, @tTipoDoc, @tEstadoDoc, @finicio, @ffinal | No determinable con certeza desde el código |
| spRep_RegVentaPagos | Genera dataset de reporte/consulta operacional. | @flagTipo, @flagRegVenta, @flagCorrelativo, @flagEstado, @flagAgrupado, @flagRedondeo, @tCliente, @tTipoDoc, @tEstadoDoc, @tCaja, @sOrden, @finicio, @ffinal, @sAno, @sMES, @sFecha, @sFecha2 | No determinable con certeza desde el código |
| spRep_RegVentaSunat | Genera dataset de reporte/consulta operacional. | @fInicio, @fFinal, @tCliente, @tTipoDoc, @tEstadoDoc, @tCaja, @sOrden, @flagRegVenta, @flagRedondeo, @DiaContable, @ValorTransGratuita, @sEstDoc | No determinable con certeza desde el código |
| spRep_RegVentaSunatAD | Genera dataset de reporte/consulta operacional. | @fInicio, @fFinal, @tCliente, @tTipoDoc, @tEstadoDoc, @tCaja, @sOrden, @flagRegVenta, @flagRedondeo, @DiaContable, @sEstDoc | No determinable con certeza desde el código |
| spRep_RegVentaSunat_formaPago | Genera dataset de reporte/consulta operacional. | @fInicio, @fFinal, @tCliente, @tTipoDoc, @tEstadoDoc, @tCaja, @sOrden, @flagRegVenta, @flagRedondeo, @DiaContable, @sEstDoc, @SqlDinamica, @SqlDinamica_NC | TCLIENTE, MNOTACREDITO, DPAGODOCUMENTO, MDOCUMENTO |
| spRep_RepClieFrecuentes | Genera dataset de reporte/consulta operacional. | @flagFranjaHoraria, @flagDetaxClie, @flagConsoxClie, @flagConsoxZona, @flagNuevxClie, @flagDetaxDia, @flagMonto, @flagTurnoOFecha, @tTipoPedido, @tTipoCliente, @tMotorizado, @tCliente, @tZona, @tGrupo, @tSubGrupo | No determinable con certeza desde el código |
| spRep_ResultadoOperativo | Genera dataset de reporte/consulta operacional. | @flagTurnoOFecha, @flagVista, @tLocal, @tTipoProducto, @tTipoPedido, @tAreaProduccion, @tGrupo, @tSubGrupo, @tCodigoProducto, @tTurno, @finicio, @ffinal, @flagFacturado, @sql | No determinable con certeza desde el código |
| spRep_Rotacion | Genera dataset de reporte/consulta operacional. | @fInicio, @fFinal, @sCriterio | TMESA, TCANALVENTA |
| spRep_SUNATtxt | Genera dataset de reporte/consulta operacional. | @fInicio, @fFinal, @tCliente, @tTipoDoc, @tEstadoDoc, @tCaja, @sOrden, @flagRegVenta, @flagRedondeo, @DiaContable, @sEstDoc | TCLIENTE |
| spRep_TiempoDelivery | Genera dataset de reporte/consulta operacional. | @tEmpacador, @tMotorizado, @finicio, @ffinal, @tipo, @sql, @comilla, @sCriterio, @CanalVenta | TCANALVENTA |
| spRep_TiempoDeliveryIntegrado | Genera dataset de reporte/consulta operacional. | @finicio, @ffinal, @R1, @R2, @sql, @comilla, @sCriterio, @scriterio, @FINICIO | No determinable con certeza desde el código |
| spRep_TiempoSalon | Genera dataset de reporte/consulta operacional. | @flagTurnoOFecha, @tLocal, @tSalon, @tTurno, @finicio, @ffinal, @sql, @Actualizaciones, @comilla, @sCriterio, @xCriterio | No determinable con certeza desde el código |
| spRep_TipoProductoVentaIntegrado | Genera dataset de reporte/consulta operacional. | @flagPropiedades, @flagProduccion, @flagVenta, @flagCortesia, @flagCuentaCte, @flagPedidosFacturados, @flagCombinacion, @flagCargo, @flagOpcion, @flagVVenta, @flagVNeto, @flagVCosto, @tTipoProducto, @tTipoPedido, @tGrupo, @tSubGrupo | No determinable con certeza desde el código |
| spRep_VentaCompMensual | Genera dataset de reporte/consulta operacional. | @sPrecio, @sAno1, @sAno2, @sMes, @sTipo1, @sTipo2, @sTipo3, @sFecha1, @sFecha2, @dHour, @Opcion, @criterio, @sql, @comilla | No determinable con certeza desde el código |
| spRep_VentaFecha | Genera dataset de reporte/consulta operacional. | @sPrecio, @sAno, @sMes, @sFecha, @dHour, @criterio, @tipooper, @sql, @comilla | MDOCUMENTO, DDOCUMENTO, MNOTACREDITO |
| spRep_VentaIntervaloIntegrado | Genera dataset de reporte/consulta operacional. | @INTERVALO, @TipoPedido1, @TipoPedido2, @TipoPedido3, @TipoPedido4, @TipoPedido5, @TipoPedido6, @DIA1, @DIA2, @DIA3, @DIA4, @DIA5, @DIA6, @DIA7, @fInicio | MPEDIDO |
| spRep_VentaMensualCanalesIntegrado | Genera dataset de reporte/consulta operacional. | @sPrecio, @sAno, @sMes, @sFecha, @dHour, @sql, @comilla | No determinable con certeza desde el código |
| spRep_VentaMensualIntegrado | Genera dataset de reporte/consulta operacional. | @sPrecio, @sAno, @sMes, @sFecha, @dHour, @tipo, @metaMensual, @diames, @sql, @comilla | No determinable con certeza desde el código |
| spUpd_MPEDIDO | Actualiza registros del proceso indicado por el nombre. | @tCliente, @tTipoPedido, @lPrioridad, @tTipoAtencion, @tMozo, @tMotorizado, @tObservacion, @nTiempo, @tPuntoVenta, @tHabitacion, @tReserva, @tPasajero, @tCompania, @tContacto, @nDescuento, @t | TMESA, MPEDIDO |
| spUpdate_DDOCUMENTO_Ina | Actualiza registros del proceso indicado por el nombre. | @tdocumento, @tcodigoProducto, @tItem, @cant, @tAplicaExoneracion, @tAplicaInafectacion, @FETCH_STATUS | DDOCUMENTO |
| spUpdate_DPEDIDO_Ina | Actualiza registros del proceso indicado por el nombre. | @pedido, @tcodigoProducto, @tItem, @cant, @tAplicaExoneracion, @tAplicaInafectacion, @FETCH_STATUS | DPEDIDO |
| sp_ActualizaTablas | Actualiza registros del proceso indicado por el nombre. | @sServidor, @sBaseDato, @sParametro, @isql, @NombreTabla, @Isql | TGRUPOACCESO, TGRUPOUSUARIO, TACCESO, TCAJA, TAREAIMPRESORA, TCLIENTE, TCOMPANIA, TTABLA, TDELIVERY, TGRUPO |
| sp_AsignaAnticipo_Pedido | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @tCodigoPedido, @tCodigoReserva, @TipoOper, @resultado, @titemPedido, @TipoPedido | DPEDIDO, MPEDIDO, MDOCUMENTO, DDOCUMENTO, TPRODUCTO |
| sp_CargaPedidoReqExternoDirectosInforest | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @fch_fFechaInicio, @fch_fFechaFin | No determinable con certeza desde el código |
| sp_CargaPedidoReqExternoInforest | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @fch_fFechaInicio, @fch_fFechaFin | No determinable con certeza desde el código |
| sp_ComparativoConsumo | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @BDALMACEN, @finicio, @ffinal, @vch_SubArea, @vch_Condicion, @SQL | No determinable con certeza desde el código |
| sp_CopiaArchivosRemotos | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @ServerRemoto, @BddInforest, @Isql | TPARAMETRO, TTABLA, TGRUPO, TSUBGRUPO, TPRODUCTO, TPROPIEDAD, TOPERADOR |
| sp_CreaTemporalSocio_SyBASE | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | Sin parámetros explícitos detectados | No determinable con certeza desde el código |
| sp_InsUptSocioDelivery_SyBASE | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | Sin parámetros explícitos detectados | TDELIVERY |
| sp_OptimizarBD | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @xbd | MDOCUMENTO |
| sp_OptimizarBD_Bizlinks | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @_BD_INFOREST, @_dias, @nSql, @sizex, @sizes, @_BD_IN | No determinable con certeza desde el código |
| sp_RepLiquidacion3 | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @finicio, @ffin, @turno, @tipooper | MNOTACREDITO, TPARAMETRO, MEGRESO |
| sp_TraeDatosPagos_SyBase | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @FecInicio, @FecFinal | DPAGODOCUMENTO, MDOCUMENTO |
| sp_TraeDatosVentas_SyBase | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @FecInicio, @FecFinal | MDOCUMENTO, TCLIENTE |
| sp_UpdFotoDelivery | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @tCodigo, @oFoto | No determinable con certeza desde el código |
| sp_UpdFotoProducto | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @tCodigo, @oFoto | TPRODUCTO |
| sp_UpdImagenCaja | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @tCodigo, @tTipo, @oFoto | No determinable con certeza desde el código |
| sp_VinculacionSAP | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @servidor, @BD, @CodLocal, @Variable, @query, @opcion, @comilla, @tabla, @tabla2 | TPRODUCTO, TINSUMO |
| ups_ObtieneFechaHora | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | No determinable con certeza desde el código |
| usp_ActualizarCabPeDoc | Actualiza registros del proceso indicado por el nombre. | @opcion, @param1, @param2, @param3, @param4, @param5, @param6, @pedido | MDOCUMENTO, LOG_PEDIDO_DOCUMENTO, DDOCUMENTO, MPEDIDO, DPEDIDO, ruc_temp, TCLIENTE, TUBIGEO |
| usp_AdmCen_ActualizarTablaCampos | Actualiza registros del proceso indicado por el nombre. | @tcodigo, @tAgrupacion, @lestado, @tUsuario | No determinable con certeza desde el código |
| usp_AdmCen_ActualizarTablaPrincipal | Actualiza registros del proceso indicado por el nombre. | @tcodigo, @lestado, @tUsuario | No determinable con certeza desde el código |
| usp_AdmCen_ObtenerCamposxTablas | Consulta y devuelve datos del proceso indicado por el nombre. | @sModulo, @sServidor, @sBaseDato, @sUsuario, @ErrorCode, @cantidadFilas, @contador, @camposIns, @campos, @keysIns, @keys, @NombreTabla, @isql, @keysva, @campo, @comilla | TIMPORTACION |
| usp_AdmCen_ObtenerDependenciaPadres | Consulta y devuelve datos del proceso indicado por el nombre. | @tcodigo, @Agrupacion, @agrupacion, @AGRUPACION | No determinable con certeza desde el código |
| usp_AdmCen_ObtenerPermisosUsuario | Consulta y devuelve datos del proceso indicado por el nombre. | @tUsuario, @codigoGrupo, @codigogrupo | No determinable con certeza desde el código |
| usp_AdmCen_ObtenerTablas | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | TIMPORTACION |
| usp_AdmCen_ObtenerTablasPrincipales | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | No determinable con certeza desde el código |
| usp_Anticipos | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @opcion, @codigocliente, @tCaja, @desde, @hasta | DPEDIDO, MPEDIDO, MDOCUMENTO, DDOCUMENTO, TRESERVA |
| usp_Aud_ObtenerEstadoImpreso | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | No determinable con certeza desde el código |
| usp_Aud_ObtenerMotivoEliminacion | Elimina o depura registros del proceso indicado por el nombre. | Sin parámetros explícitos detectados | No determinable con certeza desde el código |
| usp_Aud_ObtenerMozos | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | No determinable con certeza desde el código |
| usp_Aud_ObtenerSalones | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | No determinable con certeza desde el código |
| usp_Aud_ObtenerUsuarios | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | TUSUARIO |
| usp_Aud_ValidadUsuario | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @susuario, @scontrasenia | No determinable con certeza desde el código |
| usp_AutSol_ActualizaEstadoSolicitudDetalle | Actualiza registros del proceso indicado por el nombre. | @tCodigoSolicitud, @tItem, @tEstadoSolicitud, @tUsuario, @nItem, @nitem | TSOLICITUDDETALLE |
| usp_AutSol_ActualizarSolicitudCab | Actualiza registros del proceso indicado por el nombre. | @numeroAutorizacion, @numeroAutorizacionAnterior, @fFechaEmision, @fFechaCaducidad, @tUsuario, @tCodigoSolicitud | TSOLICITUD, TSOLICITUDDETALLE |
| usp_AutSol_ActualizarTiposDeDocumentos | Actualiza registros del proceso indicado por el nombre. | @tTipo, @tCaja, @tTipoEmision, @tSerie, @tNumeroAutorizacion, @fregistro, @tCaducidad, @tcaducidad, @tcaja, @tTipoEMision, @tserie | TTIPODOCUMENTOIMPRESORA |
| usp_AutSol_AnularSolicitud | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @tcodigoSolicitud, @tUsuario | TSOLICITUD |
| usp_AutSol_InsertarSolicitudCab | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @codigoTramite, @codigoEstadoSolicitud, @numeroAutorizacion, @numeroAutorizacionAnterior, @fFechaEmision, @fFechaCaducidad, @tUsuario, @tCodigoSolicitud, @codigoSolicitud, @codigoEstadoSoli | TSOLICITUD |
| usp_AutSol_InsertarSolicitudDetalle | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @tCodigoSolicitud, @tTipoDocumento, @tSerieEstablecimiento, @tSerieCaja, @tFolioInicial, @tFolioFinal, @tEstadoSolicitud, @tUsuario, @tCaja, @nCorrelativo, @nItem, @nitem, @tEstadoSolicit | TSOLICITUDDETALLE |
| usp_AutSol_ObtenerCaja | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | TCAJA |
| usp_AutSol_ObtenerCliente | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | No determinable con certeza desde el código |
| usp_AutSol_ObtenerEstablecimientoSeries | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | TTIPODOCUMENTOIMPRESORA |
| usp_AutSol_ObtenerEstadoDocumento | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | No determinable con certeza desde el código |
| usp_AutSol_ObtenerEstadoTramite | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | TTABLA |
| usp_AutSol_ObtenerSolicitud | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | TSOLICITUD |
| usp_AutSol_ObtenerSolicitudes | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | No determinable con certeza desde el código |
| usp_AutSol_ObtenerTipoDocumento | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | No determinable con certeza desde el código |
| usp_AutSol_ObtenerTipoDocumentoSerieAProcesar | Consulta y devuelve datos del proceso indicado por el nombre. | @tTipo, @tCodigoSolicitud | TTIPODOCUMENTOIMPRESORA, TSOLICITUDDETALLE |
| usp_AutSol_ObtenerTipoDocumentoSerieActualizar | Actualiza registros del proceso indicado por el nombre. | @tCodigoSolicitud | No determinable con certeza desde el código |
| usp_AutSol_ObtenerTipoDocumentoSerieAntes | Consulta y devuelve datos del proceso indicado por el nombre. | @tCodigoSolicitud | No determinable con certeza desde el código |
| usp_AutSol_ObtenerTipoTramite | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | TTRAMITE |
| usp_AutSol_ObtenerTiposDocumentos | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | No determinable con certeza desde el código |
| usp_AutSol_ObtenerTiposDocumentosActualizar | Actualiza registros del proceso indicado por el nombre. | @tCodigoSolicitud | TSOLICITUDDETALLE |
| usp_AutSol_ObtenerTramite | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | TTRAMITE |
| usp_AutSol_ObtieneAutorizacionAnterior | Consulta y devuelve datos del proceso indicado por el nombre. | @codigo | TTRAMITE |
| usp_AutSol_ValidadUsuario | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @susuario, @scontrasenia | No determinable con certeza desde el código |
| usp_AutSol_obtenerSolicitudXAutorizacion | Consulta y devuelve datos del proceso indicado por el nombre. | @tNumeroAutorizacion | TSOLICITUD |
| usp_ControlServicioWindows | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @NombreServicio, @Accion, @Resultado, @Comando, @MensajeExito, @AccionUpper | No determinable con certeza desde el código |
| usp_FE_ObtieneCodigoBHQ | Consulta y devuelve datos del proceso indicado por el nombre. | @Documento, @TipoImagen, @NotaCredito | No determinable con certeza desde el código |
| usp_FE_factObtieneCodigoBHQ | Consulta y devuelve datos del proceso indicado por el nombre. | @Documento, @TipoImagen, @NotaCredito, @Resp, @FEpape | TPARAMETRO, MNOTACREDITO, TTIPODOCUMENTO, MDOCUMENTO, TTIPODOCUMENTOIMPRESORA |
| usp_GenInsertarDiaContable | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @tUsuario, @fDiaContable | No determinable con certeza desde el código |
| usp_GenObtieneDiaContable | Consulta y devuelve datos del proceso indicado por el nombre. | @lDiaContable, @sHoraCierre, @tUsuario, @fDiaContable | No determinable con certeza desde el código |
| usp_InforestAdm_ObtenerOfertasReporte | Genera dataset de reporte/consulta operacional. | @Oferta | No determinable con certeza desde el código |
| usp_InforestAdm_ObtenerOperadoresReporteLocal | Genera dataset de reporte/consulta operacional. | Sin parámetros explícitos detectados | No determinable con certeza desde el código |
| usp_InforestAdm_ObtenerProductosReporteLocal | Genera dataset de reporte/consulta operacional. | @tcodigoCanal, @tTipoReporte, @filtro | TPRODUCTO |
| usp_InforestAdm_ObtenerProductosReporteLocal2 | Genera dataset de reporte/consulta operacional. | @tcodigoCanal, @tTipoReporte, @filtro | TPRODUCTO |
| usp_InforestAdm_ObtenerPrropiedadesReporte | Genera dataset de reporte/consulta operacional. | Sin parámetros explícitos detectados | No determinable con certeza desde el código |
| usp_InforestCon_ObtenerReporteLiquidacionVentas | Genera dataset de reporte/consulta operacional. | @tTurno, @fRegistroi, @fRegistrof, @tUsuario, @flagDiaContable, @sSectorVenta, @cortesia, @flagdiaContable | TCAJA |
| usp_InforestCon_ObtenerReporteLiquidacionVentas_NC | Genera dataset de reporte/consulta operacional. | @tTurno, @fRegistroi, @fRegistrof, @tUsuario, @flagDiaContable, @sSectorVenta, @cortesia, @flagdiaContable | MNOTACREDITO, TCAJA |
| usp_Inforest_ActualizaCostos | Actualiza registros del proceso indicado por el nombre. | @tNombreInforest, @tNombreAlmacen, @tLocal, @sql, @comilla | No determinable con certeza desde el código |
| usp_Inforest_DescargoVenta | Gestiona stock, descargo o kardex. | @Almacen, @fechaIni, @fechaFin, @sTemporal, @Local, @Pedido, @tipooper, @Isql, @Comilla, @comilla | No determinable con certeza desde el código |
| usp_Inforest_DescargoVentaPorInsumo | Gestiona stock, descargo o kardex. | @Almacen, @fechaIni, @fechaFin, @sTemporal, @Local, @Pedido, @tipooper, @Isql, @Comilla, @comilla | TDESCARGOINSUMO |
| usp_Inforest_Impresion | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @Codigo, @tipooper, @val, @Mont | DDOCUMENTO, DPAGODOCUMENTO, MDOCUMENTO |
| usp_Inforest_InicializaCostos | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | Sin parámetros explícitos detectados | TPRODUCTO, TPROPIEDAD |
| usp_Inforest_InsertaDocumento | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @Pedido, @Documento, @tTipoDocumento, @tCodigoCliente, @tEstadoDocumento, @tCaja, @tTurno, @tSalon, @tUsuario, @tUsuarioAutoriza, @fDiaContable, @tDescuento, @tConsumo, @lImpresionMonedaExtranjera, @tautorizacion, @tcodigocontrol, @Cortesia, @fInicio, @fCaducidad, @tContribuyenteEspecial, @tipooper, @femision, @nPrecioNeto, @nimpuesto1, @nimpuesto2, @nimpuesto3, @nventa, @nDescuento, @Recargo, @tReservaInf, @anticipo, @Ley1, @TotalLey1, @tOrdenCompra | DPEDIDO, DDOCUMENTO, MDOCUMENTO |
| usp_Inforest_InsertaLogDescargo | Gestiona stock, descargo o kardex. | @num_nCorrelativo, @dtm_fRegInicio, @vch_tUsuario, @dtm_fFecInicio, @dtm_fFecFin, @num_nCantPedidos, @num_nCantPedOk, @num_nCantPedError, @vch_tEquIp, @vch_tEquName, @vch_tEquUser, @vchMsgError | No determinable con certeza desde el código |
| usp_Inforest_InsertarLogErrores | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @Ttabla, @Proceso, @CodError, @ErrorProcedure, @ErrorLine, @ErrorMensaje, @DatoAlternativo, @Observaciones, @Usuario, @tipooper, @ID, @ErrorMe | LOG_INFOREST |
| usp_Inforest_ListaConectado | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @pc | No determinable con certeza desde el código |
| usp_Inforest_ObtenerClientesFrecuente | Consulta y devuelve datos del proceso indicado por el nombre. | @tcodigo, @tapellido, @tnombre, @sql, @comilla, @filtro | No determinable con certeza desde el código |
| usp_Inforest_ObtenerInsumosCriticos | Consulta y devuelve datos del proceso indicado por el nombre. | @tCodigoPedido | No determinable con certeza desde el código |
| usp_Inforest_ObtieneAccesoMenu | Consulta y devuelve datos del proceso indicado por el nombre. | @tusuario, @tModulo, @tFormularioName, @tcodigoProducto, @cant, @ninsumo, @FETCH_STATUS | TPRODUCTO, CPEDIDO, TUSUARIO |
| usp_Inforest_ObtieneCentralPedidos_Adicicion | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | MPEDIDO, TDELIVERY, DPEDIDO |
| usp_Inforest_ObtieneClientesFactura | Consulta y devuelve datos del proceso indicado por el nombre. | @tcodigoDelivery, @tTipoDocumento, @nMonto, @tcaja | TCAJA |
| usp_Inforest_ObtieneCodigoQR_Bol | Consulta y devuelve datos del proceso indicado por el nombre. | @tDocumento, @cadenaQR, @nitEmisor, @numFactura, @numAutorizacion, @fechaEmision, @nTotal, @codControl, @nitReceptor, @nDescuento | TPARAMETRO, MDOCUMENTO |
| usp_Inforest_ObtieneCodigoQR_CB_HASH_FACT | Consulta y devuelve datos del proceso indicado por el nombre. | @tDocumento, @tipoCodigo, @cadena, @nitEmisor, @TIPODOCU, @IGV, @numFactura, @serie, @caja, @fechaEmision, @nTotal, @codDocCli | MNOTACREDITO, TPARAMETRO, TTIPODOCUMENTO, MDOCUMENTO, TTIPODOCUMENTOIMPRESORA |
| usp_Inforest_ObtieneCombos | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | No determinable con certeza desde el código |
| usp_Inforest_ObtieneDocumentos | Consulta y devuelve datos del proceso indicado por el nombre. | @FechaIni, @FechaFin, @Caja | No determinable con certeza desde el código |
| usp_Inforest_ObtieneDocumentos_NC_bizlink | Consulta y devuelve datos del proceso indicado por el nombre. | @FechaIni, @FechaFin, @Caja, @nSQL, @BD, @XDOC, @tDocumento, @cant | MNOTACREDITO, TESTADOBIZLINK, MDOCUMENTO |
| usp_Inforest_ObtieneDocumentos_bizlink | Consulta y devuelve datos del proceso indicado por el nombre. | @FechaIni, @FechaFin, @Caja, @newfechaini, @nSQL, @BD, @XDOC, @tDocumento, @cant | MDOCUMENTO, TESTADOBIZLINK, TCLIENTE |
| usp_Inforest_ObtieneOperadoresXProducto | Consulta y devuelve datos del proceso indicado por el nombre. | @tProducto | No determinable con certeza desde el código |
| usp_Inforest_ObtienePedidosPorMesas | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | No determinable con certeza desde el código |
| usp_Inforest_ObtienePedidosPorSocios | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | No determinable con certeza desde el código |
| usp_Inforest_ObtieneProductos | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | No determinable con certeza desde el código |
| usp_Inforest_ObtienePropiedad | Consulta y devuelve datos del proceso indicado por el nombre. | @tAlmacen, @tBdAlmacen, @tLocal, @sql, @comilla, @CodigoOperador, @DescripcionOperador, @s | No determinable con certeza desde el código |
| usp_Inforest_ObtienePropiedadesBusqueda | Consulta y devuelve datos del proceso indicado por el nombre. | @CodigoOperador, @CodigoProducto, @CodigoPedido, @tItem, @bdAlmacen, @sLocal, @sql, @comilla, @tparametro | TPARAMETRO, MPEDIDO, VISIBILIDADPROPIEDADXCANAL |
| usp_Inforest_ObtienePropiedadesComboBusqueda | Consulta y devuelve datos del proceso indicado por el nombre. | @CodigoOperador, @CodigoProducto, @CodigoPedido, @tItem, @tItemCombo, @bdAlmacen, @sLocal, @sql, @comilla, @tparametro | TPARAMETRO, MPEDIDO, VISIBILIDADPROPIEDADXCANAL |
| usp_Inforest_ObtieneRecargas | Consulta y devuelve datos del proceso indicado por el nombre. | @FechaIni, @FechaFin, @tipo, @opcion, @param | TMOVIMIENTOTARJETASRFID |
| usp_Inforest_ObtieneRptGuia | Consulta y devuelve datos del proceso indicado por el nombre. | @tGuiaTransporte, @Almacen, @ruc, @razon, @direccion, @email, @texto, @texto2, @logo, @Isql | TPARAMETRO, TPRODUCTO |
| usp_Inforest_ObtieneSubAlmacen | Consulta y devuelve datos del proceso indicado por el nombre. | @lMultiAreaSubGrupo, @lMultiAreaCaja, @sCaja, @sSubGrupo, @tSubAlmacen, @tAreaProduccion, @count, @scaja, @ssubgrupo | TAREASUBGRUPO |
| usp_Inforest_ObtieneUsuarios | Consulta y devuelve datos del proceso indicado por el nombre. | @Modulo | No determinable con certeza desde el código |
| usp_Inforest_PedidosCentralPedido | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @fechaInicio, @fechaFin | MPEDIDO, TDELIVERY, DPEDIDO |
| usp_Inforest_RevertirInsumosCriticos | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @tCodigoPedido, @sitem | No determinable con certeza desde el código |
| usp_Inforest_RevertirInsumosCriticosCabecera | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @tCodigoPedido | No determinable con certeza desde el código |
| usp_Inforest_ValidaClienteSel | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @tTipoDoc, @tCodCliente, @tRespuesta, @lRuc, @lOtros, @tTipoIden, @VTipoResponsable, @pais | No determinable con certeza desde el código |
| usp_Inforest_ValidaDeliveryCliente | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @tcodigoDelivery, @tcodigoClienteF, @cant, @correla | No determinable con certeza desde el código |
| usp_Inforest_killListaConec | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @pc, @kill | No determinable con certeza desde el código |
| usp_ListDocumentosFE | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @caja, @FechaIni, @FechaFin, @tipooper | MNOTACREDITO |
| usp_ObtenerCaja | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | TCAJA |
| usp_ObtenerCliente | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | No determinable con certeza desde el código |
| usp_ObtenerEstadoDocumento | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | No determinable con certeza desde el código |
| usp_ObtenerTipoDocumento | Consulta y devuelve datos del proceso indicado por el nombre. | Sin parámetros explícitos detectados | No determinable con certeza desde el código |
| usp_Rec_Consumo | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @fi, @ff, @con | MDOCUMENTO, DPAGODOCUMENTO, TTARJETACREDITO |
| usp_RepInforest_DescargoVenta | Genera dataset de reporte/consulta operacional. | @Almacen, @fechaIni, @fechaFin, @sTemporal, @Local, @Grupo, @SubGrupo, @Insumo, @Area, @Descargo, @tipooper, @Isql, @Comilla, @filtro, @comilla | TPARAMETRO, DPEDIDO |
| usp_Seg_cLientes | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @tModulo, @tDbName, @nCantidad, @nlicencias, @nlicenciasActuales | No determinable con certeza desde el código |
| usp_Seg_verConexiones | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @dbName, @ProgramName | No determinable con certeza desde el código |
| usp_TicketComandaVistaPrevia | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @tCodigoPedido | No determinable con certeza desde el código |
| usp_ValidaTipoIdentidad | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @tCodigoIdentidad, @Valor, @Validacion, @nLongitud, @lMayorIgual, @lMenorIgual, @lIgual, @lTipoDato | No determinable con certeza desde el código |
| usp_WS_Inforest | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @Codigo, @tipooper | TTARJETACREDITO, DPREPAGO, DDOCUMENTO, MDOCUMENTO, DPAGODOCUMENTO, MPEDIDO, TDELIVERY, TCLIENTE |
| usp_inforest_CentralPedidoRegistro | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @fInicio, @fFinal, @tCliente, @tEstadoEntrega, @tProducto, @tCanalVenta | DPEDIDO, MPEDIDO, TDELIVERY |
| usp_validarUsuario | No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @susuario, @scontrasenia | No determinable con certeza desde el código |
| Calcular_BaseImponible_gc (función) | Función SQL auxiliar; No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @opcion, @tNotaCredito | MNOTACREDITO |
| CreaTabla (función) | Función SQL auxiliar; No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @Flag, @Date | No determinable con certeza desde el código |
| ObteniendoDatosDOC (función) | Función SQL auxiliar; No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @opcion, @tCodigoPedidoo | DDOCUMENTO, MDOCUMENTO, DPAGODOCUMENTO, TTARJETACREDITO, MPEDIDO, TMESA, TESTADOBIZLINK |
| ObteniendoExoIna (función) | Función SQL auxiliar; No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @opcion, @Tdocumento | DDOCUMENTO |
| ObteniendoOrigenVenta (función) | Función SQL auxiliar; No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @Tdocumento | DDOCUMENTO, MPEDIDO |
| ObteniendoRedondeo (función) | Función SQL auxiliar; No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @x | No determinable con certeza desde el código |
| fn_cortesia_calculo (función) | Función SQL auxiliar; No determinable con certeza desde el código; el nombre sugiere una rutina operacional específica. | @tipo, @tDocumento, @tCodigoProducto, @tItem | No determinable con certeza desde el código |

## C. Catálogo completo de vistas detectadas en `4. Vistas.sql` y `8. InfoFact.sql`

| Vista | Propósito inferido | Tablas fuente |
| --- | --- | --- |
| vAreaChef | No determinable con certeza desde el código; el nombre sugiere una vista de apoyo. | No determinable con certeza desde el código |
| vBajaBoletaInfoFact | Vista de salida para facturación electrónica InfoFact. | No determinable con certeza desde el código |
| vBajaFacturaInfoFact | Vista de salida para facturación electrónica InfoFact. | No determinable con certeza desde el código |
| vBajasDetalleInfoFact | Vista de salida para facturación electrónica InfoFact. | TESTADOINFOFACT |
| vBalanza | No determinable con certeza desde el código; el nombre sugiere una vista de apoyo. | No determinable con certeza desde el código |
| vCajaCodigoControl | No determinable con certeza desde el código; el nombre sugiere una vista de apoyo. | No determinable con certeza desde el código |
| vCompania | No determinable con certeza desde el código; el nombre sugiere una vista de apoyo. | No determinable con certeza desde el código |
| vDatosAdicionalesInfoFact | Vista de salida para facturación electrónica InfoFact. | MDOCUMENTO, TPARAMETRO |
| vDelivery | Vista operativa de delivery/logística. | No determinable con certeza desde el código |
| vDocumentoAnticiposInfoFact | Abstracción de documentos emitidos o tributarios. | DDOCUMENTO, MDOCUMENTO |
| vDocumentoCabeceraInfoFact | Abstracción de documentos emitidos o tributarios. | MPEDIDO, DDOCUMENTO, MGUIATRANSPORTE, MDOCUMENTO, TUBIGEO, TTIPOCAMBIO, TCAJA, TPARAMETRO |
| vDocumentoCreditosInfoFact | Abstracción de documentos emitidos o tributarios. | MDOCUMENTO, TTIPOCAMBIO, TPARAMETRO, MNOTACREDITO |
| vDocumentoDetalleInfoFact | Abstracción de documentos emitidos o tributarios. | DDOCUMENTO, MDOCUMENTO, TTIPOCAMBIO, TPARAMETRO |
| vDocumentoGrilla | Abstracción de documentos emitidos o tributarios. | No determinable con certeza desde el código |
| vDocumentoLeyendasInfoFact | Abstracción de documentos emitidos o tributarios. | MDOCUMENTO, MNOTACREDITO, TPARAMETRO |
| vDocumentoRegistroVentas | Abstracción de documentos emitidos o tributarios. | DDOCUMENTO, MDOCUMENTO, TPARAMETRO, MNOTACREDITO |
| vDocumentoRegistroVentas_TransGratuita | Abstracción de documentos emitidos o tributarios. | DDOCUMENTO, MDOCUMENTO, TPARAMETRO, MNOTACREDITO |
| vEgreso | No determinable con certeza desde el código; el nombre sugiere una vista de apoyo. | TCENTROCOSTO |
| vEstadoFrecuente | Vista catálogo/paramétrica para combos, estados o tipos. | No determinable con certeza desde el código |
| vEstadoSolicitud | Vista catálogo/paramétrica para combos, estados o tipos. | No determinable con certeza desde el código |
| vEstadoSolicitudDetalle | Vista catálogo/paramétrica para combos, estados o tipos. | No determinable con certeza desde el código |
| vGuiaTransporte | No determinable con certeza desde el código; el nombre sugiere una vista de apoyo. | No determinable con certeza desde el código |
| vMaitre | No determinable con certeza desde el código; el nombre sugiere una vista de apoyo. | No determinable con certeza desde el código |
| vMotivoAnulacion | No determinable con certeza desde el código; el nombre sugiere una vista de apoyo. | No determinable con certeza desde el código |
| vMotivoNotaCredito | No determinable con certeza desde el código; el nombre sugiere una vista de apoyo. | TTABLA |
| vMotivoReserva | No determinable con certeza desde el código; el nombre sugiere una vista de apoyo. | TTABLA |
| vNotaCreditoCabeceraInfoFact | Abstracción de nota de crédito para operación/FE/impresión. | MNOTACREDITO, MDOCUMENTO, TUBIGEO, TTIPOCAMBIO, TCAJA, DDOCUMENTO, TPARAMETRO, DNOTACREDITO |
| vNotaCreditoDetalleInfoFact | Abstracción de nota de crédito para operación/FE/impresión. | MNOTACREDITO, DDOCUMENTO, MDOCUMENTO, TTIPOCAMBIO, TPARAMETRO, DNOTACREDITO |
| vNotaCreditoImpresora | Abstracción de nota de crédito para operación/FE/impresión. | No determinable con certeza desde el código |
| vNotaCreditoSunat | Abstracción de nota de crédito para operación/FE/impresión. | MNOTACREDITO, MDOCUMENTO, DDOCUMENTO, TPARAMETRO, DNOTACREDITO |
| vOrigenVenta | No determinable con certeza desde el código; el nombre sugiere una vista de apoyo. | No determinable con certeza desde el código |
| vPaisOrigen | No determinable con certeza desde el código; el nombre sugiere una vista de apoyo. | No determinable con certeza desde el código |
| vPaloteoProduccionPropiedades | No determinable con certeza desde el código; el nombre sugiere una vista de apoyo. | No determinable con certeza desde el código |
| vPaloteoProduccionPropiedadesCombos | No determinable con certeza desde el código; el nombre sugiere una vista de apoyo. | No determinable con certeza desde el código |
| vProductoSunat | No determinable con certeza desde el código; el nombre sugiere una vista de apoyo. | No determinable con certeza desde el código |
| vSectorVentaCajaR | No determinable con certeza desde el código; el nombre sugiere una vista de apoyo. | No determinable con certeza desde el código |
| vSucursal | No determinable con certeza desde el código; el nombre sugiere una vista de apoyo. | No determinable con certeza desde el código |
| vTPROGRAMAPRECIOS_DETA | No determinable con certeza desde el código; el nombre sugiere una vista de apoyo. | TPROGRAMAPRECIOS_DETA |
| vTarjetaCredito | No determinable con certeza desde el código; el nombre sugiere una vista de apoyo. | No determinable con certeza desde el código |
| vTienda | No determinable con certeza desde el código; el nombre sugiere una vista de apoyo. | No determinable con certeza desde el código |
| vTipoClienteFrecuente | Vista catálogo/paramétrica para combos, estados o tipos. | No determinable con certeza desde el código |
| vTipoDocumento | Vista catálogo/paramétrica para combos, estados o tipos. | No determinable con certeza desde el código |
| vTipoGrupoCliente | Vista catálogo/paramétrica para combos, estados o tipos. | TTABLA |
| vTipoResponsable | Vista catálogo/paramétrica para combos, estados o tipos. | No determinable con certeza desde el código |

## D. Configuraciones específicas por país

| Script | Configuración observada |
|---|---|
| `scriptPeruAlIniciar.sql` | Fija `PAISORIGEN=000`, moneda local `S/.`, secundaria `US$` y nombres tributarios/monetarios orientados a Perú. |
| `scriptChileAlIniciar.sql` | Fija `PAISORIGEN=004`, moneda `CL$`/US$ y semántica monetaria chilena. |
| `scriptBoliviaAlIniciar.sql` | Fija `PAISORIGEN=001`, desactiva impuesto2/impuesto3, ajusta longitud tributaria, multiarea y montos impositivos por documento. |
| `scriptEcuadorAlIniciar.sql` | Fija `PAISORIGEN=002`, moneda base dólares y variantes para emisión ecuatoriana. |
| `scriptEspanaAlIniciar.sql` | Fija `PAISORIGEN=005`, moneda `€` y US$ para operación española. |
| `scriptArgentinaAlIniciar.sql` | Fija `PAISORIGEN=003`, recrea `MONEDA`, `TIPOIDENTIDAD` y `TIPORESPONSABLE` con semántica AFIP (CUIT/CUIL/CDI/DNI/Pasaporte). |

### Observaciones adicionales por país desde el código

- **Perú**: `InfoFact.sql` usa códigos SUNAT, moneda `PEN`, país `PE`, ubigeo peruano, detracción/retención, y vistas `vDocumentoCabeceraInfoFact` / `vNotaCreditoCabeceraInfoFact`.
- **Bolivia**: el script ajusta impuestos secundarios a cero, modifica restricciones operativas (`lMultiArea`, `tSubalmacen`) y la longitud de identificación tributaria.
- **Ecuador**: existe además `Ejecuta Columnas_Ecuador.sql`, que extiende la longitud de claves de guía de transporte a `varchar(16)`.
- **Argentina**: reconstruye catálogos tributarios completos (`TIPOIDENTIDAD`, `TIPORESPONSABLE`) y monedas, indicando una adaptación mucho más profunda del modelo base.
- **Chile/España**: el ajuste observado es principalmente monetario y de `PAISORIGEN`, sin reestructuración tributaria equivalente a Argentina.

## E. Integración InfoFact / facturación electrónica (`8. InfoFact.sql`)

### E.1 Objetos observados
- Vistas FE como `vNotaCreditoCabeceraInfoFact` y `vDocumentoCabeceraInfoFact`.
- Procedimientos FE adicionales en `5. SP.sql` con prefijos `usp_FE_`, `usp_ListDocumentosFE`, `spRep_SUNATtxt`, entre otros.

### E.2 Qué agrega funcionalmente
1. **Normalización tributaria de documentos**: transforma `MDOCUMENTO` y `MNOTACREDITO` a estructuras con `TipoDocumento`, `TipoOperacion`, fechas/hora de emisión, moneda y referencias formateadas.
2. **Datos de emisor y receptor**: proyecta datos de `TPARAMETRO`, `TUBIGEO`, `vCliente`, `vTipoIdentidad` y `vTipoDocumento` para construir payloads FE.
3. **Cálculo de bases e impuestos**: separa gravadas, gratuitas, inafectas, exoneradas, IGV, recargos, detracción y retención, incluso contemplando impresión en moneda extranjera.
4. **Soporte de crédito y contingencia**: evalúa `lCajaContingencia`, `tEstadoDocumento='03'`, días de crédito y documentos de referencia.
5. **Respuesta / trazabilidad FE**: expone columnas como `lEstadoFacturacion`, `tRespFacturacion`, `tImprTermica`, tokens y observaciones para cierre del ciclo FE.

### E.3 Riesgos específicos de esta capa
- Está **fuertemente acoplada a Perú/SUNAT** (`PEN`, `PE`, ubigeo peruano, códigos de operación concretos).
- Mezcla reglas tributarias, composición de documento y formato de integración en vistas SQL monolíticas.
- Reusar esta lógica en otro stack requerirá convertir vistas SQL complejas en servicios de dominio y adapters FE por país.
