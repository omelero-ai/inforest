# Matriz de Lógica de Negocio — Legacy VB6 + SQL → .NET 8

| ID | Regla | En VB6 | En SQL | Módulo Legacy | Destino .NET | Estado |
|---|---|---|---|---|---|---|
| BR-001 | Canales de venta con precios diferenciados | `modDeclaracion.bas`, `modProcedimiento.bas` | `TPRODUCTO`, `TCANALVENTA` | Punto de Venta | `Domain/Services/CanalPricingService` | ANALYSIS |
| BR-002 | Productos con flags de impuesto | — | `TPRODUCTO`, `TPARAMETRO` | Maestros / Ventas | `Domain/Entities/Producto` + `TaxPolicy` | ANALYSIS |
| BR-003 | Estado de mesa y disponibilidad | `frmVenta.frm`, `frmMesa.frm`, `frmMantenimientoMesa.frm` | `TMESA`, `MPEDIDO`, `vEstadoMesa` | Punto de Venta | `Domain/Entities/Mesa` | ANALYSIS |
| BR-004 | Turno de caja activo por terminal | `modPuntoVenta.bas`, `frmCaja.frm` | `MTURNO`, `TCAJA` | Caja / Punto de Venta | `Application/TurnoCaja` | ANALYSIS |
| BR-005 | Día contable automático/manual | `clsDiaContable.cls`, `modPuntoVenta.bas` | `TPARAMETRO`, `TDIACONTABLE`, `usp_GenObtieneDiaContable`, `usp_GenInsertarDiaContable` | Caja / Contable | `Domain/Services/DiaContableService` | ANALYSIS |
| BR-006 | Validación de credenciales y permisos | `frmAcceso.frm`, `ClsSeguridad.cls`, `modSeguridadInfhotel.bas` | `TUSUARIO`, `TGRUPOUSUARIO`, `TACCESO`, `TGRUPOACCESO` | Seguridad | `Application/IAuthService`, `Infrastructure/Security/AuthService`, `Infrastructure/Security/RbacService`, `Desktop/Form1` | IN_PROGRESS |
| BR-007 | Cierre de turno y cuadre | `frmCaja.frm`, `frmLiquidacion.frm` | `MTURNO`, `MCIERRE`, `TDIACONTABLE` | Caja | `Application/CierreTurno` | ANALYSIS |
| BR-008 | Descarga de almacén en venta | `clsAlmacen.cls` | `usp_Inforest_DescargoVenta`, `usp_Inforest_DescargoVentaPorInsumo`, `USP_MODIFICARINSUMOS`, `TINSUMO` | Inventario / Almacén | `Infrastructure/InventoryGateway` | ANALYSIS |
| BR-009 | XML KDS por pedido/ítem/condimento | `modKDS.bas`, `frmCheffControl.frm` | `USP_KDS_ObtenerDetallePedidox`, `DPEDIDOKDS` | Cocina / KDS | `Infrastructure/KdsAdapter` | ANALYSIS |
| BR-010 | Routing de ítems a áreas de cocina | `frmPedido.frm`, `frmVenta.frm` | `TPRODUCTOAREA`, `TAREAIMPRESORA`, `DPEDIDO` | Cocina / Ventas | `Domain/Services/AreaProduccionRouter` | ANALYSIS |
| BR-011 | Administración centralizada multi-local | `modPuntoVenta.bas` | `usp_AdmCen_*`, `TLOCAL`, `TTIENDA` | Administración centralizada | `Infrastructure/CentralSync` | ANALYSIS |
| BR-012 | Validación de licencia por hardkey | `License.cls`, `modSeguridadInfhotel.bas` | `THARDKEY`, `TPARAMETROVERSION` (INFSEGURIDAD) | Seguridad / Licencia | `Application/ILicenseService`, `Infrastructure/Security/LicenseService` | IN_PROGRESS |
| BR-013 | Cobro antes de impresión | `frmPago.frm`, `modDeclaracion.bas` | `TCAJA`, `TPARAMETRO`, `DPAGODOCUMENTO` | Caja / Pagos | `Application/PagoDocumento` | ANALYSIS |
| BR-014 | Productos con área y visibilidad por canal | `frmProducto.frm`, `frmProductoBoton.frm` | `TPRODUCTOAREA`, `VISIBILIDADPROPIEDADXCANAL`, `TVISIBILIDADTARJETACREDITOXCANAL` | Maestros / POS | `Domain/Services/ProductoVisibilidadService` | ANALYSIS |
| BR-SQL-001 | spIns_MPEDIDO — inserción de pedido con correlativo, código anual y actualización de TPEDIDOMESA | Invocado vía `clsComando` / formularios de venta | `spIns_MPEDIDO`, `TPARAMETRO`, `TMESA`, `MPEDIDO`, `TPEDIDOMESA` | Pedidos | `Application/Commands/CreatePedidoCommand` | ANALYSIS |
| BR-SQL-002 | spUpd_MPEDIDO — actualización integral de pedido | Invocado vía `clsComando` / formularios de edición | `spUpd_MPEDIDO`, `TMESA`, `MPEDIDO` | Pedidos | `Application/Commands/UpdatePedidoCommand` | ANALYSIS |
| BR-SQL-003 | TPARAMETRO flags — configuración funcional global | `modPuntoVenta.bas` carga parámetros al iniciar | `TPARAMETRO` | Configuración global | `Shared/Configuration/RuntimeFlags` | ANALYSIS |
| BR-SQL-004 | TCAJA flags — configuración operativa por caja | `modPuntoVenta.bas` carga `TCAJA` al iniciar | `TCAJA` | Configuración de caja | `Shared/Configuration/CajaProfile` | ANALYSIS |
