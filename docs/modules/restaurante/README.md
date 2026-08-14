# Módulo Restaurante — INFOREST

> Este documento cubre todos los sub-módulos del sistema INFOREST, ya que el sistema está organizado en un único dominio de restaurante.

---

## Descripción

INFOREST es un sistema de gestión de restaurante. A diferencia de sistemas modulares independientes, sus 7 ejecutables operan sobre la misma base de datos y comparten código. Este documento cubre el sistema completo.

## Responsabilidad

Gestión integral de la operación de restaurantes:
- Toma de pedidos y atención en salón
- Facturación y cobro
- Cocina (comandas, KDS)
- Delivery y despacho a domicilio
- Inventario y almacén
- Administración y configuración
- Reportes gerenciales
- Facturación electrónica multi-país

---

## Legacy

### Ubicación

```
legacy-restaurant/
├── restaurant-vb6/           ← Código fuente VB6
│   ├── Formularios/          ← 401 formularios .frm
│   ├── Modulos/              ← 32 módulos .bas
│   ├── Clases/               ← 10 clases .cls
│   └── Reportes/             ← 206 reportes Crystal Reports
└── database-sql-server/      ← Scripts SQL Server
    ├── 1. Estructura.sql     ← 126 tablas
    ├── 4. Vistas.sql         ← 105 vistas
    ├── 5. SP.sql             ← 105+ stored procedures
    └── opcionales/           ← Scripts por país
```

### Ejecutables (Proyectos VB6)

| Proyecto | EXE | Formularios | Módulos | Clases | Propósito |
|---|---|---|---|---|---|
| `InfoRest.vbp` | `InfoRest.exe` | ~120 | 18 | 10 | Punto de venta principal |
| `CajaRapida.vbp` | `CajaRapida.exe` | ~100 | 16 | 8 | Caja rápida / fast food |
| `Adicion.vbp` | `Adicion.exe` | 37 | 12 | 6 | Adición de comandas |
| `Administracion.vbp` | `Administracion.exe` | 151 | 13 | 7 | Administración y maestros |
| `Consulta.vbp` | `Consulta.exe` | 124 | 15 | 9 | Consultas y reportes |
| `Despachador.vbp` | `Despachador.exe` | 25 | 10 | 3 | Despacho y delivery |
| `Motorizados.vbp` | `Motorizado.exe` | 2 | 7 | 3 | Control de motorizados |

### Formularios Principales

#### Operación y Ventas
| Formulario | Propósito |
|---|---|
| `frmAdicion.frm` | Adición de ítems a pedido |
| `frmAgrupacion.frm` | Agrupación de pedidos |
| `frmCaja.frm` | Gestión de caja |
| `frmCajaRapida.frm` | Caja rápida |
| `frmCajaDetalle.frm` | Detalle de caja |
| `frmCentralPedidos.frm` | Central de pedidos |
| `frmCheffControl.frm` | Control de cocina |
| `frmDelivery` (múltiples) | Operación delivery |
| `frmDespachador.frm` | Despacho |

#### Administración y Maestros
| Formulario | Propósito |
|---|---|
| `frmArea.frm` | Gestión de áreas |
| `frmCargo.frm` | Cargos |
| `frmCentroCosto.frm` | Centros de costo |
| `frmCliente.frm` | Maestro de clientes |
| `frmClienteDelivery.frm` | Clientes delivery |
| `frmChofer.frm` | Choferes/motorizados |
| `frmProducto` (inferido) | Maestro de productos |

#### Reportes (`frmRep*`)
> ~76 formularios de reporte. Ver Reportes Crystal Reports más abajo.

#### Utilidades
| Formulario | Propósito |
|---|---|
| `frmAcceso.frm` | Control de acceso/login |
| `frmBackup.frm` | Respaldo de BD |
| `frmBusca.frm` | Búsqueda genérica |
| `frmCalendario.frm` | Selector de fecha |
| `FrmUpdate.frm` | Actualización del sistema |
| `frmOptimizar.frm` | Optimización de BD |

### Módulos VB6

> Ver inventario completo en [docs/architecture/legacy-architecture.md](../../architecture/legacy-architecture.md)

| Módulo | Rol |
|---|---|
| `modDeclaracion.bas` | 543 variables globales — estado de sesión y configuración |
| `modProcedimiento.bas` | Utilidades núcleo del sistema |
| `modPuntoVenta.bas` | Entry point del POS |
| `modKDS.bas` | Integración Kitchen Display System |
| `modImpresoraFiscal.bas` | Impresora fiscal Epson (Argentina) |
| `modBlueVision.bas` | Display cliente BlueVision |
| ... | Ver lista completa en arquitectura Legacy |

### Clases VB6

| Clase | Rol |
|---|---|
| `clsComando.cls` | Wrapper ADODB.Command para SPs |
| `ClsDocumento.cls` | Operaciones de documentos |
| `clsAlmacen.cls` | Kardex y descargos de inventario |
| `ClsSeguridad.cls` | Cifrado (XOR+César — débil) |
| `clsDiaContable.cls` | Gestión del día contable |

### Stored Procedures (Selección)

| SP | Módulo Funcional | Descripción |
|---|---|---|
| `spIns_MPEDIDO` | Pedidos | Inserción de pedido |
| `spUpd_MPEDIDO` | Pedidos | Actualización de pedido |
| `spRep_RegVenta` | Reportes | Registro de ventas |
| `spRep_Anulacion` | Reportes | Anulaciones |
| `spRep_Descuento` | Reportes | Descuentos |
| `spRep_Propina` | Reportes | Propinas |
| `spRep_VentaFecha` | Reportes | Ventas por fecha |
| `spRep_Ranking` | Reportes | Ranking de productos |
| `spRep_TiempoDelivery` | Delivery | Tiempos de delivery |
| `USP_KDS_ObtenerProductoPedido` | KDS | Productos para KDS |
| `USP_MODIFICARINSUMOS` | Almacén | Gestión de insumos |
| `usp_Aud_ObtenerSalones` | Auditoría | Salones auditados |

> Ver inventario completo: [docs/database/stored-procedures.md](../../database/stored-procedures.md)

### Tablas Principales

| Tabla | Dominio | Descripción |
|---|---|---|
| `MPEDIDO` | Pedidos | Cabecera de pedido |
| `DPEDIDO` | Pedidos | Detalle de pedido |
| `CPEDIDO` | Pedidos | Combos en pedido |
| `APEDIDO` | Pedidos | Ítems anulados (auditoría) |
| `MDOCUMENTO` | Facturación | Documento de venta |
| `DDOCUMENTO` | Facturación | Detalle de documento |
| `MTURNO` | Caja | Turno de caja |
| `TCAJA` | Caja | Configuración de caja |
| `DPAGODOCUMENTO` | Pagos | Pagos por documento |
| `TPRODUCTO` | Maestros | Maestro de productos |
| `TGRUPO` | Maestros | Grupos de productos |
| `TCLIENTE` | Maestros | Maestro de clientes |
| `TUSUARIO` | Seguridad | Usuarios del sistema |
| `TPARAMETRO` | Config | Parámetros del sistema |
| `TDIACONTABLE` | Contable | Día contable activo |

> Ver inventario completo: [docs/database/legacy-database.md](../../database/legacy-database.md)

### Reportes Crystal Reports

> 206 reportes en total organizados por tipo:

| Prefijo | Tipo | Cantidad Aprox. |
|---|---|---|
| `Crt*` | Crystal Reports integrados | ~15 |
| `dsrPaloteo*` | Paloteo (conteo por producto) | ~15 |
| `dsrPedidos*` | Pedidos | ~10 |
| `dsrRep*` | Reportes operativos | ~50 |
| `dsrConsulta*` | Consultas | ~30 |
| Otros | Misceláneos | ~86 |

> Ubicación: `legacy-restaurant/restaurant-vb6/Reportes/`

### Integraciones

| Sistema | Módulo | Tecnología | Propósito |
|---|---|---|---|
| KDS | `modKDS.bas` | XML sobre directorio | Comandas a cocina |
| BlueVision/TVS | `modBlueVision.bas` | COM ActiveX | Display cliente |
| Facturación Electrónica | `ClsDocumento.cls`, SP FE | Múltiples | Comprobantes electrónicos |
| PinPad DLL3500 | `DLL3500.bas` | Win32 DLL | Pago con tarjeta |
| CashDro | `modProcedimientoNuevo.bas` | HTTP API | Cajón automático |
| Rappi | `modDespachador.bas` | SQL SP | Pedidos Rappi |
| FPay/MercadoPago QR | `modProcedimiento.bas` | SP + motor | Pagos QR |
| Impresora Fiscal Epson | `modImpresoraFiscal.bas` | OCX | Argentina |
| Biometría SecuGen | `FpLibX_Const.bas` | OCX | Autenticación |

---

## Nueva Arquitectura

### Proyecto

> **Status: NOT_STARTED** — No existe código .NET 8 en el repositorio.

**Ubicación:** `modern-net8/` (carpeta vacía)

### Namespace

UNKNOWN — pendiente de definición arquitectónica

### Controllers / Services / Domain / Infrastructure

NOT_STARTED — Ver [target-architecture.md](../../architecture/target-architecture.md)

---

## Reglas de Negocio

> Ver documento completo: [docs/migration/business-rules.md](../../migration/business-rules.md)

| ID | Regla | Estado |
|---|---|---|
| BR-001 | Canales de venta con precios diferenciados | NOT_STARTED |
| BR-002 | Productos con hasta 15 flags de impuesto | NOT_STARTED |
| BR-003 | Pedido cabecera/detalle con combos | NOT_STARTED |
| BR-004 | Integración hotelera en pedido | NOT_STARTED |
| BR-005 | Botones de producto dinámicos por canal | NOT_STARTED |
| BR-006 | Configuración de caja con 30+ flags | NOT_STARTED |
| BR-007 | Obligación de cierre de turno | NOT_STARTED |
| BR-008 | Pago antes de impresión | NOT_STARTED |
| BR-009 | Día contable automático | NOT_STARTED |
| BR-010 | XML KDS estándar | NOT_STARTED |
| BR-011 | Routing de ítems a áreas de cocina | NOT_STARTED |
| BR-012 | Administración centralizada multi-local | NOT_STARTED |
| BR-013 | Descargo automático de insumos | NOT_STARTED |
| BR-014 | Validación de licencia por dongle | NOT_STARTED |

---

## Dependencias

### Base de Datos
- SQL Server `INFOREST` — base principal
- SQL Server `INFSEGURIDAD` — seguridad y auditoría
- SQL Server `ALMACEN` — inventario
- SQL Server `CENTRALDELIVERY` — delivery centralizado
- SQL Server `FACTURACION` — facturación electrónica

### Hardware
- Impresoras térmicas (ESC/POS)
- Cajón de dinero
- CashDro (cajón automático)
- PinPad DLL3500
- Impresora fiscal Epson (Argentina)
- Biometría SecuGen
- KDS (pantalla cocina)
- BlueVision/TVS (display cliente)

### Software/COM
- Crystal Reports 6/9
- MSCOMCTL.OCX, MCI32.OCX, MSBIND.DLL
- IFEpson.ocx (impresora fiscal)
- CAJA_PINPAD.dll
- Chilkat Mail
- BlueVision COM ActiveX

### Archivos de Configuración
- `INFOREST.INI` — configuración principal
- `ALMACEN.INI` — almacén
- `FACTURACION.INI` — facturación
- `DLL3500.INI` — PinPad
- `RUTA.INI` — rutas

---

## Estado de Migración

| Aspecto | Estado |
|---|---|
| Análisis Legacy | IN_PROGRESS — README completado, módulos en análisis |
| Reglas de negocio documentadas | PARTIAL — 14 reglas identificadas, pendiente profundización |
| Arquitectura Target | NOT_STARTED |
| Código .NET 8 | NOT_STARTED |
| Pruebas | NOT_STARTED |

---

## Pendientes

- [ ] Analizar detalladamente cada uno de los 401 formularios
- [ ] Documentar reglas de negocio de cada formulario clave
- [ ] Definir arquitectura .NET 8 (ADRs pendientes)
- [ ] Crear proyecto .NET 8 base
- [ ] Definir estrategia para 206 reportes Crystal Reports
- [ ] Definir estrategia para hardware POS
- [ ] Documentar propósito de `modMasticar.bas` y `modPvCorp.bas`
- [ ] Documentar integración hotelera (`Infhotel.ini`)

---

## Riesgos

| Riesgo | Impacto | Probabilidad |
|---|---|---|
| 543 variables globales en modDeclaracion — estado oculto | Alto | Alta |
| Reglas de negocio en formularios UI (no separadas) | Alto | Alta |
| Hardware POS sin soporte directo .NET | Alto | Alta |
| 206 reportes Crystal Reports sin migración | Medio | Alta |
| Credenciales SQL hardcodeadas en código | Crítico | Confirmado |
| Cifrado débil en uso | Alto | Confirmado |

---

## Evidencia

- [legacy-restaurant/README.md](../../../legacy-restaurant/README.md) — CONFIRMED — análisis técnico completo
- [legacy-restaurant/restaurant-vb6/](../../../legacy-restaurant/restaurant-vb6/) — CONFIRMED — código fuente disponible
- [legacy-restaurant/database-sql-server/](../../../legacy-restaurant/database-sql-server/) — CONFIRMED — scripts SQL disponibles
- [modern-net8/README.md](../../../modern-net8/README.md) — CONFIRMED — carpeta vacía, sin implementación
