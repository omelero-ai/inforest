# Matriz de Trazabilidad — INFOREST Legacy → .NET 8

> Status: NOT_STARTED — No existe código .NET 8. Esta matriz se completa a medida que avanza la migración.
>
> Última actualización: 2026-08-11

---

## Instrucciones de Uso

Esta matriz relaciona cada componente Legacy con su equivalente en .NET 8.

**Actualizar esta matriz obligatoriamente** cuando:
- Se inicia la migración de un componente (`ANALYSIS` → `IN_PROGRESS`)
- Se completa la implementación (`IN_PROGRESS` → `MIGRATED`)
- Se valida el comportamiento (`MIGRATED` → `COMPLETED`)

**Estados válidos:** `NOT_STARTED` | `ANALYSIS` | `IN_PROGRESS` | `MIGRATED` | `VALIDATING` | `COMPLETED` | `BLOCKED` | `NOT_APPLICABLE` | `UNKNOWN`

---

## Módulos Ejecutables

| Legacy | Tipo | .NET 8 | Tipo | Estado | Evidencia |
|---|---|---|---|---|---|
| `InfoRest.vbp` / `InfoRest.exe` | VBP Project | — | — | NOT_STARTED | — |
| `CajaRapida.vbp` / `CajaRapida.exe` | VBP Project | — | — | NOT_STARTED | — |
| `Adicion.vbp` / `Adicion.exe` | VBP Project | — | — | NOT_STARTED | — |
| `Administracion.vbp` / `Administracion.exe` | VBP Project | — | — | NOT_STARTED | — |
| `Consulta.vbp` / `Consulta.exe` | VBP Project | — | — | NOT_STARTED | — |
| `Despachador.vbp` / `Despachador.exe` | VBP Project | — | — | NOT_STARTED | — |
| `Motorizados.vbp` / `Motorizado.exe` | VBP Project | — | — | NOT_STARTED | — |

---

## Módulos BAS (Lógica de Negocio)

| Legacy | Tipo | .NET 8 | Tipo | Estado | Notas |
|---|---|---|---|---|---|
| `modDeclaracion.bas` | Global State | — | — | NOT_STARTED | 543 variables globales — requiere análisis especial |
| `modProcedimiento.bas` | Core Utils | — | — | NOT_STARTED | Actualizador, FPay, botones, QR |
| `modProcedimientoNuevo.bas` | Extensions | — | — | NOT_STARTED | CashDro |
| `modPuntoVenta.bas` | App Entry | — | — | NOT_STARTED | Sub Main POS |
| `modCajaRapida.bas` | App Entry | — | — | NOT_STARTED | Sub Main caja |
| `modAdicion.bas` | App Entry | — | — | NOT_STARTED | Sub Main adición |
| `modAdministracion.bas` | App Entry | — | — | NOT_STARTED | Sub Main admin |
| `modConsulta.bas` | App Entry | — | — | NOT_STARTED | Sub Main consultas |
| `modConsultaIntregrada.bas` | Multi-local | — | — | NOT_STARTED | Consulta multi-local |
| `modDespachador.bas` | App Entry | — | — | NOT_STARTED | Sub Main despacho |
| `modMotorizado.bas` | App Entry | — | — | NOT_STARTED | Sub Main motorizado |
| `modKDS.bas` | Integration | — | — | NOT_STARTED | XML KDS |
| `modBlueVision.bas` | Integration | — | — | NOT_STARTED | BlueVision TVS |
| `modImpresoraFiscal.bas` | Integration | — | — | NOT_STARTED | Epson fiscal |
| `modAuditoria.bas` | Audit | — | — | NOT_STARTED | Auditoría |
| `modAuditoriaEquipo.bas` | Audit | — | — | NOT_STARTED | Auditoría equipo |
| `modAuditoriaIntegral.bas` | Audit | — | — | NOT_STARTED | INFSEGURIDAD |
| `modBarcode.bas` | Utility | — | — | NOT_STARTED | Códigos de barras |
| `modSeguridadInfhotel.bas` | License | — | — | NOT_STARTED | Licencias |
| `modConexionIp.bas` | Network | — | — | NOT_STARTED | Conectividad |
| `modCrearInis.bas` | Config | — | — | NOT_STARTED | Creación INIs |
| `modTime.bas` | Utility | — | — | NOT_STARTED | Control de tiempo |
| `modMasticar.bas` | Unknown | — | — | UNKNOWN | Propósito no determinado |
| `modCheffControl.bas` | Kitchen | — | — | NOT_STARTED | Chef control |
| `modGuias.bas` | Dispatch | — | — | NOT_STARTED | Guías transporte |
| `modPvCorp.bas` | Unknown | — | — | UNKNOWN | Propósito no determinado |
| `CodigoControl.bas` | Fiscal | — | — | NOT_STARTED | Bolivia código control |
| `DLL3500.bas` | Hardware | — | — | NOT_STARTED | PinPad DLL3500 |
| `FpLibX_Const.bas` | Hardware | — | — | NOT_STARTED | Biometría constantes |
| `ModPictureBoxCustom.bas` | UI Utility | — | — | NOT_STARTED | PictureBox custom |
| `ModuloHardKey.bas` | License | — | — | NOT_STARTED | Dongle |
| `VBZipBas.bas` | Utility | — | — | NOT_STARTED | ZIP |

---

## Clases (Domain Objects)

| Legacy | Tipo | .NET 8 | Tipo | Estado | Notas |
|---|---|---|---|---|---|
| `clsComando.cls` | Data Access | — | — | NOT_STARTED | Wrapper SP ADO |
| `ClsDocumento.cls` | Domain | — | — | NOT_STARTED | Documentos almacén |
| `clsAlmacen.cls` | Domain | — | — | NOT_STARTED | Kardex, descargos |
| `ClsSeguridad.cls` | Security | — | — | NOT_STARTED | Reemplazar con BCrypt |
| `clsDiaContable.cls` | Domain | — | — | NOT_STARTED | Día contable |
| `claCorreoElectronico.cls` | Notification | — | — | NOT_STARTED | Email |
| `clsxml.cls` | Utility | — | — | NOT_STARTED | XML |
| `clsTrama.cls` | Integration | — | — | NOT_STARTED | FE Paperlees |
| `License.cls` | License | — | — | NOT_STARTED | Validación licencia |
| `Mapping.cls` | Mapping | — | — | UNKNOWN | Propósito no determinado |

---

## Tablas de Base de Datos (Selección Principal)

| Legacy (SQL Server) | Tipo | .NET 8 | Tipo | Estado | Notas |
|---|---|---|---|---|---|
| `MPEDIDO` | Table | — | — | NOT_STARTED | Cabecera pedido |
| `DPEDIDO` | Table | — | — | NOT_STARTED | Detalle pedido |
| `CPEDIDO` | Table | — | — | NOT_STARTED | Combos en pedido |
| `APEDIDO` | Table | — | — | NOT_STARTED | Auditoría anulados |
| `MDOCUMENTO` | Table | — | — | NOT_STARTED | Documento de venta |
| `DDOCUMENTO` | Table | — | — | NOT_STARTED | Detalle documento |
| `MTURNO` | Table | — | — | NOT_STARTED | Turno de caja |
| `TCAJA` | Table | — | — | NOT_STARTED | Config caja |
| `TPRODUCTO` | Table | — | — | NOT_STARTED | Maestro productos |
| `TGRUPO` | Table | — | — | NOT_STARTED | Grupos productos |
| `TSUBGRUPO` | Table | — | — | NOT_STARTED | Subgrupos |
| `TCLIENTE` | Table | — | — | NOT_STARTED | Maestro clientes |
| `TUSUARIO` | Table | — | — | NOT_STARTED | Usuarios |
| `TPARAMETRO` | Table | — | — | NOT_STARTED | Config sistema |
| `DPAGODOCUMENTO` | Table | — | — | NOT_STARTED | Pagos |
| `MCIERRE` | Table | — | — | NOT_STARTED | Cierre contable |
| `TDIACONTABLE` | Table | — | — | NOT_STARTED | Día contable |

> Ver inventario completo: [docs/database/legacy-database.md](../database/legacy-database.md)

---

## Stored Procedures Clave

| Legacy SP | Módulo | .NET 8 | Tipo | Estado | Notas |
|---|---|---|---|---|---|
| `spIns_MPEDIDO` | Pedidos | — | — | NOT_STARTED | Inserción pedido |
| `spUpd_MPEDIDO` | Pedidos | — | — | NOT_STARTED | Actualización pedido |
| `spRep_RegVenta` | Reportes | — | — | NOT_STARTED | Reporte ventas |
| `spRep_Anulacion` | Reportes | — | — | NOT_STARTED | Reporte anulaciones |
| `USP_KDS_ObtenerProductoPedido` | KDS | — | — | NOT_STARTED | Productos para KDS |
| `USP_MODIFICARINSUMOS` | Almacén | — | — | NOT_STARTED | Gestión insumos |
| `usp_Aud_ObtenerSalones` | Auditoría | — | — | NOT_STARTED | Auditoría salones |

> Ver inventario completo: [docs/database/stored-procedures.md](../database/stored-procedures.md)

---

## Reportes Crystal Reports

> 206 reportes en total. Ver inventario en [docs/modules/restaurante/README.md](../modules/restaurante/README.md)

| Legacy Reporte | Tipo | .NET 8 | Tipo | Estado |
|---|---|---|---|---|
| `CrtRepInventarioIntegrado.dsr` | Crystal | — | — | NOT_STARTED |
| `CrtRegComprasIntegrado.dsr` | Crystal | — | — | NOT_STARTED |
| `dsrPedidos*.dsr` | Crystal | — | — | NOT_STARTED |
| `dsrPaloteoVenta*.dsr` | Crystal | — | — | NOT_STARTED |
| *(206 en total)* | Crystal | — | — | NOT_STARTED |

---

## Integraciones Externas

| Legacy | Tecnología | .NET 8 | Tipo | Estado |
|---|---|---|---|---|
| KDS (Kitchen Display) | XML sobre directorio | — | — | NOT_STARTED |
| BlueVision/TVS | COM ActiveX | — | — | NOT_STARTED |
| Facturación Electrónica | Múltiples SDKs | — | — | NOT_STARTED |
| PinPad DLL3500 | Win32 DLL | — | — | NOT_STARTED |
| CashDro | HTTP API | — | — | NOT_STARTED |
| Rappi | SQL SP | — | — | NOT_STARTED |
| FPay/MercadoPago QR | SP + motor | — | — | NOT_STARTED |
| Biometría SecuGen | OCX Win32 | — | — | NOT_STARTED |
| Impresora Fiscal Epson | OCX Win32 | — | — | NOT_STARTED |
| Chilkat Email | COM | — | — | NOT_STARTED |

---

*Esta matriz debe actualizarse con cada avance de migración.*
