# Brechas Conocidas — INFOREST

> Inventario de brechas, pendientes y elementos que requieren atención durante la migración.
>
> Status: IN_PROGRESS — baseline .NET 8 de Fase 3 validado; los gaps remanentes se mantienen como backlog controlado para fases posteriores.

---

## Brechas Arquitectónicas

| ID | Descripción | Impacto | Prioridad | Estado |
|---|---|---|---|---|
| GAP-001 | Arquitectura .NET 8 aún no cubre módulos funcionales completos de operación | Alto | Alta | IN_PROGRESS — ADR-001 a ADR-012 aceptados y baseline transversal validado; faltan slices funcionales P3-04+ |
| GAP-002 | No existe todavía un módulo funcional completo listo para producción en .NET 8 | CRÍTICO | Alta | IN_PROGRESS — existe baseline ejecutable con seguridad/KDS/reportes/periféricos, pero ningún ejecutable reemplaza al Legacy end-to-end |
| GAP-003 | 206 reportes Crystal Reports — plan de migración definido, implementación en progreso | Alto | Alta | IN_PROGRESS — Etapa 10 |
| GAP-004 | Hardware POS (PinPad, fiscal, cajón) sin abstracción .NET | Alto | Alta | IN_PROGRESS — P3-11: PinPad (P/Invoke), CashDro y BlueVision implementados. Biometría SecuGen y Fiscal Epson BLOCKED (OCX 32-bit sin SDK .NET). |
| GAP-005 | Integraciones fiscales reales multi-país siguen pendientes pese a la arquitectura base ya definida | Alto | Alta | IN_PROGRESS — P3-11: IPaisPolicy + FacturacionElectronicaFactory + 5 gateways FE implementados como stubs. Integración real con SUNAT/AFIP/SII/SRI/SIAT pendiente. |
| GAP-006 | Base de datos Target no definida | Alto | Alta | OPEN |

---

## Brechas de Seguridad

| ID | Descripción | Archivo Legacy | Impacto | Prioridad |
|---|---|---|---|---|
| SEC-001 | Credenciales SQL hardcodeadas en código fuente | `modPuntoVenta.bas` Sub Main | CRÍTICO | Alta |
| SEC-002 | Cifrado débil XOR+César para datos sensibles | `ClsSeguridad.cls` | Alto | Alta |
| SEC-003 | SQL inline potencialmente vulnerable a injection | Múltiples `*.frm` | Alto | Media |
| SEC-004 | Variables globales sin control de acceso (543 vars) | `modDeclaracion.bas` | Medio | Media |
| SEC-005 | Licencia por dongle físico sin alternativa moderna | `License.cls`, `ModuloHardKey.bas` | Bajo | Baja |
| SEC-006 | `TUSUARIO.tPassword` (`varchar(8)`) no permite persistir BCrypt sin tabla sidecar — **RESUELTO**: tabla sidecar `TUSUARIO_HASH` implementada (ADR-013, script `migrations/001_TUSUARIO_HASH.sql`) | `TUSUARIO` | Alto | Alta — COMPLETED |

---

## Funcionalidades Legacy Sin Equivalente .NET

> Lista de componentes Legacy que siguen sin equivalente .NET completo o productivo.

| Componente Legacy | Tipo | Estado Target |
|---|---|---|
| Todos los 401 formularios VB6 | UI Forms | NOT_STARTED |
| Todos los 32 módulos BAS | Business Logic | NOT_STARTED |
| Todas las 10 clases | Domain Objects | NOT_STARTED |
| Todos los 206 reportes Crystal | Reports | NOT_STARTED |
| Integraciones hardware críticas (SecuGen, Epson fiscal, periféricos productivos) | Hardware | PARTIAL |

---

## Componentes de Propósito Incierto

> Componentes cuyo propósito no es determinable con certeza desde el código existente.

| Componente | Evidencia | Estado |
|---|---|---|
| `modMasticar.bas` | Nombre sugeriere procesamiento de datos, propósito exacto UNKNOWN | UNKNOWN |
| `modPvCorp.bas` | Posiblemente punto de venta corporativo, no confirmado | UNKNOWN |
| `Mapping.cls` | Mapeo de datos, uso exacto no determinable | UNKNOWN |
| `Infhotel.ini` | Integración con sistema hotelero, detalles UNKNOWN | UNKNOWN |

---

## Dependencias de Hardware Sin Plan de Migración

| Hardware | Módulo Legacy | Complejidad Migración | Plan .NET |
|---|---|---|---|
| Impresora térmica (ESC/POS) | TTIPODOCUMENTOIMPRESORA | Media | UNKNOWN |
| Cajón de dinero (RJ11) | API serial | Baja | UNKNOWN |
| CashDro (cajón automático) | `modProcedimientoNuevo.bas` | Media | `ICashDroService` + `CashDroService` (Process.Start). Validado en tests; wiring operativo real pendiente. |
| PinPad DLL3500 | `DLL3500.bas`, `CAJA_PINPAD.dll` | Alta — DLL Win32 | `IPinPadService` + `PinPadService` (P/Invoke). Falta validación con DLL/hardware real. |
| Impresora fiscal Epson | `modImpresoraFiscal.bas`, `IFEpson.ocx` | Alta — OCX Win32 | UNKNOWN |
| Biometría SecuGen | `FpLibX_Const.bas`, `sgfplibx.ocx` | Alta — OCX Win32 | UNKNOWN |
| KDS pantalla cocina | `modKDS.bas` | Media — XML sobre directorio | PARTIAL — existe baseline .NET para XML/directorio, falta pantalla operativa y confirmación con hardware real |
| BlueVision/TVS | `modBlueVision.bas` | Alta — COM Win32 | PARTIAL — `IBlueVisionService` + `BlueVisionHttpClient`; falta validar servidor/protocolo real en ambiente operativo |

---

## Integraciones Fiscales por País Sin Plan

| País | Integración | Módulo Legacy | Plan .NET |
|---|---|---|---|
| Perú | SUNAT FE, RUC, IGV | SP + `scriptPeruAlIniciar.sql` | `PeruFEGateway` stub + `PeruPaisPolicy`; integración OSE/SUNAT real pendiente |
| Chile | SII, IVA | `scriptChileAlIniciar.sql` | `ChileFEGateway` stub + `ChilePaisPolicy` |
| Bolivia | SIN código de control | `CodigoControl.bas` | `BoliviaFEGateway` stub + `BoliviaPaisPolicy`; código de control/fiscalidad real pendiente |
| Ecuador | SRI | `scriptEcuadorAlIniciar.sql` | `EcuadorFEGateway` stub + `EcuadorPaisPolicy` |
| Argentina | AFIP impresora fiscal | `modImpresoraFiscal.bas` | `ArgentinaFEGateway` stub + `ArgentinaPaisPolicy`; impresora fiscal sigue BLOCKED |
| España | IVA España | `scriptEspanaAlIniciar.sql` | `EspanaPaisPolicy`; gateway FE dedicado no implementado |

---

## TODOs Detectados en el Análisis

| Elemento | Descripción | Archivo/Ubicación |
|---|---|---|
| `modMasticar.bas` | Propósito no determinado — requiere revisión | `Modulos/modMasticar.bas` |
| `modPvCorp.bas` | Propósito no determinado — requiere revisión | `Modulos/modPvCorp.bas` |
| `Mapping.cls` | Uso no determinado — requiere revisión | `Clases/Mapping.cls` |
| `INFHOTEL.INI` | Integración hotelera — detalles sin documentar | Configuración runtime |
| Triggers SQL | Triggers eliminados en runtime (`ElminaTriguer`) — existencia real no confirmada | `modProcedimiento.bas` |
| HardKey físico | La validación de presencia/conexión del dongle (`hk.VerificaConexion`, `hk.IniciaConexion`) aún no tiene equivalente .NET | `frmAcceso.frm`, `License.cls` |

---

## Deuda Técnica del Legacy (A Resolver en Target)

| Deuda | Descripción | Archivo Legacy |
|---|---|---|
| Código spagueti | UI + lógica + datos mezclados en formularios | Todos los `.frm` |
| 543 variables globales | Estado global no controlado | `modDeclaracion.bas` |
| Hardcoding generalizado | Strings literales, credenciales, rutas | Múltiples módulos |
| Sin pruebas automatizadas | Cero tests en el Legacy | — |
| Crystal Reports dependiente | Motor de reportes propietario y legado | `Reportes/` |
| Timeout de 600s en SPs | `clsComando` — SP timeout muy alto | `Clases/clsComando.cls` |
| Timers para polling | Timer-based polling (250ms–60000ms) | Múltiples formularios |

---

*Este documento debe actualizarse a medida que se identifican nuevas brechas durante la migración.*

---

## Etapa 9 — Delivery, Despacho, Motorizados

| ID | Gap | Descripción | Archivo Origen | Estado |
|---|---|---|---|---|
| GAP-DEL-001 | sp_UpdFotoDelivery | Actualización de foto del cliente delivery — requiere storage de imágenes no definido | `frmClienteDeliveryDetalle.frm` | BLOCKED |
| GAP-DEL-002 | ADR-009 pendiente | Estrategia de sincronización multi-local: polling actual a CENTRALDELIVERY no escala. Decidir entre polling, Service Bus o SQL replication | `modDespachador.bas` | OPEN |
| GAP-DEL-003 | ADR-010 pendiente | Integración Rappi: actualmente stub. Definir si se mantiene adaptador embebido o se expone endpoint | `frmOrdenesConsola.frm`, `frmEnvioDatosRappi.frm` | OPEN |
| GAP-DEL-004 | RappiOrderAdapter HTTP | `RappiOrderAdapter` retorna lista vacía. El conector HTTP real a la API de Rappi no está implementado | `Inforest.Infrastructure/Delivery/RappiOrderAdapter.cs` | NOT_STARTED |
| GAP-DEL-005 | Reportes FastReport Etapa 9 | 8 reportes del módulo Despachador/Motorizado: RepAnaliticoMotorizado, RepControlMotorizado, RepTiempoDelivery, RepPlanillaMovilidad, RepClienteDelivery, RepPedido, variantes integradas — handlers y repositorio implementados en Etapa 10; plantillas .frx pendientes | `Reportes/` | IN_PROGRESS — Etapa 10 |
| GAP-DEL-006 | Formularios WinForms faltantes | Pendientes: AsignacionMotorizadoForm, ReasignacionMotorizadoForm, ClienteDeliveryForm, PedidoDeliveryForm, OrdenesConsolaForm, BusquedaDeliveryForm | `frmAsignacionMotorizado.frm` etc. | IN_PROGRESS |
