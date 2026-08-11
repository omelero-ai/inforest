# Brechas Conocidas — INFOREST

> Inventario de brechas, pendientes y elementos que requieren atención durante la migración.
>
> Status: PARTIAL — análisis basado en el código Legacy. Se actualizará a medida que avance la migración.

---

## Brechas Arquitectónicas

| ID | Descripción | Impacto | Prioridad | Estado |
|---|---|---|---|---|
| GAP-001 | No existe arquitectura .NET 8 definida | CRÍTICO — bloquea toda migración | Alta | OPEN |
| GAP-002 | No existe código .NET 8 en el repositorio | CRÍTICO | Alta | OPEN |
| GAP-003 | 206 reportes Crystal Reports sin plan de migración | Alto | Alta | OPEN |
| GAP-004 | Hardware POS (PinPad, fiscal, cajón) sin abstracción .NET | Alto | Alta | OPEN |
| GAP-005 | Multi-país (impuestos, FE) sin arquitectura definida | Alto | Alta | OPEN |
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

---

## Funcionalidades Legacy Sin Equivalente .NET

> Lista de componentes Legacy que NO tienen equivalente en el Target (porque no existe código .NET 8).

| Componente Legacy | Tipo | Estado Target |
|---|---|---|
| Todos los 401 formularios VB6 | UI Forms | NOT_STARTED |
| Todos los 32 módulos BAS | Business Logic | NOT_STARTED |
| Todas las 10 clases | Domain Objects | NOT_STARTED |
| Todos los 206 reportes Crystal | Reports | NOT_STARTED |
| Todas las integraciones hardware | Hardware | NOT_STARTED |

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
| CashDro (cajón automático) | `modProcedimientoNuevo.bas` | Media | UNKNOWN |
| PinPad DLL3500 | `DLL3500.bas`, `CAJA_PINPAD.dll` | Alta — DLL Win32 | UNKNOWN |
| Impresora fiscal Epson | `modImpresoraFiscal.bas`, `IFEpson.ocx` | Alta — OCX Win32 | UNKNOWN |
| Biometría SecuGen | `FpLibX_Const.bas`, `sgfplibx.ocx` | Alta — OCX Win32 | UNKNOWN |
| KDS pantalla cocina | `modKDS.bas` | Media — XML sobre directorio | UNKNOWN |
| BlueVision/TVS | `modBlueVision.bas` | Alta — COM Win32 | UNKNOWN |

---

## Integraciones Fiscales por País Sin Plan

| País | Integración | Módulo Legacy | Plan .NET |
|---|---|---|---|
| Perú | SUNAT FE, RUC, IGV | SP + `scriptPeruAlIniciar.sql` | UNKNOWN |
| Chile | SII, IVA | `scriptChileAlIniciar.sql` | UNKNOWN |
| Bolivia | SIN código de control | `CodigoControl.bas` | UNKNOWN |
| Ecuador | SRI | `scriptEcuadorAlIniciar.sql` | UNKNOWN |
| Argentina | AFIP impresora fiscal | `modImpresoraFiscal.bas` | UNKNOWN |
| España | IVA España | `scriptEspanaAlIniciar.sql` | UNKNOWN |

---

## TODOs Detectados en el Análisis

| Elemento | Descripción | Archivo/Ubicación |
|---|---|---|
| `modMasticar.bas` | Propósito no determinado — requiere revisión | `Modulos/modMasticar.bas` |
| `modPvCorp.bas` | Propósito no determinado — requiere revisión | `Modulos/modPvCorp.bas` |
| `Mapping.cls` | Uso no determinado — requiere revisión | `Clases/Mapping.cls` |
| `INFHOTEL.INI` | Integración hotelera — detalles sin documentar | Configuración runtime |
| Triggers SQL | Triggers eliminados en runtime (`ElminaTriguer`) — existencia real no confirmada | `modProcedimiento.bas` |

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
