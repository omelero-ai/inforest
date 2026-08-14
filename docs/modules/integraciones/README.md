# Módulo Integraciones — Periféricos Especializados y FE por País

> Etapa P3-11 — Estado: `IN_PROGRESS`
>
> Última actualización: 2026-08-12

---

## Resumen

Este módulo cubre la integración de periféricos especializados de POS y la Facturación Electrónica (FE) por país en la migración INFOREST VB6 → .NET 8.

---

## Periféricos integrados

### PinPad (cobro con tarjeta)

| Ítem | Valor |
|---|---|
| Legacy | `DLL3500.bas`, `CAJA_PINPAD.dll` |
| Regla | BR-PERIPH-001 |
| ADR | ADR-009 |
| Interfaz .NET | `IPinPadService` |
| Implementación real | `PinPadService` (P/Invoke a `caja_pinpad.dll`) |
| Stub | `NullPinPadService` |
| Habilitación | `appsettings.json` → `Hardware.PinPad.Habilitado` |
| Operaciones | VENTA, VENTA_CASH, VENTA_CUOTAS, ANULACION, CIERRE, SIMULACION_CUOTAS, TEST |
| Estado | IN_PROGRESS |

**Flujo de operación:**
1. `fiOpenPort(configPath)` → abre comunicación con el terminal
2. `fiStartOperation(tipoOp, timeout, &respuesta)` → ejecuta operación
3. Parsear respuesta (`RET_OK=0`, `RET_NOK=-1`, `RET_RUNNING=1`, errores `-2` a `-9`)
4. `fiClosePort()` → cierra puerto

---

### CashDro (cajón automático de dinero)

| Ítem | Valor |
|---|---|
| Legacy | `modProcedimientoNuevo.bas`, `MotorCashDrow.exe` |
| Regla | BR-PERIPH-002 |
| ADR | ADR-011 |
| Interfaz .NET | `ICashDroService` |
| Implementación real | `CashDroService` (`Process.Start` a `MotorCashDrow.exe`) |
| Stub | `NullCashDroService` |
| Habilitación | `appsettings.json` → `Hardware.CashDro.Habilitado` |
| Estado | IN_PROGRESS |

**Notas:** La integración es por proceso externo, no DLL. Si el ejecutable no existe se lanza `InfrastructureException("CASHDRO_NOT_FOUND")`.

---

### BlueVision TVS (visualización en pantallas de mesa)

| Ítem | Valor |
|---|---|
| Legacy | `modBlueVision.bas`, `BlueVision_Core_TVS.dll` (COM) |
| Regla | BR-PERIPH-003 |
| ADR | ADR-012 |
| Interfaz .NET | `IBlueVisionService` |
| Implementación real | `BlueVisionHttpClient` (HttpClient nativo, no COM) |
| Stub | `NullBlueVisionService` |
| Habilitación | `appsettings.json` → `Hardware.BlueVision.Habilitado` |
| Credenciales | `Hardware.BlueVision.Login / ClearPassword / Url` (no BLUEVISION.INI) |
| Estado | IN_PROGRESS |

**Flujo:** POST `/api/session` → POST `/api/ticket` → POST `/api/ticket/line` × N ítems

---

### Biometría SecuGen

| Ítem | Valor |
|---|---|
| Legacy | `FpLibX_Const.bas`, `sgfplibx.ocx` (OCX 32-bit) |
| Regla | BR-PERIPH-004 |
| ADR | ADR-009 |
| Interfaz .NET | `IBiometriaService` |
| Implementación real | **GAP** — `sgfplibx.ocx` es OCX de 32 bits sin SDK .NET oficial |
| Stub activo | `NullBiometriaService` (siempre exitoso) |
| Habilitación | `appsettings.json` → `Hardware.Biometria.Habilitado` |
| Estado | BLOCKED — GAP-004 |

**Gap:** El OCX `sgfplibx.ocx` no tiene SDK .NET oficial conocido. Alternativas: proceso host 32-bit con IPC, o migración a SecuGen SDK moderno cuando se reemplace hardware.

---

### Impresora Fiscal Epson (Argentina)

| Ítem | Valor |
|---|---|
| Legacy | `modImpresoraFiscal.bas`, `IFEpson.ocx` (OCX 32-bit, Argentina) |
| Regla | BR-PERIPH-005 |
| ADR | ADR-009 |
| Interfaz .NET | `IImpresoraFiscalService` |
| Implementación real | **GAP** — `IFEpson.ocx` es OCX de 32 bits sin SDK .NET oficial |
| Stub activo | `NullImpresoraFiscalService` |
| Habilitación | `appsettings.json` → `Hardware.ImpresoraFiscal.Habilitado` |
| Solo para | Argentina (`CodigoPais = "AR"`, `IPaisPolicy.RequiereFiscal = true`) |
| Estado | BLOCKED — GAP-004 |

**Gap:** Alternativa a evaluar: protocolo serial RS-232/USB Epson Fiscal directo o proceso host 32-bit.

---

## Facturación Electrónica por País

### Arquitectura

```
Application Layer:
  IFacturacionElectronicaGateway (contrato polimórfico por país)

Infrastructure Layer:
  FacturacionElectronicaFactory  ← resuelve gateway por CodigoPais
  ├── PeruFEGateway      (SUNAT/OSE — UBL 2.1)
  ├── ArgentinaFEGateway (AFIP WS-FE — CAE)
  ├── ChileFEGateway     (SII DTE)
  ├── EcuadorFEGateway   (SRI RIDE)
  ├── BoliviaFEGateway   (SIAT/NIT)
  └── NullFEGateway      (stub/default)
```

### Estado por país

| País | Gateway | Autoridad Fiscal | Estándar XML | Estado Implementación |
|---|---|---|---|---|
| Perú | `PeruFEGateway` | SUNAT / OSE | UBL 2.1 | STUB — integración real pendiente |
| Argentina | `ArgentinaFEGateway` | AFIP WS-FE / WSAA | XML AFIP | STUB — integración real pendiente |
| Chile | `ChileFEGateway` | SII | DTE XML | STUB — integración real pendiente |
| Ecuador | `EcuadorFEGateway` | SRI | RIDE XML | STUB — integración real pendiente |
| Bolivia | `BoliviaFEGateway` | SIAT | XML Bolivia | STUB — integración real pendiente |
| España | — | AEAT | — | NOT_APPLICABLE — FE no requerida |

### Habilitación

```json
{
  "FE": {
    "Habilitado": true,
    "QRHabilitado": true,
    "RutaXml": "fe/xml/"
  }
}
```

### DocumentoFE (dominio)

Modelo de dominio en `Inforest.Domain/Entities/Hardware/DocumentoFE.cs`:
- Estados: `Pendiente → Enviado → Autorizado / Rechazado / Anulado`
- Método `RegistrarAutorizacion(numeroCAE, qrData, xml)` al recibir respuesta exitosa del proveedor
- Método `RegistrarRechazo(motivo)` al recibir rechazo

---

## Políticas de País

### Arquitectura

```
Application Layer:
  IPaisPolicy (interfaz — CodigoPais, RequiereFiscal, RequiereFE, LabelDocTrib, etc.)

Infrastructure Layer:
  PaisPolicyFactory  ← resuelve política por CodigoPais (default = Perú)
  ├── PeruPaisPolicy
  ├── ArgentinaPaisPolicy
  ├── ChilePaisPolicy
  ├── EcuadorPaisPolicy
  ├── BoliviaPaisPolicy
  └── EspanaPaisPolicy
```

### Tabla de políticas

| CodigoPais | LabelDocTrib | TipoFE | Moneda | RequiereFiscal | RequiereFE |
|---|---|---|---|---|---|
| PE | R.U.C. | SUNAT | PEN | No | Sí |
| AR | C.U.I.T. | AFIP | ARS | **Sí** | Sí |
| CL | R.U.T. | SII | CLP | No | Sí |
| EC | R.U.C. | SRI | USD | No | Sí |
| BO | N.I.T. | SIAT | BOB | No | Sí |
| ES | N.I.F. | AEAT | EUR | No | No |

---

## Tablas Legacy Relacionadas

| Tabla | Base | Propósito |
|---|---|---|
| `TCONFIGURAPERIFERICO` | INFOREST | Configuración de periférico por caja |
| `TCAJATERMINAL` | INFOREST | Terminales físicas (PinPad, caja) |
| `TTERMINAL` | INFOREST | Maestro de terminales |
| `TESTADOBIZLINK` | INFOREST | Estado de integración BizLink/PinPad |
| `TINTEGRACIONES` | INFOREST | Integraciones activas por local |
| `TUSUARIO` | INFSEGURIDAD | Biometría almacenada por usuario (campo biométrico) |
| `TPARAMETRO` | INFOREST | `lFacturacionE`, `TipoFacturacion`, `pais` — flags de FE y país |
| `TCAJA` | INFOREST | `lFacturacionE` por caja, flags de periféricos |
| Tablas de `FACTURACION` | FACTURACION | Comprobantes electrónicos emitidos (base auxiliar Legacy) |

---

## Gaps Conocidos

| GAP | Descripción | Impacto | Estado |
|---|---|---|---|
| GAP-004 (parcial) | Biometría SecuGen — OCX 32-bit sin SDK .NET | Biometría deshabilitada hasta resolución | OPEN |
| GAP-004 (parcial) | Impresora Fiscal Epson — OCX 32-bit sin SDK .NET | Argentina sin fiscal hasta resolución | OPEN |
| GAP-005 (parcial) | Gateways FE Perú, Chile, Ecuador, Bolivia — solo stubs | FE no operativa hasta completar integración | IN_PROGRESS |
| — | Gateways FE — credenciales y endpoints reales pendientes por país | Requiere configuración por cliente | OPEN |

---

## Referencias

- [ADR-009](../architecture/architecture-decisions.md#adr-009) — Interop Win32
- [ADR-010](../architecture/architecture-decisions.md#adr-010) — FE por país
- [ADR-011](../architecture/architecture-decisions.md#adr-011) — CashDro
- [ADR-012](../architecture/architecture-decisions.md#adr-012) — BlueVision
- [BR-PERIPH-001 a BR-PERIPH-007](../migration/business-rules.md)
- [Traceability Matrix](../migration/traceability-matrix.md)
