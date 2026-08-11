# Contexto del Sistema — INFOREST

> Status: CONFIRMED (basado en análisis del repositorio)

---

## Descripción del Sistema

INFOREST es un sistema ERP gastronómico que gestiona la operación completa de establecimientos de restauración. Funciona como una plataforma multi-ejecutable donde distintos operadores interactúan simultáneamente con el mismo sistema.

---

## Diagrama de Contexto (C4 — Nivel 1)

```mermaid
C4Context
    title Sistema INFOREST — Contexto

    Person(mozo, "Mozo/Mesero", "Toma pedidos en salón y los envía a cocina")
    Person(cajero, "Cajero", "Gestiona cobros, cierre de turno, documentos de venta")
    Person(admin, "Administrador", "Configura el sistema, gestiona productos, usuarios, reportes")
    Person(motorizado, "Motorizado", "Gestiona entregas de delivery")
    Person(chef, "Jefe de Cocina", "Visualiza y gestiona comandas en cocina")
    Person(gerente, "Gerente", "Consulta reportes de ventas y operaciones")

    System(inforest, "INFOREST", "Sistema ERP Gastronómico\nVB6 + SQL Server\nGestiona: ventas, caja, delivery, cocina, inventario, reportes")

    System_Ext(kds, "KDS", "Kitchen Display System\nPantalla de cocina")
    System_Ext(bluevision, "BlueVision/TVS", "Pantalla display para clientes")
    System_Ext(fe, "Facturación Electrónica", "SUNAT, SRI, SIN, AFIP, SII\nsegún país")
    System_Ext(pinpad, "Terminal PinPad", "Pago con tarjeta\nDLL3500")
    System_Ext(cashdro, "CashDro", "Cajón automático de efectivo")
    System_Ext(sqlserver, "SQL Server", "INFOREST\nINFSEGURIDAD\nALMACEN\nCENTRALDELIVERY\nFACTURACION")
    System_Ext(rappi, "Rappi / Delivery Apps", "Pedidos externos de plataformas delivery")
    System_Ext(fpay, "FPay / MercadoPago", "Pagos QR y pasarelas de pago")

    Rel(mozo, inforest, "Toma pedidos, gestiona mesas")
    Rel(cajero, inforest, "Cobra, cierra turno, emite documentos")
    Rel(admin, inforest, "Configura, administra")
    Rel(motorizado, inforest, "Gestiona entregas")
    Rel(chef, inforest, "Visualiza comandas")
    Rel(gerente, inforest, "Consulta reportes")

    Rel(inforest, kds, "Envía comandas XML")
    Rel(inforest, bluevision, "Envía ticket digital")
    Rel(inforest, fe, "Emite comprobantes electrónicos")
    Rel(inforest, pinpad, "Procesa pago con tarjeta")
    Rel(inforest, cashdro, "Controla cajón automático")
    Rel(inforest, sqlserver, "Persiste datos")
    Rel(inforest, rappi, "Recibe pedidos externos")
    Rel(inforest, fpay, "Procesa pagos QR")
```

---

## Usuarios del Sistema

| Actor | Ejecutable Principal | Funciones Principales |
|---|---|---|
| Mozo/Mesero | `InfoRest.exe` | Toma de pedidos, gestión de mesas, envío a cocina |
| Cajero | `InfoRest.exe`, `CajaRapida.exe` | Cobro, emisión de documentos, cierre de turno |
| Administrador | `Administracion.exe` | Configuración, maestros, usuarios, reportes administrativos |
| Operador de Adición | `Adicion.exe` | Agregado de ítems a pedidos en curso |
| Despachador | `Despachador.exe` | Central de pedidos delivery, asignación a motorizados |
| Motorizado | `Motorizado.exe` | Seguimiento de entregas a domicilio |
| Jefe de Cocina | `InfoRest.exe` (frmCheffControl) | Visualización y control de comandas |
| Gerente/Consultor | `Consulta.exe` | Consultas e informes de gestión |

---

## Sistemas Externos Integrados

| Sistema | Tipo | Tecnología Legacy | Módulo VB6 | Estado |
|---|---|---|---|---|
| KDS | Pantalla cocina | XML sobre directorio compartido | `modKDS.bas` | CONFIRMED |
| BlueVision/TVS | Display cliente | COM `BlueVision_Core_TVS` | `modBlueVision.bas` | CONFIRMED |
| Facturación Electrónica | Fiscal | Múltiples (por país) | `ClsDocumento.cls`, SP FE | CONFIRMED |
| PinPad DLL3500 | Pago tarjeta | `CAJA_PINPAD.dll` | `DLL3500.bas` | CONFIRMED |
| CashDro | Cajón automático | API HTTP (timer 2s) | `modProcedimientoNuevo.bas` | CONFIRMED |
| Rappi | Delivery externo | SP dedicado | `modDespachador.bas` | CONFIRMED |
| FPay / MercadoPago QR | Pago QR | `usp_EjecutaMotorServiciosFPAY` | `modProcedimiento.bas` | CONFIRMED |
| Impresora Fiscal Epson | FE Argentina | `ifepson.ocx` | `modImpresoraFiscal.bas` | CONFIRMED |
| Biometría SecuGen | Seguridad | `sgfplibx.ocx` | `FpLibX_Const.bas` | CONFIRMED |
| Chilkat Mail | Email | Librería Chilkat | `claCorreoElectronico.cls` | CONFIRMED |
| APIs RUC/DNI | Consulta fiscal | HTTP | `FrmConsultaRUC.frm` | CONFIRMED |

---

## Restricciones del Contexto

| Restricción | Descripción | Estado |
|---|---|---|
| Multi-país | Scripts SQL y lógica específica por país | CONFIRMED |
| Multi-local | Configuración `AdministracionCentralizada=ON` en INI | CONFIRMED |
| Hardware dependiente | Impresoras térmicas, cajones, PinPads, biometría | CONFIRMED |
| Windows only (Legacy) | COM/ActiveX requieren Windows | CONFIRMED |
| Licencia por hardware | Dongle físico validado por `License.cls` | CONFIRMED |
| Credenciales embebidas | Credenciales SQL en código fuente VB6 | CONFIRMED — riesgo de seguridad |

---

## Referencias

- [Análisis técnico del Legacy](../../legacy-restaurant/README.md)
- [Arquitectura Legacy](legacy-architecture.md)
- [Decisiones arquitectónicas](architecture-decisions.md)
