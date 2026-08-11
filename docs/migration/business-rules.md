# Reglas de Negocio — INFOREST

> Reglas de negocio extraídas del sistema Legacy VB6.
>
> Fuente: Análisis de código en `legacy-restaurant/restaurant-vb6/` y scripts SQL en `legacy-restaurant/database-sql-server/`.
>
> Status: PARTIAL — análisis inicial. Se añaden reglas a medida que se profundiza el análisis por módulo.

---

## Formato de Regla

```
BR-XXX
Nombre:
Origen: Legacy/<archivo>
Archivo: legacy-restaurant/restaurant-vb6/<ruta>
Procedimiento/Función:
Descripción:
Condición:
Resultado:
Excepciones:
Destino .NET:
Estado: NOT_STARTED | IN_PROGRESS | MIGRATED | COMPLETED
Evidencia: CONFIRMED | PARTIAL | UNKNOWN
```

---

## Ventas y Pedidos

### BR-001
**Nombre:** Canales de venta con precios diferenciados

**Origen:** Legacy/modDeclaracion.bas + TPRODUCTO

**Archivo:** `legacy-restaurant/restaurant-vb6/Modulos/modDeclaracion.bas`

**Descripción:** El sistema soporta hasta 5 canales de venta, cada uno con precio independiente por producto.

**Condición:** Canal de venta = 01 (Local), 02 (Delivery), 03 (Llevar), 04 (Canal4), 05 (Canal5)

**Resultado:** Se aplica el precio correspondiente al canal: `nPrecioVenta`, `nPrecioDelivery`, `nPrecioLlevar`, `nPrecioCanal4`, `nPrecioCanal5` de `TPRODUCTO`

**Excepciones:** Si el precio del canal es 0, puede aplicarse el precio base

**Destino .NET:** Servicio de precios por canal

**Estado:** NOT_STARTED

**Evidencia:** CONFIRMED

---

### BR-002
**Nombre:** Productos con hasta 15 flags de impuesto

**Origen:** Legacy/TPRODUCTO

**Archivo:** `legacy-restaurant/database-sql-server/1. Estructura.sql`

**Descripción:** Cada producto puede tener hasta 15 flags de impuesto (`lImpuesto1..15`) para manejar flexibilidad tributaria extrema.

**Condición:** El flag `lImpuestoN` = true indica que el impuesto N aplica al producto

**Resultado:** Se calculan los impuestos activos según TPARAMETRO (nombre, porcentaje de cada impuesto)

**Excepciones:** Si el flag está inactivo, no se aplica el impuesto aunque esté configurado en TPARAMETRO

**Destino .NET:** Servicio de cálculo de impuestos

**Estado:** NOT_STARTED

**Evidencia:** CONFIRMED

---

### BR-003
**Nombre:** Pedido cabecera/detalle con soporte de combos

**Origen:** Legacy/MPEDIDO, DPEDIDO, CPEDIDO

**Archivo:** `legacy-restaurant/database-sql-server/1. Estructura.sql`

**Descripción:** Un pedido tiene cabecera (`MPEDIDO`) con múltiples ítems (`DPEDIDO`). Los combos tienen su propio detalle (`CPEDIDO`) referenciando al ítem padre.

**Condición:** Si el producto es un combo, se registran los sub-productos en CPEDIDO

**Resultado:** El pedido queda registrado en MPEDIDO + DPEDIDO + CPEDIDO (si aplica)

**Excepciones:** Los ítems anulados van a APEDIDO (auditoría)

**Destino .NET:** Entidades Pedido, DetallePedido, DetalleCombo

**Estado:** NOT_STARTED

**Evidencia:** CONFIRMED

---

### BR-004
**Nombre:** Registro de pedido con campos de integración hotelera

**Origen:** Legacy/MPEDIDO

**Archivo:** `legacy-restaurant/database-sql-server/1. Estructura.sql`

**Descripción:** El pedido en MPEDIDO incluye campos para integración con sistemas hoteleros: `tHabitacion`, `tReserva`, `tPasajero`, `tCompania`, `tContacto`.

**Condición:** Cuando el sistema opera integrado con un sistema hotelero

**Resultado:** El pedido queda vinculado a una habitación/reserva del hotel

**Excepciones:** Si no hay integración hotelera, estos campos quedan vacíos

**Destino .NET:** UNKNOWN — requiere análisis de módulo hotelero

**Estado:** NOT_STARTED

**Evidencia:** CONFIRMED (estructura de tabla) / PARTIAL (lógica de uso)

---

### BR-005
**Nombre:** Botones de producto cargados dinámicamente por canal de venta

**Origen:** Legacy/modProcedimiento.bas

**Archivo:** `legacy-restaurant/restaurant-vb6/Modulos/modProcedimiento.bas`

**Descripción:** Los botones de producto en la pantalla de venta se asignan dinámicamente desde `TPRODUCTO.nBoton` y `nBotonRapido`, filtrados por canal de venta y unidad de negocio.

**Condición:** Al iniciar el módulo de venta o al cambiar de canal

**Resultado:** La pantalla muestra los productos del canal activo en su posición de botón configurada

**Excepciones:** UNKNOWN — sin análisis detallado del formulario

**Destino .NET:** Servicio de configuración de menú/pantalla POS

**Estado:** NOT_STARTED

**Evidencia:** CONFIRMED

---

## Caja y Pagos

### BR-006
**Nombre:** Configuración de comportamiento por caja (30+ flags)

**Origen:** Legacy/TCAJA

**Archivo:** `legacy-restaurant/database-sql-server/1. Estructura.sql`

**Descripción:** Cada caja tiene 30+ flags de comportamiento en `TCAJA` que controlan: si requiere comanda, si requiere contraseña, tipos de consumo permitidos, si requiere precuenta, si permite adición, filtros de tipo de pedido, etc.

**Condición:** Los flags se cargan al iniciar el ejecutable

**Resultado:** El comportamiento del POS varía por caja según los flags configurados

**Excepciones:** PARTIAL — requiere análisis detallado de cada flag

**Destino .NET:** Configuración de caja por instancia

**Estado:** NOT_STARTED

**Evidencia:** CONFIRMED

---

### BR-007
**Nombre:** Obligación de cierre de turno antes de operar

**Origen:** Legacy/TCAJA.lObligaCierre

**Archivo:** `legacy-restaurant/database-sql-server/1. Estructura.sql`

**Descripción:** Si `lObligaCierre = true` en TCAJA, el operador debe cerrar el turno antes de poder iniciar operaciones.

**Condición:** `TCAJA.lObligaCierre = 1`

**Resultado:** El sistema bloquea operaciones hasta que se cierre el turno anterior

**Excepciones:** UNKNOWN

**Destino .NET:** Middleware de validación de turno

**Estado:** NOT_STARTED

**Evidencia:** CONFIRMED

---

### BR-008
**Nombre:** Pago antes de impresión

**Origen:** Legacy/TCAJA.lPagoAntesImpresion (inferido)

**Archivo:** `legacy-restaurant/restaurant-vb6/Modulos/modDeclaracion.bas`

**Descripción:** Existe un flag que exige registrar el pago antes de imprimir el comprobante.

**Condición:** Flag `lPagoAntesImpresion` activo

**Resultado:** No se imprime comprobante hasta confirmar el pago

**Excepciones:** UNKNOWN

**Destino .NET:** Servicio de flujo de cobro

**Estado:** NOT_STARTED

**Evidencia:** PARTIAL

---

### BR-009
**Nombre:** Día contable automático o manual con hora de corte

**Origen:** Legacy/modPuntoVenta.bas + TDIACONTABLE

**Archivo:** `legacy-restaurant/restaurant-vb6/Modulos/modPuntoVenta.bas`

**Descripción:** El día contable puede avanzar automáticamente a una hora de corte configurada en `TPARAMETRO`, o puede ser manual.

**Condición:** Si el parámetro de hora de corte está activo y se supera la hora configurada

**Resultado:** El día contable se avanza automáticamente

**Excepciones:** UNKNOWN — comportamiento exacto no determinado con certeza

**Destino .NET:** Servicio de gestión de día contable

**Estado:** NOT_STARTED

**Evidencia:** CONFIRMED (estructura) / PARTIAL (lógica exacta)

---

## Cocina / KDS

### BR-010
**Nombre:** XML KDS con estructura estándar por pedido/ítem/condimento

**Origen:** Legacy/modKDS.bas

**Archivo:** `legacy-restaurant/restaurant-vb6/Modulos/modKDS.bas`

**Descripción:** El sistema genera XMLs para KDS con estructura: `<Transaction><Order>` con campos de terminal, tipo de transacción, estado, tipo de pedido, mesero, mesa, usuario, ítems con nombre/cantidad/estación KDS y condimentos.

**Condición:** Al registrar/modificar un pedido que tiene ítems con `<KDSStation>` asignado

**Resultado:** XML escrito en directorio configurado (`sOrderInfo` del INI); la estación KDS lo lee

**Excepciones:** En caso de combo, el nombre del ítem incluye iniciales del combo como prefijo

**Destino .NET:** Servicio KDS — generación y envío de órdenes

**Estado:** NOT_STARTED

**Evidencia:** CONFIRMED

---

### BR-011
**Nombre:** Dirección de ítems a áreas de cocina

**Origen:** Legacy/TPRODUCTOAREA + TAREAIMPRESORA

**Archivo:** `legacy-restaurant/database-sql-server/1. Estructura.sql`

**Descripción:** Cada producto tiene áreas de impresión/cocina asignadas (`TPRODUCTOAREA`). Al registrar el ítem en el pedido, se envía la comanda a las áreas correspondientes.

**Condición:** El producto tiene área de cocina asignada en `TPRODUCTOAREA`

**Resultado:** La comanda se imprime/envía al área de cocina configurada

**Excepciones:** Si `lImprimeArea = false`, no se envía al área

**Destino .NET:** Servicio de comandas y routing de cocina

**Estado:** NOT_STARTED

**Evidencia:** CONFIRMED

---

## Administración Centralizada

### BR-012
**Nombre:** Administración centralizada multi-local

**Origen:** Legacy/modPuntoVenta.bas + INFOREST.INI

**Archivo:** `legacy-restaurant/restaurant-vb6/Modulos/modPuntoVenta.bas`

**Descripción:** Si `INFOREST.INI [AdministracionCentralizada] CENTRALIZADA=ON`, el sistema consulta maestros (productos, parámetros) desde un servidor central.

**Condición:** Flag `CENTRALIZADA=ON` en INI

**Resultado:** Los maestros de productos y configuración se obtienen del servidor central, no del local

**Excepciones:** UNKNOWN — comportamiento exacto en caso de desconexión del central

**Destino .NET:** Multi-tenant o administración centralizada

**Estado:** NOT_STARTED

**Evidencia:** CONFIRMED

---

## Inventario / Almacén

### BR-013
**Nombre:** Descargo automático de insumos por venta

**Origen:** Legacy/clsAlmacen.cls

**Archivo:** `legacy-restaurant/restaurant-vb6/Clases/clsAlmacen.cls`

**Descripción:** Al completar una venta, el sistema descarga automáticamente los insumos (`TINSUMO`) asociados a los productos vendidos mediante SP dedicados.

**Condición:** El producto tiene insumos configurados en `TINSUMO`

**Resultado:** Stock de insumos reducido en la BD Almacén

**Excepciones:** PARTIAL — condiciones exactas de descargo no completamente documentadas

**Destino .NET:** Servicio de inventario / descargo automático

**Estado:** NOT_STARTED

**Evidencia:** CONFIRMED

---

## Seguridad / Licencia

### BR-014
**Nombre:** Validación de licencia por hardware key (dongle)

**Origen:** Legacy/License.cls + modSeguridadInfhotel.bas

**Archivo:** `legacy-restaurant/restaurant-vb6/Clases/License.cls`

**Descripción:** El sistema valida la licencia mediante un dongle físico al iniciar. El vencimiento de licencia dispara alertas por email vía Chilkat.

**Condición:** Al inicio del ejecutable

**Resultado:** Si el dongle es válido y la licencia no está vencida, el sistema continúa; si no, bloquea

**Excepciones:** PARTIAL — comportamiento exacto ante dongle ausente no determinado con certeza

**Destino .NET:** Definir modelo de licenciamiento (ver ADR-001)

**Estado:** NOT_STARTED

**Evidencia:** CONFIRMED

---

*Este documento se amplía con cada análisis de módulo.*
