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

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Domain/Services/TaxPolicy.cs`, `modern-net8/tests/Inforest.Domain.Tests/Services/TaxPolicyAndVisibilidadTests.cs`

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

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Domain/Services/TaxPolicy.cs`, `modern-net8/tests/Inforest.Domain.Tests/Services/TaxPolicyAndVisibilidadTests.cs`

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

**Estado:** IN_PROGRESS

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

**Estado:** IN_PROGRESS

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

**Estado:** IN_PROGRESS

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

**Estado:** IN_PROGRESS

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

**Destino .NET:** `IProduccionCocinaService` + `KdsXmlDispatcher` + `KdsLegacyGateway`

**Estado:** IN_PROGRESS

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

**Destino .NET:** `AreaProduccionRouter` + `ComandaAreaDispatcher` + `KitchenConfigurationProvider`

**Estado:** IN_PROGRESS

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

**Estado:** IN_PROGRESS

**Evidencia:** CONFIRMED

---

## Seguridad / Licencia

### BR-015
**Nombre:** Login por módulo con sesión operativa y auditoría de acceso

**Origen:** Legacy/frmAcceso.frm + modAuditoriaIntegral.bas + usp_Inforest_ObtieneUsuarios

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmAcceso.frm`

**Procedimiento/Función:** `imgOpcion_Click`, `AccesoInicio`, `registroAccesoAuditoria`

**Descripción:** El acceso al sistema filtra usuarios por módulo, valida password legacy o banda magnética, abre una sesión operativa por caja/terminal y registra el ingreso/salida en `MMOVIMIENTOACCESO`.

**Condición:** Al iniciar sesión en el ejecutable POS/Administración/Consulta.

**Resultado:** Solo ingresan usuarios activos habilitados para el módulo; la sesión queda asociada a usuario, caja, terminal y correlativo de auditoría.

**Excepciones:** La migración .NET reemplaza el cifrado XOR/César por BCrypt y deja la validación física de HardKey como gap controlado.

**Destino .NET:** `IAuthService`, `ISessionService`, `IRbacService`, `IAuditoriaService`, `Form1`

**Estado:** IN_PROGRESS

**Evidencia:** CONFIRMED

---

### BR-014
**Nombre:** Validación de licencia por hardware key (dongle)

**Origen:** Legacy/License.cls + modSeguridadInfhotel.bas

**Archivo:** `legacy-restaurant/restaurant-vb6/Clases/License.cls`

**Descripción:** El sistema valida la licencia mediante un dongle físico al iniciar. El vencimiento de licencia dispara alertas por email vía Chilkat.

**Condición:** Al inicio del ejecutable

**Resultado:** Si el dongle es válido y la licencia no está vencida, el sistema continúa; si no, bloquea

**Excepciones:** PARTIAL — comportamiento exacto ante dongle ausente no determinado con certeza

**Destino .NET:** `ILicenseService` + `LicenseService` con contrato SQL Legacy y hardkey físico pendiente

**Estado:** IN_PROGRESS

**Evidencia:** CONFIRMED

---

## Reglas SQL

> Reglas extraídas directamente de `legacy-restaurant/database-sql-server/5. SP.sql` y relacionadas con la lógica operativa más crítica de ventas e inventario.

### BR-VENTAS-SQL-001
**Nombre:** Correlativo global de pedido desde `TPARAMETRO`

**Origen:** SQL/spIns_MPEDIDO

**Archivo:** `legacy-restaurant/database-sql-server/5. SP.sql`

**Procedimiento/Función:** `spIns_MPEDIDO`

**Descripción:** Antes de insertar un pedido, el SP incrementa `TPARAMETRO.nCorrelativo` y usa el nuevo valor como correlativo interno del pedido.

**Condición:** Inserción de un nuevo pedido.

**Resultado:** `TPARAMETRO.nCorrelativo` aumenta en 1 y `MPEDIDO.nCorrelativo` recibe ese valor.

**Excepciones:** No hay control transaccional explícito en el script para concurrencia; requiere validación en migración.

**Destino .NET:** `Application/Commands/CreatePedidoCommand`

**Estado:** NOT_STARTED

**Evidencia:** CONFIRMED

---

### BR-VENTAS-SQL-002
**Nombre:** Código de pedido anual con prefijo de año

**Origen:** SQL/spIns_MPEDIDO

**Archivo:** `legacy-restaurant/database-sql-server/5. SP.sql`

**Procedimiento/Función:** `spIns_MPEDIDO`

**Descripción:** El código del pedido se genera con los dos últimos dígitos del año más una secuencia de 8 dígitos basada en el máximo `tCodigoPedido` del año actual.

**Condición:** Inserción de un nuevo pedido.

**Resultado:** `MPEDIDO.tCodigoPedido` queda con formato `YY########`.

**Excepciones:** Si no existe pedido previo del año, la secuencia inicia en `00000001`.

**Destino .NET:** `Domain/ValueObjects/PedidoId`

**Estado:** NOT_STARTED

**Evidencia:** CONFIRMED

---

### BR-VENTAS-SQL-003
**Nombre:** Derivación automática de salón desde la mesa

**Origen:** SQL/spIns_MPEDIDO

**Archivo:** `legacy-restaurant/database-sql-server/5. SP.sql`

**Procedimiento/Función:** `spIns_MPEDIDO`

**Descripción:** Si se informa una mesa, el salón real del pedido se obtiene desde `TMESA.tSalon`; solo si no hay mesa se usa el salón recibido como parámetro.

**Condición:** `@tMesa <> ''`.

**Resultado:** `MPEDIDO.tSalon` se persiste con el salón asociado a la mesa.

**Excepciones:** Si no hay mesa, el SP conserva `@tSalon`.

**Destino .NET:** `Domain/Services/MesaAssignmentService`

**Estado:** NOT_STARTED

**Evidencia:** CONFIRMED

---

### BR-VENTAS-SQL-004
**Nombre:** Inserción diferenciada por canal central de pedidos

**Origen:** SQL/spIns_MPEDIDO

**Archivo:** `legacy-restaurant/database-sql-server/5. SP.sql`

**Procedimiento/Función:** `spIns_MPEDIDO`

**Descripción:** El SP consulta `vTipoPedido.lCanalCentralPedidos`; según el canal, cambia qué columnas se completan durante la inserción (`tEntregarA`, `tTiporecepcion`, `fRegistro`).

**Condición:** `COUNT(codigo)` en `vTipoPedido` con `lCanalCentralPedidos = 0` o distinto.

**Resultado:** El pedido se inserta con diferentes valores operativos para venta local vs canal centralizado.

**Excepciones:** La diferencia exacta debe preservarse por estrategia de compatibilidad funcional, no por simplificación.

**Destino .NET:** `Application/Policies/TipoPedidoPolicy`

**Estado:** NOT_STARTED

**Evidencia:** CONFIRMED

---

### BR-VENTAS-SQL-005
**Nombre:** Persistencia de descuentos autorizados y origen de venta

**Origen:** SQL/spIns_MPEDIDO + spUpd_MPEDIDO

**Archivo:** `legacy-restaurant/database-sql-server/5. SP.sql`

**Procedimiento/Función:** `spIns_MPEDIDO`, `spUpd_MPEDIDO`

**Descripción:** Tanto alta como edición de pedido persisten `nDescuento`, `tDescuento`, `tObservacionDescuento`, `tUsuarioDescuento` y `CodigoOrigenVenta`.

**Condición:** Pedido creado o modificado con descuento y/o canal de origen.

**Resultado:** El pedido conserva trazabilidad de descuento autorizado y del canal/origen de venta.

**Excepciones:** No se observa validación de permisos dentro del SP; se asume resuelta en VB6 o capa superior.

**Destino .NET:** `Application/Commands/CreateOrUpdatePedidoCommand`

**Estado:** NOT_STARTED

**Evidencia:** CONFIRMED

---

### BR-VENTAS-SQL-006
**Nombre:** Sincronización de `TPEDIDOMESA` tras crear pedido

**Origen:** SQL/spIns_MPEDIDO

**Archivo:** `legacy-restaurant/database-sql-server/5. SP.sql`

**Procedimiento/Función:** `spIns_MPEDIDO`

**Descripción:** Después de insertar el pedido, el SP actualiza `TPEDIDOMESA` asignando el nuevo código donde exista `tCodigoPedido = ''`.

**Condición:** Inserción de un nuevo pedido.

**Resultado:** Las relaciones temporales mesa↔pedido quedan vinculadas al nuevo pedido persistido.

**Excepciones:** Requiere revisión de concurrencia para mesas múltiples o inserciones simultáneas.

**Destino .NET:** `Infrastructure/Repositories/PedidoMesaRepository`

**Estado:** NOT_STARTED

**Evidencia:** CONFIRMED

---

### BR-VENTAS-SQL-007
**Nombre:** Regularización automática de pedidos no entregados con programación nula

**Origen:** SQL/spIns_MPEDIDO

**Archivo:** `legacy-restaurant/database-sql-server/5. SP.sql`

**Procedimiento/Función:** `spIns_MPEDIDO`

**Descripción:** Tras cada inserción, el SP marca como entregados (`lEntregado = 1`) y fija `fProgramacion = fFecha` para pedidos recientes con `fProgramacion IS NULL` y no entregados.

**Condición:** Existencia de pedidos creados en el último día sin programación y sin entrega.

**Resultado:** Se corrigen pedidos incompletos para evitar estados operativos inválidos.

**Excepciones:** Es una corrección lateral del SP; debe analizarse si corresponde migrarla igual o encapsularla como job/regla separada.

**Destino .NET:** `Application/Policies/PedidoConsistencyPolicy`

**Estado:** NOT_STARTED

**Evidencia:** CONFIRMED

---

### BR-VENTAS-SQL-008
**Nombre:** Actualización integral del pedido respetando mesa/salón vigentes

**Origen:** SQL/spUpd_MPEDIDO

**Archivo:** `legacy-restaurant/database-sql-server/5. SP.sql`

**Procedimiento/Función:** `spUpd_MPEDIDO`

**Descripción:** La edición del pedido actualiza cliente delivery, tipo de pedido, prioridad, atención, mozo, motorizado, observaciones, datos hoteleros, descuentos, programación, invitado/pariente, entrega, origen de venta, mesa y salón.

**Condición:** Modificación de un pedido existente.

**Resultado:** `MPEDIDO` se actualiza en una sola operación SQL, recalculando `tSalon` desde `TMESA` si se cambió la mesa.

**Excepciones:** Si no hay mesa, conserva el salón previamente registrado en `MPEDIDO`.

**Destino .NET:** `Application/Commands/UpdatePedidoCommand`

**Estado:** NOT_STARTED

**Evidencia:** CONFIRMED

---

### BR-VENTAS-SQL-009
**Nombre:** Reporte de ventas con filtros dinámicos por cliente/documento/estado/caja/tipo de pago

**Origen:** SQL/spRep_RegVenta

**Archivo:** `legacy-restaurant/database-sql-server/5. SP.sql`

**Procedimiento/Función:** `spRep_RegVenta`

**Descripción:** El reporte arma criterios dinámicos en SQL para filtrar por cliente, tipo de documento, estado del documento, caja y tipo de pago.

**Condición:** Ejecución del reporte con alguno de esos parámetros informado.

**Resultado:** El dataset final de ventas se restringe a los criterios solicitados por el operador.

**Excepciones:** El SP usa SQL dinámico; la migración debe preservar el resultado sin reproducir riesgos de concatenación indiscriminada.

**Destino .NET:** `Application/Queries/ReporteRegistroVentaQuery`

**Estado:** NOT_STARTED

**Evidencia:** CONFIRMED

---

### BR-VENTAS-SQL-010
**Nombre:** Reporte de ventas con orden configurable y fecha operativa / día contable

**Origen:** SQL/spRep_RegVenta

**Archivo:** `legacy-restaurant/database-sql-server/5. SP.sql`

**Procedimiento/Función:** `spRep_RegVenta`

**Descripción:** El reporte soporta orden por correlativo, montos o fechas, utiliza tablas temporales (`#DBTRANS`, `#DBTRANS1`, `#TEMP`) y bifurca la lógica según `@diaContable` para fecha calendario vs fecha operativa.

**Condición:** Ejecución del reporte con combinación de `@sOrden`, `@flagAnoMes`, `@diaContable`, rangos de fecha y hora de corte.

**Resultado:** El registro de ventas consolida documentos y notas de crédito conforme a la fecha operativa esperada por negocio.

**Excepciones:** La migración debe validar contra Legacy la semántica exacta de hora de corte y agrupaciones.

**Destino .NET:** `Application/Queries/ReporteRegistroVentaQuery`

**Estado:** NOT_STARTED

**Evidencia:** CONFIRMED

---

### BR-ALMACEN-SQL-001
**Nombre:** Modificación de insumo con normalización de descripción y actualización condicional de stock

**Origen:** SQL/USP_MODIFICARINSUMOS

**Archivo:** `legacy-restaurant/database-sql-server/5. SP.sql`

**Procedimiento/Función:** `USP_MODIFICARINSUMOS`

**Descripción:** El SP normaliza la descripción del insumo a mayúsculas y, según `@MODULO`, actualiza o no el stock (`NSTOCK`) además de usuario, caja, activo, bandera `LINSUMO` y fecha de modificación.

**Condición:** Edición de un insumo desde mantenimiento.

**Resultado:** `TINSUMO` queda actualizado con auditoría básica y consistencia de descripción.

**Excepciones:** Si `@MODULO = 0`, el stock no se modifica.

**Destino .NET:** `Application/Commands/UpdateInsumoCommand`

**Estado:** NOT_STARTED

**Evidencia:** CONFIRMED

---

### BR-SQL-CMD-001
**Nombre:** Timeout de 600 segundos para ejecución de stored procedures

**Origen:** Legacy/clsComando.cls

**Archivo:** `legacy-restaurant/restaurant-vb6/Clases/clsComando.cls`

**Procedimiento/Función:** `CreateCmdSp`

**Descripción:** El Legacy establece `Cmd.CommandTimeout = 600` para todos los SPs ejecutados a través de `clsComando`. Este timeout cubre SPs de reporte pesados y operaciones batch en horarios de alta concurrencia.

**Condición:** Toda ejecución de SP vía `clsComando.cls`

**Resultado:** La conexión espera hasta 600 segundos antes de lanzar error de timeout.

**Excepciones:** Ninguna — el timeout es uniforme para todos los SPs.

**Destino .NET:** `SpExecutor` — constante `CommandTimeoutSeconds = 600`

**Estado:** COMPLETED

**Evidencia:** CONFIRMED

---

*Este documento se amplía con cada análisis de módulo.*

---

## Módulo Despachador / Motorizado (Etapa 9)

### BR-DEL-001

**Nombre:** Tipo de pedido determina si requiere motorizado y cliente frecuente

**Origen:** `TTIPOPEDIDODETALLE.lActivaMotorizado`, `lObligaMotorizado`, `lObligaClienteFrecuente`

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmNuevoDelivery.frm`

**Descripción:** Según el tipo de pedido, puede ser obligatorio asignar un motorizado y/o un cliente frecuente de delivery.

**Condición:** Al crear o confirmar un pedido delivery

**Resultado:** Si `lObligaMotorizado=true` el sistema exige motorizado; si `lObligaClienteFrecuente=true` exige cliente frecuente registrado en TDELIVERY.

**Excepciones:** Si el canal es `lCanalCentralPedidos`, la lógica puede diferir.

**Destino .NET:** `CrearPedidoDeliveryCommand`, `TipoPedidoDelivery`

**Estado:** MIGRATED

---

### BR-DEL-002

**Nombre:** Tarifa motorizado varía por tipo de día: LV/SD/ES

**Origen:** `TMOTORIZADODATOS.nTarifaLV`, `nTarifaSD`, `nTarifaES`

**Archivo:** `legacy-restaurant/restaurant-vb6/Modulos/modDespachador.bas`

**Descripción:** Cada motorizado tiene tres tarifas configuradas: lunes-viernes (LV), sábado-domingo (SD) y especial/feriado (ES). La tarifa activa en un momento dado está en `TPARAMETRO.tTarifaActualMotorizado`.

**Condición:** Al consultar la tarifa a aplicar en un día específico

**Resultado:** Se retorna `nTarifaLV`, `nTarifaSD` o `nTarifaES` según el tipo de día.

**Excepciones:** La tarifa activa puede ser sobreescrita manualmente vía `frmTarifaMotorizado`.

**Destino .NET:** `Motorizado.ObtenerTarifaPorTipoDia()`, `ActualizarTarifaMotorizadoCommand`

**Estado:** MIGRATED

---

### BR-DEL-003

**Nombre:** Flag `lCD` en TCAJA habilita modo CENTRALDELIVERY

**Origen:** `modDespachador.bas: CD = Calcular("select lCD...")`

**Archivo:** `legacy-restaurant/restaurant-vb6/Modulos/modDespachador.bas`

**Descripción:** Si `TCAJA.lCD = true`, la caja opera en modo Central de Pedidos, conectándose a `CENTRALDELIVERY` para sincronizar pedidos.

**Condición:** Al iniciar el módulo Despachador

**Resultado:** Se habilita `CentralPedidosRepository` y se muestra panel de pedidos centralizados.

**Excepciones:** Si `lCD = false`, solo se muestra el panel local.

**Destino .NET:** `ObtenerPedidosCentralQuery`, `CentralPedidosRepository`

**Estado:** MIGRATED

---

### BR-DEL-004

**Nombre:** Flag `lOrdenesRappi` en TCAJA habilita integración Rappi

**Origen:** `modDespachador.bas: lOrdenesRappi`

**Archivo:** `legacy-restaurant/restaurant-vb6/Modulos/modDespachador.bas`

**Descripción:** Si `TCAJA.lOrdenesRappi = true`, la consola muestra las órdenes recibidas desde Rappi y permite enviar datos.

**Condición:** Al iniciar el módulo Despachador

**Resultado:** Se habilita `OrdenesConsolaForm` y `EnvioDatosRappiCommand`.

**Excepciones:** Si `false`, la pestaña Rappi no se muestra.

**Destino .NET:** `RappiOrderAdapter`, `ObtenerOrdenesExternasQuery`

**Estado:** MIGRATED (stub; conector HTTP Rappi pendiente — GAP-DEL-004)

---

### BR-DEL-005

**Nombre:** `tTarifaActualMotorizado` en TPARAMETRO indica tarifa vigente

**Origen:** `modDespachador.bas`, `frmTarifaMotorizado.frm`

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmTarifaMotorizado.frm`

**Descripción:** `TPARAMETRO.tTarifaActualMotorizado` almacena el código de tarifa activa actualmente. Al cambiar tarifa desde `frmTarifaMotorizado`, se actualiza este campo.

**Condición:** Al ejecutar `ActualizarTarifaMotorizadoCommand`

**Resultado:** Se actualiza `TPARAMETRO.tTarifaActualMotorizado` con el código del nuevo tipo de día.

**Excepciones:** Requiere permisos de supervisión.

**Destino .NET:** `MotorizadoRepository.ActualizarTarifaActivaAsync()`

**Estado:** MIGRATED

---

### BR-DEL-006

**Nombre:** `Motorizado.exe` inicia directamente en `frmLlegadaSalida`

**Origen:** `modMotorizado.bas: frmLlegadaSalida.Show`

**Archivo:** `legacy-restaurant/restaurant-vb6/Modulos/modMotorizado.bas`

**Descripción:** El ejecutable Motorizado no muestra login; va directamente a la pantalla de control de llegada/salida del motorizado.

**Condición:** Al iniciar el proceso `Motorizado.exe`

**Resultado:** Se muestra `frmLlegadaSalida` como primera pantalla.

**Excepciones:** Ninguna — comportamiento fijo en Sub Main.

**Destino .NET:** `Program.cs` del módulo Motorizado → `Application.Run(new LlegadaSalidaForm(...))`

**Estado:** MIGRATED

---

### BR-DEL-007

**Nombre:** TLOCAL controla multi-local: IP, BD INFOREST por local, flag réplica

**Origen:** `TLOCAL` schema

**Archivo:** `legacy-restaurant/database-sql-server/1. Estructura.sql`

**Descripción:** Cada local registrado en `TLOCAL` tiene su propia IP, nombre de base de datos INFOREST y flag `lReplica`. El Despachador puede consolidar pedidos de múltiples locales.

**Condición:** Al consultar pedidos multi-local

**Resultado:** `LocalRepository` retorna locales activos; `CentralPedidosRepository` accede a `CENTRALDELIVERY`.

**Excepciones:** Si `lActivo = false`, el local no se consulta.

**Destino .NET:** `LocalRestaurante`, `ILocalRepository`, `LocalRepository`

**Estado:** MIGRATED

---

### BR-DEL-008

**Nombre:** Central Pedidos usa conexión separada a CENTRALDELIVERY configurada en INI

**Origen:** `modDespachador.bas: sRutaCD, sMDBCD`

**Archivo:** `legacy-restaurant/restaurant-vb6/Modulos/modDespachador.bas`

**Descripción:** La base `CENTRALDELIVERY` tiene su propia cadena de conexión, configurada en INFOREST.INI sección `[CentralDelivery]`. El módulo Despachador la lee al iniciar.

**Condición:** Al iniciar con `lCD = true`

**Resultado:** Se usa `IDbConnectionFactory` con `databaseName = "CENTRALDELIVERY"`.

**Excepciones:** Si no está configurada, se deshabilita el panel Central.

**Destino .NET:** `IDbConnectionFactory`, `CentralPedidosRepository`, `appsettings.json`

**Estado:** MIGRATED

---

### BR-DEL-009

**Nombre:** Estado delivery se modifica vía `sp_CD_Modificar_EstadoDelivery_Cabecera`

**Origen:** `frmDespachador.frm`

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmDespachador.frm`

**Descripción:** Cualquier cambio de estado de un pedido delivery (pendiente → en preparación → en camino → entregado) se persiste llamando al SP `sp_CD_Modificar_EstadoDelivery_Cabecera` sobre la base `CENTRALDELIVERY`.

**Condición:** Al ejecutar `ActualizarEstadoDeliveryCommand`

**Resultado:** Se llama el SP con el nuevo estado; ambas bases (INFOREST y CENTRALDELIVERY) quedan sincronizadas.

**Excepciones:** Si `lCD = false`, el estado se actualiza solo en INFOREST.

**Destino .NET:** `CentralPedidosRepository.ModificarEstadoDeliveryAsync()`

**Estado:** MIGRATED

---

### BR-DEL-010

**Nombre:** Tipo de pedido tiene flags canal: `lCanalCentralPedidos`, `lCanalDelivery`

**Origen:** `TTIPOPEDIDODETALLE`

**Archivo:** `legacy-restaurant/database-sql-server/1. Estructura.sql`

**Descripción:** Cada tipo de pedido puede habilitar/deshabilitar canales: `lCanalCentralPedidos` permite que aparezca en el Central de Pedidos; `lCanalDelivery` habilita el flujo delivery con motorizado.

**Condición:** Al filtrar pedidos en el panel Despachador o Central Pedidos

**Resultado:** Solo se muestran pedidos cuyo tipo tiene habilitado el canal correspondiente.

**Excepciones:** Tipos sin flags = comportamiento local standard.

**Destino .NET:** `TipoPedidoDelivery`, `ObtenerDeliveryPendientesQuery`, `ObtenerPedidosDespachadorQuery`

**Estado:** MIGRATED

---

### BR-DEL-011

**Nombre:** Actualización de foto de cliente delivery vía SP

**Origen:** `frmClienteDeliveryDetalle.frm: GuardarFoto()`

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmClienteDeliveryDetalle.frm`

**Procedimiento/Función:** `sp_UpdFotoDelivery`

**Descripción:** Cuando se selecciona una foto del cliente de delivery, el formulario construye un parámetro binario y ejecuta `sp_UpdFotoDelivery(@tCodigo, @oFoto)` para persistir `TDELIVERY.iFoto`.

**Condición:** Al guardar la foto del cliente delivery.

**Resultado:** Se actualiza la columna binaria `iFoto` del registro en `TDELIVERY` según `tCodigoDelivery`.

**Excepciones:** Si el cliente no existe, no debe ejecutarse la actualización.

**Destino .NET:** `ActualizarFotoClienteDeliveryHandler`, `IClienteDeliveryRepository.ActualizarFotoAsync`, `ClienteDeliveryRepository.ActualizarFotoAsync`

**Estado:** IN_PROGRESS

---

### BR-DEL-012

**Nombre:** Confirmar y revertir entrega de pedido (lEntregado) — frmCentralPedidos

**Origen:** `frmCentralPedidos.frm: Case 3 (confirmar) / Case 5 (revertir)`

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmCentralPedidos.frm`

**Procedimiento/Función:** SQL directo sobre MPEDIDO

**Descripción:** El formulario de Central de Pedidos permite confirmar que un pedido fue entregado al cliente y también revertir dicha confirmación. La confirmación actualiza los campos `lEntregado=1`, `tusuarioentregado` y `fregentregado=GETDATE()`. La reversión actualiza `lEntregado=0`. Ambas operaciones actualizan el usuario responsable.

**Condición:** Al ejecutar "Confirmar Entrega" (Case 3) o "Revertir Entrega" (Case 5) sobre un pedido seleccionado.

**Resultado:**
- Confirmar: `MPEDIDO.lEntregado=1`, `tusuarioentregado=@sUsuario`, `fregentregado=GETDATE()`
- Revertir: `MPEDIDO.lEntregado=0`, `tusuarioentregado=@sUsuario`, `fregentregado=GETDATE()`

**Excepciones:**
- No se puede confirmar entrega si el pedido ya tiene `lEntregado=1`.
- No se puede revertir si el pedido no está en estado `lEntregado=1`.
- La reversión siempre requiere supervisor (ver BR-DEL-013).

**Destino .NET:** `ConfirmarEntregaCentralHandler`, `RevertirEntregaCentralHandler`, `IPedidoDeliveryRepository.ConfirmarEntregaAsync`, `IPedidoDeliveryRepository.RevertirEntregaAsync`

**Estado:** MIGRATED

---

### BR-DEL-013

**Nombre:** Supervisor requerido para entrega sin pago y para revertir entrega

**Origen:** `frmCentralPedidos.frm: Case 3 (Supervisor("22")), Case 5 (Supervisor("22"))`

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmCentralPedidos.frm`

**Procedimiento/Función:** `Supervisor("22")` en `modProcedimiento.bas`

**Descripción:** Cuando se intenta confirmar la entrega de un pedido que no ha sido cancelado (estado pago = "NO PAGADO", "POR COBRAR" o "ANTICIPO"), el sistema exige autorización de supervisor mediante el código de acción "22". Asimismo, cualquier reversión de entrega también requiere autorización de supervisor con el mismo código "22".

**Condición:**
- Confirmar entrega con pedido en estado "NO PAGADO", "POR COBRAR" o "ANTICIPO" sin supervisor → muestra pregunta y pide clave.
- Revertir entrega siempre → pide clave de supervisor.

**Resultado:** Si el supervisor no autoriza, la operación se cancela. Si autoriza, la operación procede.

**Excepciones:** Si el pedido está pagado, la confirmación procede sin supervisor.

**Destino .NET:** `ConfirmarEntregaCentralHandler` (campo `SupervisorAutorizado`), `RevertirEntregaCentralHandler` (campo `SupervisorAutorizado`), `CentralPedidosForm.PromptSupervisor()`

**Estado:** MIGRATED

---

### BR-DEL-014

**Nombre:** Modificar fecha programada de entrega (fProgramacion / fRegistro)

**Origen:** `frmCentralPedidos.frm: Case 2`

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmCentralPedidos.frm`

**Procedimiento/Función:** SQL directo sobre MPEDIDO

**Descripción:** El formulario de Central de Pedidos permite cambiar la fecha y hora programada de entrega de un pedido. Actualiza simultáneamente los campos `fregistro` y `fProgramacion` de MPEDIDO con la nueva fecha seleccionada.

**Condición:** Al ejecutar "Modificar Fecha" sobre un pedido seleccionado que no haya sido entregado.

**Resultado:** `MPEDIDO.fregistro = @nuevaFecha`, `MPEDIDO.fProgramacion = @nuevaFecha`

**Excepciones:** No se puede modificar la fecha de un pedido ya entregado (`lEntregado=1`). La nueva fecha no puede ser vacía (`default`).

**Destino .NET:** `ModificarFechaProgramadaDeliveryHandler`, `IPedidoDeliveryRepository.ModificarFechaProgramadaAsync`, `CentralPedidosForm` (Case 2)

**Estado:** MIGRATED

---

---

### BR-REP-001

**Nombre:** Reporte de Comanda — modo detalle vs. cabecera

**Origen:** `frmRepComanda.frm`

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmRepComanda.frm`

**Procedimiento/Función:** `spRep_Comanda`

**Descripción:** El reporte de comanda tiene dos modos: `@flagTipo=1` muestra un registro por ítem DPEDIDO (detallado); `@flagTipo=0` muestra agrupado por comanda. La opción de ordenamiento (`@SOrden`) puede ser por Pedido, Comanda, Mozo o Fecha.

**Condición:** Al ejecutar `ObtenerReporteComandaQuery`

**Resultado:** Dataset con columnas: tCodigoPedido, tComanda, NombreProducto, Mozo, nCantidad, PrecioUnitario, PrecioTotal, fFecha, Usuario, tDocumento, Estado, tObservacion

**Excepciones:** Si `@sCriterio` no está vacío, se agrega como cláusula WHERE adicional en el SQL dinámico del SP.

**Destino .NET:** `ObtenerReporteComandaHandler`, `FrmComandaReporte.cs`, `RepComanda.frx`

**Estado:** IN_PROGRESS

---

### BR-REP-002

**Nombre:** Reporte de Propinas — filtro por condición dinámica

**Origen:** `frmRepPropina.frm`

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmRepPropina.frm`

**Procedimiento/Función:** `spRep_Propina`

**Descripción:** Muestra propinas de tipo tarjeta (`tTipoPago='02'` y `nPropina>0`) en el rango de fechas. El SP construye SQL dinámico. El parámetro `@sCondicion` permite filtrar adicionalmente por mozo, motorizado o tarjeta.

**Condición:** Al ejecutar `ObtenerReportePropinaQuery`

**Resultado:** Dataset con Mozo, Motorizado, Tarjeta, Propina, Documento, Fecha, Trabajador, TipoPedido, FactorRetencion

**Excepciones:** Si `@sCondicion` no está vacío, se añade como AND al WHERE del SQL dinámico.

**Destino .NET:** `ObtenerReportePropinaHandler`, `FrmPropinaReporte.cs`, `RepPropina.frx`

**Estado:** IN_PROGRESS

---

### BR-REP-003

**Nombre:** Reporte de Clientes Principales — monto mínimo y modo

**Origen:** `frmRepPrincipal.frm`

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmRepPrincipal.frm`

**Procedimiento/Función:** `spRep_PrincipalCliente`

**Descripción:** Identifica clientes con consumo acumulado mayor o igual a `@sMonto` en el rango de fechas. `@flagTipo=1` retorna el detalle por documento; `@flagTipo=0` retorna el resumen por cliente con totales. `@SCliente` vacío = todos los clientes.

**Condición:** Al ejecutar `ObtenerReportePrincipalClienteQuery`

**Resultado:** Modo detalle: Cliente, Empresa, Documento, Fecha, Neto, Impuestos, Venta. Modo resumen: Cliente, Empresa, CantidadDocumentos, Totales.

**Excepciones:** La subquery de clientes calificados usa el mismo rango de fechas que el resultado principal.

**Destino .NET:** `ObtenerReportePrincipalClienteHandler`, `RepPrincipalClienteDetalle.frx`, `RepPrincipalClienteResumen.frx`

**Estado:** IN_PROGRESS

---

### BR-REP-004

**Nombre:** Reporte Cuenta Corriente Integrado — tres modos de visualización

**Origen:** `frmRepCtaCteIntegrado.frm`

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmRepCtaCteIntegrado.frm`

**Procedimiento/Función:** `spRep_CtaCteIntegrado`

**Descripción:** `@flagTipo='1'`=consolidado por cliente, `'2'`=detalle por documento, `'3'`=estado de cuenta. Consulta las tablas MDOCUMENTO y TCLIENTE (cuenta corriente habilitada en MPEDIDO.tClienteCtaCte).

**Condición:** Al ejecutar `ObtenerReporteCtaCteIntegradoQuery`

**Resultado:** Dataset con información de cuenta corriente según el modo seleccionado.

**Destino .NET:** `ObtenerReporteCtaCteIntegradoHandler`, `RepCtaCteIntegradoConsolidado.frx`, `RepCtaCteIntegradoDetallado.frx`

**Estado:** IN_PROGRESS

---

### BR-REP-005

**Nombre:** Paloteo Comparativo — modo turno vs. fecha y cantidad vs. valor

**Origen:** `frmRepPaloteoComparativo.frm`

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmRepPaloteoComparativo.frm`

**Procedimiento/Función:** `spRep_PaloteoComparativo`

**Descripción:** Compara ventas por producto desglosado en canales (Salon/Delivery/Llevar/Canal4/Canal5). `@flagTurnoOFecha=0` filtra por turno; `=1` filtra por rango de fechas. `@flagTipoValor=0` muestra cantidad; `=1` muestra valor monetario. `@flagNFacturado=1` incluye solo pedidos con documento emitido.

**Condición:** Al ejecutar `ObtenerReportePaloteoComparativoQuery`

**Resultado:** Columnas: Codigo, Grupo, SubGrupo, Producto, Valor, Produccion, Venta, Cortesia, CtaCte, Canal1-5

**Destino .NET:** `ObtenerReportePaloteoComparativoHandler`, `FrmPaloteoComparativoReporte.cs`, `RepPaloteoComparativo.frx`

**Estado:** IN_PROGRESS

---

### BR-REP-006

**Nombre:** Paloteo Sub-Productos — apertura por componentes de combo

**Origen:** `frmRepPaloteoSubProd.frm`

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmRepPaloteoSubProd.frm`

**Procedimiento/Función:** `spRep_PaloteoSubProd`

**Descripción:** Muestra productos vendidos desagregados en sub-productos (componentes de combos). Usa la vista `vPRODUCTOXPRODUCTO` para obtener los sub-productos de cada ítem vendido. `@flagTurno=0` filtra por turno; `=1` por rango de fecha.

**Condición:** Al ejecutar `ObtenerReportePaloteoSubProdQuery`

**Resultado:** Local, Salón, Mesa, TipoProducto, Grupo, SubGrupo, Producto, Cantidad, Venta, SubProducto, CantProd

**Destino .NET:** `ObtenerReportePaloteoSubProdHandler`, `RepPaloteoSubProd.frx`

**Estado:** IN_PROGRESS

---

### BR-REP-SQL-DYN-001

**Nombre:** Expresión de precio — protección contra SQL injection en SPs dinámicos

**Origen:** `frmRepPaloteoVentaIntegrado.frm`, `frmRepRankingIntegrado.frm`, `frmRepDiarioVentaIntegrado.frm`, `frmRepAnaliticoMotorizadoIntegrado.frm`

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmRepPaloteoVentaIntegrado.frm`

**Descripción:** Los SPs de integrado reciben un parámetro `@sPrecio` que se inyecta directamente como expresión SQL en el SELECT. En VB6, este valor se construye en el formulario (ej. `"dbo.DPEDIDO.nVenta"`). En .NET 8, NUNCA se acepta este valor como texto libre del usuario: solo se permite a través del enum `ExpresionPrecio` con valores controlados (Venta, Neto, Costo). `ExpresionPrecioExtensions.ToSqlExpresion()` convierte el enum a la expresión SQL exacta que el SP Legacy espera.

**Condición:** Al construir los parámetros de los handlers de reportes integrados

**Resultado:** Expresión SQL segura, sin caracteres de injection (`;`, `--`, `'`)

**Excepciones:** Ninguna — el enum cubre todos los casos del Legacy.

**Destino .NET:** `ExpresionPrecio`, `ExpresionPrecioExtensions`, handlers de reportes integrados

**Estado:** MIGRATED

---

### BR-REP-009

**Nombre:** Venta Mensual Integrado — hora de corte del día contable

**Origen:** `frmRepVentaMensualIntegrado.frm`, `frmRepDiarioVentaIntegrado.frm`

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmRepVentaMensualIntegrado.frm`

**Procedimiento/Función:** `spRep_VentaMensualIntegrado`

**Descripción:** El parámetro `@dHour` define la hora de corte del día contable. Si un pedido se registra antes de esa hora, se asigna al día anterior. Esto permite que un restaurante que cierra a las 2 AM tenga la venta correctamente asignada al día operativo (no calendario). El SP agrupa por día y por tipo de pedido para mostrar venta de Salon, Delivery, Llevar y canales adicionales.

**Condición:** Al ejecutar `ObtenerReporteVentaMensualIntegradoQuery`

**Resultado:** Fila por día con venta por canal, cantidad de pedidos, pax y costo

**Excepciones:** `@tipo='D'` = informe diario; `'M'` = informe mensual acumulado

**Destino .NET:** `ObtenerReporteVentaMensualIntegradoHandler`, `FrmVentaMensualIntegradoReporte.cs`, `RepVentaMensualIntegrado.frx`

**Estado:** IN_PROGRESS

---

### BR-REP-011

**Nombre:** KDS — Reporte de tiempos de preparación por pedido

**Origen:** KDS monitor (módulo cocina)

**Archivo:** `legacy-restaurant/database-sql-server/5. SP.sql`

**Procedimiento/Función:** `USP_KDS_ResporteTiempoPedido`

**Descripción:** Calcula tiempos de preparación KDS (cocina) por pedido. `fSalida` es cuando el ítem fue despachado desde la cocina; `fEnvio` o `fRegistro` es cuando entró al KDS. El SP calcula mínimo, máximo y promedio de tiempo por producto en el rango de fechas. El tiempo promedio se calcula como la media de todos los ítems del producto.

**Condición:** Al ejecutar `ObtenerReporteTiempoKdsPedidoQuery`

**Resultado:** Pedido, Producto, TiempoCorto (HH:mm:ss), TiempoLargo, TiempoPromedio

**Excepciones:** Si `DPEDIDOKDS.fSalida` es NULL, el registro se excluye del cálculo.

**Destino .NET:** `ObtenerReporteTiempoKdsPedidoHandler`, `FrmTiempoKdsReporte.cs`, `RepTiempoKdsPedido.frx`

**Estado:** IN_PROGRESS

---

### BR-REP-012

**Nombre:** KDS — Reporte de tiempos de preparación por producto

**Origen:** KDS monitor (módulo cocina)

**Archivo:** `legacy-restaurant/database-sql-server/5. SP.sql`

**Procedimiento/Función:** `USP_KDS_ResporteTiempoProducto`

**Descripción:** Similar a BR-REP-011 pero agrupa por Grupo/SubGrupo/Producto para análisis estadístico de producción. Permite filtrar por grupo, subgrupo y producto. El SP usa las tablas DPEDIDOKDS, DPEDIDO, TPRODUCTO, TGRUPO y TSUBGRUPO.

**Condición:** Al ejecutar `ObtenerReporteTiempoKdsProductoQuery`

**Resultado:** Grupo, SubGrupo, Producto, TiempoCorto, TiempoLargo, TiempoPromedio

**Destino .NET:** `ObtenerReporteTiempoKdsProductoHandler`, `FrmTiempoKdsReporte.cs`, `RepTiempoKdsProducto.frx`

**Estado:** IN_PROGRESS

---

## Reglas de Periféricos Especializados e Integraciones por País (P3-11)

### BR-PERIPH-001

**Nombre:** PinPad — operaciones de cobro con tarjeta

**Origen:** Legacy/DLL3500.bas

**Archivo:** `legacy-restaurant/restaurant-vb6/Modulos/DLL3500.bas`

**Procedimiento/Función:** `fiOpenPort`, `fiClosePort`, `fiStartOperation`, `fiGetStatus` via `caja_pinpad.dll`

**Descripción:** El PinPad se comunica via DLL nativa Win32 (`caja_pinpad.dll`). El Legacy declara las funciones con `Declare Function` e invoca por operación (`OP_VENTA`, `OP_ANULACION`, `OP_CIERRE`, `OP_SIMULACION_CUOTAS`, `OP_TEST_COMUNICACIONES`, `OP_DUPLICADO`). El flujo de estados es `RET_RUNNING (1)` → `RET_OK (0)` / `RET_NOK (-1)`. Errores específicos: `ERR_COM`, `ERR_PPAD_NO_RESP`, `ERR_SOCKET`, `ERR_HOST_NO_RESP`, `ERR_PRINTER`, `ERR_NAK_RECEIVED`, `ERR_LRC_PPAD`. El separador de campos es `FS = Chr(28)`.

**Condición:** Se activa al seleccionar pago con tarjeta en `frmPago.frm` / `frmPagoPinPad.frm`.

**Resultado:** Cadena de respuesta parseada con `MensajePinPad()`. Imprime cupón con `ImprimeCabecera()`. El formato de encabezado varía por país: Argentina usa `C.U.I.T.`, Perú usa `R.U.C.`.

**Excepciones:** `ERR_COM (-2)`, `ERR_PPAD_NO_RESP (-3)`, `ERR_SOCKET (-5)`, `ERR_HOST_NO_RESP (-6)`. Se muestra MsgBox al operador.

**Destino .NET:** `IPinPadService`, `PinPadService` (P/Invoke), `NullPinPadService`, `ProcesarPagoPinPadHandler`, `ObtenerTerminalesPinPadHandler`, `FrmPago`, `FrmPagoPinPad`

**Estado:** IN_PROGRESS — flujo WinForms y selección de terminal migrados; parsing completo de respuesta/cupón legacy sigue parcial

---

### BR-PERIPH-002

**Nombre:** CashDro — cajón automático de dinero

**Origen:** Legacy/modProcedimientoNuevo.bas

**Archivo:** `legacy-restaurant/restaurant-vb6/Modulos/modProcedimientoNuevo.bas`

**Procedimiento/Función:** `IniciarMotorCashDrow(Argumentos As String)`

**Descripción:** El Legacy lanza el proceso externo `CashDrow\MotorCashDrow.exe` mediante `ShellExecute`. La comunicación es por proceso/argumentos, no por DLL directa. Si el ejecutable no existe, muestra MsgBox de error. Los argumentos determinan la operación (abrir, cerrar, estado).

**Condición:** Al procesar un pago o apertura de turno cuando CashDro está habilitado en la configuración de caja.

**Resultado:** El motor CashDro ejecuta la operación de cajón. El Legacy no espera retorno síncrono.

**Excepciones:** Si `MotorCashDrow.exe` no existe, se muestra mensaje al administrador.

**Destino .NET:** `ICashDroService`, `CashDroService` (Process.Start), `NullCashDroService`

**Estado:** IN_PROGRESS

---

### BR-PERIPH-003

**Nombre:** BlueVision/TVS — sistema de visualización de tickets en mesa

**Origen:** Legacy/modBlueVision.bas

**Archivo:** `legacy-restaurant/restaurant-vb6/Modulos/modBlueVision.bas`

**Procedimiento/Función:** `TVS_EnviarTicket(RsImpresion, sArea, sEvento, sPedido)`

**Descripción:** BlueVision TVS es un sistema de visualización de tickets para mesas. El Legacy se comunica via SDK COM (`BlueVision_Core_TVS.dll`) con autenticación por login/password leídos de `BLUEVISION.INI`. El flujo es: 1) CreateSession con TvsLogin, 2) SaveTicket con TvsTicket (id GUID, mode, room, table, total, TPV, evento, mozo), 3) SaveTicketLine por cada ítem de `DPEDIDO`. El total se calcula con `TVS_ObtenerTotal()`. El GUID se genera con `CoCreateGuid` via ole32.dll.

**Condición:** Al guardar/modificar un pedido cuando BlueVision está habilitado en la configuración.

**Resultado:** Ticket y líneas registrados en el servidor BlueVision. El Legacy escribe log con `TVS_EscribirLog()`.

**Excepciones:** Cualquier error COM se captura en `ErrorHandler` y se loguea.

**Destino .NET:** `IBlueVisionService`, `BlueVisionHttpClient` (HttpClient nativo), `NullBlueVisionService`, `BlueVisionOptions`

**Estado:** IN_PROGRESS

---

### BR-PERIPH-004

**Nombre:** Biometría SecuGen — lectura y verificación de huella dactilar

**Origen:** Legacy/FpLibX_Const.bas, sgfplibx.ocx

**Archivo:** `legacy-restaurant/restaurant-vb6/Modulos/FpLibX_Const.bas`

**Procedimiento/Función:** SDK SecuGen FpLibX via OCX `sgfplibx.ocx`

**Descripción:** El Legacy usa el SDK SecuGen FP (FpLibX) para captura y verificación de huella. Los códigos de error están definidos en `FpLibX_Const.bas`: dispositivo no encontrado (`ERROR_DEVICE_NOT_FOUND = 55`), timeout (`ERROR_TIME_OUT = 54`), ya abierto (`ERROR_DEV_ALREADY_OPEN = 59`), fallo de extracción (`ERROR_EXTRACT_FAIL = 105`), fallo de match (`ERROR_MATCH_FAIL = 106`). El template puede ser ANSI378 (256), SG400 (512) o ISO19794 (768). Se usa en login y en operaciones que requieren autorización biométrica.

**Condición:** Cuando la configuración de biometría está habilitada en `TPARAMETRO`/`TCAJA`.

**Resultado:** Template de huella capturado; verificación con score de coincidencia.

**Excepciones:** Múltiples códigos de error de dispositivo y de algoritmo de matching.

**Destino .NET:** `IBiometriaService`, `BiometriaSecuGenService` (P/Invoke a SDK SecuGen .NET si disponible), `NullBiometriaService`. GAP: OCX sin SDK .NET oficial disponible — usar Null hasta reemplazo de hardware.

**Estado:** IN_PROGRESS

---

### BR-PERIPH-005

**Nombre:** Impresora Fiscal Epson — emisión de documentos fiscales (Argentina)

**Origen:** Legacy/modImpresoraFiscal.bas, IFEpson.ocx

**Archivo:** `legacy-restaurant/restaurant-vb6/Modulos/modImpresoraFiscal.bas`

**Procedimiento/Función:** `SeteoFactura(PrinterOCX As PrinterFiscal)` via OCX `IFEpson.ocx`

**Descripción:** La impresora fiscal Epson (LX-300 Fiscal, FX-880 Fiscal) se controla via OCX `IFEpson.ocx` (32 bits, propiedad intelectual Epson Argentina). El setup define: 1) `SetPaperSize(cols, filas)` para papel continuo 12"×10", 2) `SetPreference` para estación de emisión Facturas A / documentos no fiscales, 3) `SetZone` para cada zona del formulario (letra comprobante, logo, razón social, datos vendedor, nro factura, CUIT, datos comprador, detalle de venta con ítems). Cada zona tiene posición (col, fila inicial, col final, fila final). Los ítems se dividen en zonas: cantidad (zona 61), descripción (zona 62), precio unitario (zona 63), tasa IVA (zona 64), tasa ajuste (zona 65).

**Condición:** Solo para Argentina (`pais = "003"`). Se activa al emitir documentos fiscales.

**Resultado:** Documento físico impreso con formato fiscal argentino regulado por AFIP.

**Excepciones:** Cada operación retorna booleano `Continuar`; si falsa, se aborta la emisión.

**Destino .NET:** `IImpresoraFiscalService`, `ImpresoraFiscalEpsonService` (protocolo serial RS-232/USB o proceso host 32-bit), `NullImpresoraFiscalService`. GAP: OCX de 32 bits sin SDK .NET oficial — registrado en known-gaps.

**Estado:** IN_PROGRESS

---

### BR-PERIPH-006

**Nombre:** Facturación Electrónica — emisión de comprobantes electrónicos por país

**Origen:** Legacy/modDeclaracion.bas, clsTrama.cls, clsxml.cls

**Archivo:** `legacy-restaurant/restaurant-vb6/Modulos/modDeclaracion.bas`

**Procedimiento/Función:** Variables globales `lFacturacionE`, `tCodigoFE`, `sRutaFE`, `sMDBFE`, `CnFE`, `lQRFE`, `RutaImgFE`, `clsTramaFE`, `TipoFacturacion`

**Descripción:** La FE en el Legacy se gestiona via base auxiliar MDB (`FACTURACION`). La variable `lFacturacionE` (boolean) habilita el flujo. `tCodigoFE` identifica el tipo de documento electrónico. `sRutaFE` es la ruta a la base MDB. `CnFE` es la conexión ADO a `FACTURACION`. `lQRFE` habilita QR en el comprobante. `clsTramaFE` (clase `clsTrama`) maneja el armado de la trama/XML. `TipoFacturacion` diferencia entre modalidades. La clase `clsxml.cls` arma el XML del comprobante según país. `registroventasunat` (comentario en código) indica integración SUNAT para Perú. La flag `FACTURACION PAPERLEES` indica comprobante sin papel físico.

**Condición:** Cuando `lFacturacionE = True` (configurado en `TPARAMETRO`/`TCAJA`).

**Resultado:** XML de comprobante electrónico generado y enviado al proveedor FE/autoridad fiscal del país.

**Excepciones:** Errores de conexión a base MDB, errores de validación de trama XML, rechazo del proveedor FE.

**Destino .NET:** `IFacturacionElectronicaGateway` (polimórfico por país), `FacturacionElectronicaFactory`, `PeruFEGateway` (SUNAT/OSE), `ArgentinaFEGateway` (AFIP), `ChileFEGateway` (SII), `EcuadorFEGateway` (SRI), `BoliviaFEGateway` (SIAT), `NullFEGateway`

**Estado:** IN_PROGRESS

---

### BR-PERIPH-007

**Nombre:** Multi-país — configuración fiscal y de periféricos por país

**Origen:** Legacy/modDeclaracion.bas, DLL3500.bas, scripts opcionales SQL

**Archivo:** `legacy-restaurant/restaurant-vb6/Modulos/modDeclaracion.bas`, `legacy-restaurant/database-sql-server/opcionales/`

**Procedimiento/Función:** Variable global `pais`, scripts `scriptPeruAlIniciar.sql`, `scriptArgentinaAlIniciar.sql`, etc.

**Descripción:** El país del local está codificado en la variable global `pais`. Valores conocidos: `"003"` = Argentina (usa C.U.I.T., impresora fiscal Epson, AFIP FE); default/vacío = Perú (usa R.U.C., SUNAT FE). Otros países: Chile (`scriptChileAlIniciar.sql`), Bolivia (`scriptBoliviaAlIniciar.sql`), Ecuador (`scriptEcuadorAlIniciar.sql`), España (`scriptEspanaAlIniciar.sql`). Cada script configura parámetros fiscales, tipos de impuestos, formatos de documento y integraciones específicas. El comportamiento de impresión del PinPad también varía por país (cabecera con RUC vs CUIT en `DLL3500.bas`).

**Condición:** Al inicializar el sistema, se lee el código de país y se aplica la configuración correspondiente.

**Resultado:** Comportamiento fiscal, de impresión y de integración adaptado al país del local.

**Excepciones:** País no reconocido → comportamiento por defecto (Perú).

**Destino .NET:** `IPaisPolicy`, `ConfiguracionPais`, `PaisPolicyFactory`, `PeruPaisPolicy`, `ArgentinaPaisPolicy`, `ChilePaisPolicy`, `EcuadorPaisPolicy`, `BoliviaPaisPolicy`, `EspanaPaisPolicy`

**Estado:** IN_PROGRESS

---

## Configuración del Sistema

### BR-SQL-003
**Nombre:** TPARAMETRO — configuración funcional global del sistema

**Origen:** Legacy/modPuntoVenta.bas + TPARAMETRO

**Archivo:** `legacy-restaurant/restaurant-vb6/Modulos/modPuntoVenta.bas`

**Procedimiento/Función:** `SELECT * FROM TPARAMETRO` al inicio del sistema; valores cargados en variables globales de `modDeclaracion.bas`

**Descripción:** La tabla `TPARAMETRO` tiene un único registro con 150+ columnas que controlan el comportamiento global del sistema: impuestos, correlativos, canales, monedas, día contable automático/manual, habilitación de módulos (almacén, KDS, delivery, FE), flags de impresión, configuración de email, integración con SAP, parámetros de país. Todos los módulos leen estas variables globales para determinar su comportamiento.

**Condición:** Al iniciar cualquier módulo (`Sub Main`) se carga `TPARAMETRO`.

**Resultado:** Variables globales disponibles en todo el sistema; comportamiento dinámico según configuración.

**Excepciones:** Si `TPARAMETRO` no existe o está vacía, el sistema no puede iniciar.

**Destino .NET:** `ConfiguracionSistema` (record tipado), `IParametroRepository.ObtenerConfiguracionAsync`, `ParametroService` con cache lazy, `IParametroService.ObtenerBoolAsync/ObtenerNumericoAsync`

**Estado:** COMPLETED

**Evidencia:** `modern-net8/src/Inforest.Domain/Entities/Configuracion/ConfiguracionSistema.cs`, `modern-net8/src/Inforest.Infrastructure/Configuracion/ParametroRepository.cs`, `modern-net8/src/Inforest.Infrastructure/Configuracion/ParametroService.cs`, `modern-net8/tests/Inforest.Infrastructure.Tests/Configuracion/ParametroRepositoryTests.cs`

---

### BR-SQL-004
**Nombre:** TCAJA — configuración operativa por caja/terminal

**Origen:** Legacy/modPuntoVenta.bas + TCAJA

**Archivo:** `legacy-restaurant/restaurant-vb6/Modulos/modPuntoVenta.bas`

**Procedimiento/Función:** `SELECT * FROM TCAJA WHERE tCaja = @sCaja` al inicio; valores cargados en variables globales de `modDeclaracion.bas`

**Descripción:** La tabla `TCAJA` tiene una fila por caja física con 100+ columnas que controlan el comportamiento de esa terminal específica: tipo de comanda (wComanda/vComanda), flags de password para cada operación, tipos de consumo habilitados, opciones de preimpresión, modo delivery, acceso a módulos específicos, configuración de balanza, integración móvil. Cada caja puede tener comportamiento diferente dentro del mismo restaurante.

**Condición:** Al iniciar una sesión en una caja determinada, se carga su configuración.

**Resultado:** Comportamiento diferenciado por terminal física; flags disponibles en toda la sesión.

**Excepciones:** Si la caja no existe en `TCAJA` o `lActivo = 0`, la caja no puede operar.

**Destino .NET:** `ConfiguracionCaja` (record tipado), `IParametroRepository.ObtenerConfiguracionCajaAsync`, `ObtenerConfiguracionCajaHandler`

**Estado:** COMPLETED

**Evidencia:** `modern-net8/src/Inforest.Domain/Entities/Configuracion/ConfiguracionCaja.cs`, `modern-net8/src/Inforest.Application/Configuracion/ConfiguracionHandlers.cs`, `modern-net8/src/Inforest.Infrastructure/Configuracion/ParametroRepository.cs`, `modern-net8/tests/Inforest.Application.Tests/Configuracion/ValidarInicioPosHandlerTests.cs`, `modern-net8/tests/Inforest.Infrastructure.Tests/Configuracion/ConfiguracionCajaTests.cs`

---

### BR-SQL-001
**Nombre:** spIns_MPEDIDO — inserción de pedido con correlativo automático

**Origen:** Legacy/spIns_MPEDIDO

**Archivo:** `legacy-restaurant/database-sql-server/5. SP.sql`

**Procedimiento/Función:** `spIns_MPEDIDO`

**Descripción:** El SP inserta una cabecera de pedido en `MPEDIDO` generando un correlativo automático basado en `TPARAMETRO.nCorrelativo` y actualizando `TPEDIDOMESA` (estado de la mesa). El código de pedido (`tCodigoPedido`) se genera como correlativo anual único. El SP garantiza atomicidad.

**Condición:** Al crear un nuevo pedido desde cualquier módulo (POS, Delivery, Adición).

**Resultado:** Registro en `MPEDIDO` + actualización `TPEDIDOMESA` + correlativo incrementado.

**Excepciones:** Si la mesa ya tiene un pedido activo (estado ocupado), error.

**Destino .NET:** `CreatePedidoCommand`, `CreatePedidoHandler` (llama SP Legacy via Dapper)

**Estado:** IN_PROGRESS

**Evidencia:** CONFIRMED

---

### BR-SQL-002
**Nombre:** spUpd_MPEDIDO — actualización integral del pedido

**Origen:** Legacy/spUpd_MPEDIDO

**Archivo:** `legacy-restaurant/database-sql-server/5. SP.sql`

**Procedimiento/Función:** `spUpd_MPEDIDO`

**Descripción:** El SP actualiza estado, datos y totales del pedido. Puede actualizar mesa, mozo, estado, turno, totales de venta y datos de delivery. Maneja la transacción completa de actualización.

**Condición:** Al modificar cualquier dato del pedido (cambio de mesa, mozo, estado).

**Resultado:** Registro MPEDIDO actualizado consistentemente.

**Excepciones:** Si el pedido está en estado 'AN' (anulado) o 'CO' (cobrado), no se puede actualizar.

**Destino .NET:** `UpdatePedidoCommand`, `UpdatePedidoHandler` (llama SP Legacy via Dapper)

**Estado:** IN_PROGRESS

**Evidencia:** CONFIRMED

---

### BR-015
**Nombre:** TaxPolicy — política de impuestos por producto

**Origen:** Legacy/modDeclaracion.bas + TPRODUCTO

**Archivo:** `legacy-restaurant/restaurant-vb6/Modulos/modDeclaracion.bas`, `legacy-restaurant/database-sql-server/1. Estructura.sql`

**Procedimiento/Función:** Variables globales `nPorcentaje1/2/3` + flags `lImpuesto1..15` en `TPRODUCTO`

**Descripción:** Cada producto tiene hasta 15 flags de impuesto booleanos (lImpuesto1..lImpuesto15). Los porcentajes globales (Impuesto1/2/3) se cargan desde TPARAMETRO al iniciar el sistema. Para calcular impuestos de una línea: monto = precioVenta × (porcentaje/100) si el flag del producto está activo. El precio neto = precioVenta / (1 + %total/100).

**Condición:** Al emitir documentos y calcular totales de pedido/documento.

**Resultado:** Impuesto1/2/3 calculado por línea de detalle. Precio neto derivado del precio de venta con impuestos.

**Excepciones:** Si ningún flag está activo, impuesto = 0 y precioNeto = precioVenta.

**Destino .NET:** `TaxPolicy` (`Domain/Services/TaxPolicy.cs`)

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Domain/Services/TaxPolicy.cs` + tests en `Inforest.Domain.Tests`

---

### BR-016
**Nombre:** ProductoVisibilidadService — visibilidad de producto por canal de venta

**Origen:** Legacy/frmProducto.frm + TPRODUCTO

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmProducto.frm`

**Procedimiento/Función:** Flags `lLocal`, `lDelivery`, `lLlevar`, `lCanal4`, `lCanal5` en `TPRODUCTO`

**Descripción:** Cada producto puede estar habilitado o deshabilitado individualmente para cada uno de los 5 canales de venta. Al listar productos en el POS, sólo se muestran los activos y habilitados para el canal activo de la caja.

**Condición:** Al cargar el menú de productos en cualquier módulo de venta.

**Resultado:** Sólo productos activos y visibles en ese canal aparecen en pantalla.

**Excepciones:** Si el producto está inactivo (lActivo = 0), no se muestra en ningún canal.

**Destino .NET:** `ProductoVisibilidadService` (`Domain/Services/ProductoVisibilidadService.cs`); delegado a `ProductoMaestro.DisponibleEnCanal()`

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Domain/Services/ProductoVisibilidadService.cs` + tests en `Inforest.Domain.Tests`

---

### BR-017
**Nombre:** InventoryGateway — descargo de almacén sobre ventas

**Origen:** Legacy/clsAlmacen.cls

**Archivo:** `legacy-restaurant/restaurant-vb6/Clases/clsAlmacen.cls`

**Procedimiento/Función:** `FunInsertamSubKardex`, `FunInsertamKardex`, `FunInsertaLogDescargo`, `EjecutarDescargoAutomatico`

**Descripción:** Al cerrar/cobrar un pedido de venta, si el sistema tiene almacén activo (lAlmacen = true en TPARAMETRO), se ejecuta el descargo de insumos/recetas contra la base ALMACEN. El SP `usp_Inforest_DescargoVenta` llena una tabla temporal con los ítems del pedido y sus recetas, y luego los descarga del kardex. Soporta descargo por receta directa y por insumos. Se registra un log de descargo por cada proceso.

**Condición:** `TPARAMETRO.lAlmacen = true` y pedido en estado 'CO' (cobrado) o cierre de lote.

**Resultado:** Movimiento en kardex/sub-kardex de almacén por la cantidad vendida × receta.

**Excepciones:** Si lAlmacen = false, se omite el descargo. Errores se registran en log de descargo.

**Destino .NET:** `IInventoryGateway` + `InventoryGateway` (`Infrastructure/Almacen/InventoryGateway.cs`)

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Application/Interfaces/IInventoryGateway.cs`, `modern-net8/src/Inforest.Infrastructure/Almacen/InventoryGateway.cs`

---

### BR-018
**Nombre:** NotificacionEmail — envío de correos electrónicos del sistema

**Origen:** Legacy/claCorreoElectronico.cls

**Archivo:** `legacy-restaurant/restaurant-vb6/Clases/claCorreoElectronico.cls`

**Procedimiento/Función:** `EnviarCorreo`, `EnviarCorreoProrroga`

**Descripción:** El sistema envía correos vía SMTP (Chilkat COM en Legacy) para: confirmaciones de reserva, recordatorios, agradecimientos y solicitudes de prórroga de licencia. Los parámetros SMTP provienen de configuración. Múltiples destinatarios separados por ";" se normalizan a ",".

**Condición:** Según flags en TPARAMETRO (lEmailConfirmacion, lEmailRecordatorio, lEmailAgradecimiento) o solicitud de prórroga de licencia.

**Resultado:** Correo enviado al destinatario configurado.

**Excepciones:** Error de SMTP es capturado y registrado en log; no corta el flujo principal.

**Destino .NET:** `INotificacionEmailService` + `SmtpEmailService` (`Infrastructure/Notifications/SmtpEmailService.cs`)

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Application/Interfaces/INotificacionEmailService.cs`, `modern-net8/src/Inforest.Infrastructure/Notifications/SmtpEmailService.cs`

---

### BR-DC-001
**Nombre:** No se puede vender sin aperturar el Día Contable

**Origen:** Legacy/frmDiaContable.frm

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmDiaContable.frm`

**Procedimiento/Función:** `cmdOpcion_Click`

**Descripción:** Si el usuario intenta salir del formulario de día contable sin haber aperturado ninguno, el sistema muestra el mensaje "No puede Iniciar Venta sin Aperturar Dia Contable" y bloquea el inicio del POS (`lIniciaPorDiaContable = False`).

**Condición:** Usuario hace clic en "Salir" sin haber aperturado el día contable.

**Resultado:** Mensaje de error. El flag `lIniciaPorDiaContable` permanece `False`, bloqueando el POS.

**Excepciones:** Si ya existe un día contable activo (modo apertura con fecha ≠ sentinel), el flag se establece en `True` automáticamente.

**Destino .NET:** `FrmDiaContable.OnSalirClick` → `IniciaPorDiaContable` property

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Desktop/Turno/FrmDiaContable.cs`

---

### BR-DC-002
**Nombre:** No se puede aperturar un nuevo Día Contable si ya existe uno activo

**Origen:** Legacy/frmDiaContable.frm

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmDiaContable.frm`

**Procedimiento/Función:** `Form_Load` modo="Apertura"

**Descripción:** Si `obtieneDiaContable()` retorna una fecha distinta al sentinel (1900/01/01), existe un día contable activo. El formulario muestra el mensaje "Para cerrar este Dia Contable se debe cerrar el turno" y deshabilita el botón de apertura.

**Condición:** `fdiacontable ≠ "19000101"` en modo "Apertura".

**Resultado:** Botón apertura oculto. Mensaje informativo. `lIniciaPorDiaContable = True`, `lDiaContableAperturado = True`.

**Excepciones:** —

**Destino .NET:** `FrmDiaContable.OnLoadAsync` modo `Apertura`

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Desktop/Turno/FrmDiaContable.cs`

---

### BR-DC-003
**Nombre:** La fecha de apertura del Día Contable no puede ser menor al último registrado

**Origen:** Legacy/frmDiaContable.frm + clsDiaContable.cls

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmDiaContable.frm`, `legacy-restaurant/restaurant-vb6/Clases/clsDiaContable.cls`

**Procedimiento/Función:** `cmdApertura_Click` → `validaFechaIngreso()`

**Descripción:** `validaFechaIngreso()` consulta `SELECT max(fdiacontable) FROM tdiacontable`. Si la fecha seleccionada es menor a esa máxima, la apertura es rechazada con el mensaje "La fecha seleccionada es menor al último Día Contable registrado".

**Condición:** `Format(dtpDiaContable.value,"yyyyMMdd") < Format(fechaMaxima,"yyyyMMdd")`.

**Resultado:** Apertura rechazada. Foco vuelve al DateTimePicker.

**Excepciones:** Si no hay registros previos (`max = NULL`), la validación pasa.

**Destino .NET:** `AperturarDiaContableHandler` + `IDiaContableService.ObtenerFechaMaximaDiaContableAsync`

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Application/Turno/DiaContableHandlers.cs`, `modern-net8/tests/Inforest.Application.Tests/Turno/DiaContableHandlerTests.cs`

---

### BR-DC-004
**Nombre:** El cierre del Día Contable actualiza directamente TDIACONTABLE

**Origen:** Legacy/frmDiaContable.frm

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmDiaContable.frm`

**Procedimiento/Función:** `cmdCerrar_Click`

**Descripción:** El cierre ejecuta directamente: `UPDATE TDIACONTABLE SET lcierre=1, tusuariocierre=<usuario>, fregistrocierre=GETDATE() WHERE fdiacontable=<fecha>`. No llama a un SP dedicado. Requiere confirmación del usuario (MsgBox Sí/No).

**Condición:** Usuario confirma el cierre.

**Resultado:** `lcierre=1` en el registro de TDIACONTABLE correspondiente.

**Excepciones:** Si el usuario cancela el diálogo, no se realiza ninguna operación.

**Destino .NET:** `CerrarDiaContableHandler` + `IDiaContableService.CerrarDiaContableAsync` (SQL directo parametrizado)

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Application/Turno/DiaContableHandlers.cs`, `modern-net8/src/Inforest.Infrastructure/Turno/DiaContableService.cs`, `modern-net8/tests/Inforest.Application.Tests/Turno/DiaContableHandlerTests.cs`

---

## Reglas de Negocio — Notas de Crédito (POS-FUNC-006)

### BR-NC-001

**Nombre:** Total de nota de crédito mayor a cero

**Origen:** Legacy/frmNotaCreditoDetalle.frm

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmNotaCreditoDetalle.frm`

**Procedimiento/Función:** `cmdOpcion_Click(1)` — Grabar

**Descripción:** El total calculado (neto + impuesto1 + impuesto2 + impuesto3) de la nota de crédito debe ser estrictamente mayor a cero.

**Condición:** `nNCTotal <= 0`

**Resultado:** Error — "El valor de la Nota de crédito debe ser mayor a cero"

**Excepciones:** Ninguna

**Destino .NET:** `NotaCredito.Emitir` (lanza `DomainException` `NC_TOTAL_INVALIDO`)

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Domain/Entities/Ventas/NotaCredito.cs`, `modern-net8/tests/Inforest.Domain.Tests/Ventas/NotaCreditoTests.cs`

---

### BR-NC-002

**Nombre:** Impuestos de nota de crédito no negativos

**Origen:** Legacy/frmNotaCreditoDetalle.frm

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmNotaCreditoDetalle.frm`

**Procedimiento/Función:** `cmdOpcion_Click(1)` — Grabar

**Descripción:** Los tres campos de impuesto de la nota de crédito no pueden tomar valores negativos.

**Condición:** `nNCImp1 < 0 OR nNCImp2 < 0 OR nNCImp3 < 0`

**Resultado:** Error — "El valor Impuesto de la Nota de crédito no debe ser negativo"

**Excepciones:** Ninguna

**Destino .NET:** `NotaCredito.Emitir` (lanza `DomainException` `NC_IMPUESTO_NEGATIVO`)

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Domain/Entities/Ventas/NotaCredito.cs`

---

### BR-NC-003

**Nombre:** Total acumulado de notas de crédito no supera el total del documento

**Origen:** Legacy/frmNotaCreditoDetalle.frm

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmNotaCreditoDetalle.frm`

**Procedimiento/Función:** `cmdOpcion_Click(1)` — Grabar (verificación vía `MNOTACREDITO SUM`)

**Descripción:** La suma de todas las notas de crédito activas emitidas contra un documento (campo `nVenta`) más la nueva nota de crédito no puede exceder el total (`nVenta`) del documento origen.

**Condición:** `totalNCPrevio + totalNCNueva > documento.Total`

**Resultado:** Error — "El valor de la nota de crédito supera el saldo disponible del documento"

**Excepciones:** Ninguna

**Destino .NET:** `EmitirNotaCreditoHandler` (retorna `NC_EXCEDE_DOCUMENTO`)

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Application/Ventas/NotaCreditoHandlers.cs`, `modern-net8/tests/Inforest.Application.Tests/Ventas/NotaCreditoHandlersTests.cs`

---

### BR-NC-004

**Nombre:** Observación/motivo de nota de crédito obligatoria

**Origen:** Legacy/frmNotaCreditoDetalle.frm

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmNotaCreditoDetalle.frm`

**Procedimiento/Función:** `cmdOpcion_Click(1)` — Grabar

**Descripción:** El campo de observación/motivo de la nota de crédito no puede estar vacío.

**Condición:** `txtObservacion.Text = ""`

**Resultado:** Error — "Ingrese el Motivo de la Nota de Crédito"

**Excepciones:** Ninguna

**Destino .NET:** `NotaCredito.Emitir` (lanza `DomainException` `NC_OBSERVACION_REQUERIDA`)

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Domain/Entities/Ventas/NotaCredito.cs`

---

### BR-NC-005

**Nombre:** Autorización de supervisor para emitir nota de crédito

**Origen:** Legacy/frmNotaCreditoDetalle.frm

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmNotaCreditoDetalle.frm`

**Procedimiento/Función:** `cmdOpcion_Click(0)` — Agregar; verifica `Supervisor("27")`

**Descripción:** La emisión de una nota de crédito requiere autorización de un supervisor (permiso "27" en `TACCESO`).

**Condición:** `Supervisor("27") = False`

**Resultado:** Error — "Clave no permitida"

**Excepciones:** Ninguna

**Destino .NET:** Delegado a `IRbacService` — pendiente integración en `EmitirNotaCreditoHandler` (gap controlado)

**Estado:** IN_PROGRESS

**Evidencia:** `modern-net8/src/Inforest.Infrastructure/Security/RbacService.cs`

---

### BR-NC-006

**Nombre:** No se puede anular una nota de crédito ya anulada

**Origen:** Legacy/frmNotaCreditoDetalle.frm

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmNotaCreditoDetalle.frm`

**Procedimiento/Función:** Estado "ANULADO" deshabilita botones de edición/anulación

**Descripción:** Una nota de crédito con estado "AN" (anulado) no puede ser anulada nuevamente.

**Condición:** `Estado == "AN"`

**Resultado:** Error — "La nota de crédito ya se encuentra anulada"

**Excepciones:** Ninguna

**Destino .NET:** `NotaCredito.Anular` (lanza `DomainException` `NC_YA_ANULADA`)

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Domain/Entities/Ventas/NotaCredito.cs`

---

## POS-FUNC-015 — Mensajería cocina/KDS

### BR-MSGCOC-001

**Nombre:** Mensaje de cocina obligatorio, sin caracteres prohibidos y en mayúsculas

**Origen:** Legacy/frmMensajeCocinaDetalle.frm

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmMensajeCocinaDetalle.frm`

**Procedimiento/Función:** `cmdOpcion_Click(0)`, `ValidaStr`

**Descripción:** El mensaje a cocina es obligatorio, se trimmea, se almacena en mayúsculas y bloquea los caracteres `&`, `'` y `"`.

**Condición:** Mensaje vacío, longitud mayor a 95 o contiene `&`, `'` o `"`.

**Resultado:** Error de validación; no se inserta ni actualiza `TMENSAJECOCINA`.

**Excepciones:** Ninguna

**Destino .NET:** `MensajeCocina.Crear/Actualizar`

**Estado:** COMPLETED

**Evidencia:** `modern-net8/src/Inforest.Domain/Entities/Cocina/MensajeCocina.cs`

---

### BR-MSGCOC-002

**Nombre:** Correlativo anual de 8 caracteres para mensajes de cocina

**Origen:** Legacy/frmMensajeCocinaDetalle.frm

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmMensajeCocinaDetalle.frm`

**Procedimiento/Función:** Alta de mensaje nuevo

**Descripción:** El código del mensaje se genera como `yy + 6 dígitos`, reiniciando el correlativo por año.

**Condición:** Inserción de un nuevo mensaje de cocina.

**Resultado:** Se genera el siguiente código disponible antes de grabar.

**Excepciones:** Si no existen mensajes del año, inicia en `yy000001`.

**Destino .NET:** `MensajeCocinaRepository.ObtenerProximoCodigoAsync` + `AgregarMensajeCocinaHandler`

**Estado:** COMPLETED

**Evidencia:** `modern-net8/src/Inforest.Infrastructure/Kitchen/MensajeCocinaRepository.cs`, `modern-net8/src/Inforest.Application/Kitchen/MensajeCocinaHandlers.cs`

---

### BR-MSGCOC-003

**Nombre:** Máximo 30 mensajes activos de cocina

**Origen:** Legacy/frmMensajeCocinaDetalle.frm

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmMensajeCocinaDetalle.frm`

**Procedimiento/Función:** `cmdOpcion_Click(0)`

**Descripción:** El sistema no permite registrar o reactivar más de 30 mensajes activos simultáneamente.

**Condición:** `lActivo = True` y el conteo de activos alcanza 30.

**Resultado:** Se bloquea la operación y se informa el máximo permitido.

**Excepciones:** Al modificar, el mensaje actual se excluye del conteo.

**Destino .NET:** `AgregarMensajeCocinaHandler`, `ModificarMensajeCocinaHandler`

**Estado:** COMPLETED

**Evidencia:** `modern-net8/src/Inforest.Application/Kitchen/MensajeCocinaHandlers.cs`

---

### BR-MSGCOC-004

**Nombre:** Listado y mantenimiento de mensajes por caja/rango

**Origen:** Legacy/frmMensajeCocina.frm, Legacy/frmMensajeCocinaDetalle.frm

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmMensajeCocina.frm`, `legacy-restaurant/restaurant-vb6/Formularios/frmMensajeCocinaDetalle.frm`

**Procedimiento/Función:** `cmdProcesa_Click`, `cmdOpcion_Click(0..2)`

**Descripción:** El listado usa `USP_LISTARMENSAJES`; con caja informada filtra por caja y sin caja usa rango de fechas. Desde el mantenimiento se puede agregar, editar y eliminar mensajes.

**Condición:** Consulta, edición o eliminación desde los formularios de mensajería.

**Resultado:** Se refleja el estado actual de `TMENSAJECOCINA` y se ejecutan los SPs legacy correspondientes.

**Excepciones:** Ninguna

**Destino .NET:** `FrmMensajeCocina`, `FrmMensajeCocinaDetalle`, `MensajeCocinaRepository`

**Estado:** COMPLETED

**Evidencia:** `modern-net8/src/Inforest.Desktop/Kitchen/FrmMensajeCocina.cs`, `modern-net8/src/Inforest.Desktop/Kitchen/FrmMensajeCocinaDetalle.cs`, `modern-net8/src/Inforest.Infrastructure/Kitchen/MensajeCocinaRepository.cs`

---

### BR-MSGCOC-005

**Nombre:** Cierre de mensajes activos al cerrar turno

**Origen:** Legacy/frmLiquidacionDetalle.frm

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmLiquidacionDetalle.frm`

**Procedimiento/Función:** Flujo de cierre de turno

**Descripción:** Al cerrar turno, los mensajes activos de la caja se cierran con usuario y fecha final mediante `USP_CERRAR_MENSAJES_CIERRETURNO`.

**Condición:** Cierre exitoso de turno para una caja.

**Resultado:** Los mensajes activos de `TMENSAJECOCINA` quedan inactivos con trazabilidad de cierre.

**Excepciones:** Si no hay mensajes activos en la caja, el SP no realiza cambios.

**Destino .NET:** `CerrarTurnoHandler` + `MensajeCocinaRepository.CerrarActivosPorCajaAsync`

**Estado:** COMPLETED

**Evidencia:** `modern-net8/src/Inforest.Application/Turno/TurnoHandlers.cs`, `modern-net8/src/Inforest.Infrastructure/Kitchen/MensajeCocinaRepository.cs`

---

## POS-FUNC-010 — Cliente y cuentas corrientes

### BR-CLIENTE-001
**Nombre:** Unicidad de código de cliente

**Origen:** Legacy/frmNuevoCliente.frm

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmNuevoCliente.frm`

**Procedimiento/Función:** `cmdGuardar_Click` / `RegistrarNuevoClientePosHandler`

**Descripción:** El código de cliente (tCodigoCliente) debe ser único en TCLIENTE.

**Condición:** Al insertar un nuevo cliente desde `FrmNuevoCliente`

**Resultado:** Si el código ya existe → Error "CLIENTE_YA_EXISTE"

**Excepciones:** Ninguna

**Destino .NET:** `RegistrarNuevoClientePosHandler`

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Application/Maestros/CuentaCorrienteHandlers.cs`

---

### BR-CLIENTE-002
**Nombre:** Validación de formato de identidad tributaria

**Origen:** Legacy/frmNuevoCliente.frm

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmNuevoCliente.frm`

**Procedimiento/Función:** `cmdValidaDNI_Click` / `Cliente.Crear`

**Descripción:** RUC debe tener 11 dígitos numéricos; DNI debe tener 8 dígitos numéricos. Otros tipos requieren al menos 4 caracteres.

**Condición:** Al crear o actualizar cliente

**Resultado:** Error con código `CLIENTE_RUC_INVALIDO` o `CLIENTE_DNI_INVALIDO` o `CLIENTE_IDENTIDAD_INVALIDA`

**Excepciones:** Tipos distintos de RUC/DNI usan validación mínima de longitud

**Destino .NET:** `Cliente.Crear` / `Cliente.Actualizar`

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Domain/Entities/Maestros/Cliente.cs`

---

### BR-CLIENTE-003
**Nombre:** Razón social obligatoria

**Origen:** Legacy/frmNuevoCliente.frm

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmNuevoCliente.frm`

**Procedimiento/Función:** Validación previa a guardado

**Descripción:** El campo razón social / empresa es obligatorio para crear o actualizar un cliente.

**Condición:** Al crear o actualizar cliente

**Resultado:** Error "CLIENTE_EMPRESA_REQUERIDA"

**Excepciones:** Ninguna

**Destino .NET:** `Cliente.Crear`

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Domain/Entities/Maestros/Cliente.cs`

---

### BR-CLIENTE-004
**Nombre:** Edición requiere cliente existente

**Origen:** Legacy/frmNuevoCliente.frm (modo edición)

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmNuevoCliente.frm`

**Procedimiento/Función:** `ActualizarClientePosHandler`

**Descripción:** Solo se puede actualizar un cliente que ya existe en TCLIENTE.

**Condición:** Al actualizar cliente desde `FrmNuevoCliente` en modo edición

**Resultado:** Error "CLIENTE_NO_ENCONTRADO" si no existe

**Excepciones:** Ninguna

**Destino .NET:** `ActualizarClientePosHandler`

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Application/Maestros/CuentaCorrienteHandlers.cs`

---

### BR-CTACTE-001
**Nombre:** Consumo no puede superar línea de crédito

**Origen:** Legacy/frmCtaCte.frm + vCompania

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmCtaCte.frm`

**Procedimiento/Función:** `CuentaCorriente.AplicarConsumo`

**Descripción:** El consumo acumulado en cuenta corriente no puede superar la línea de crédito asignada (TDELIVERY.nLinea).

**Condición:** Al aplicar consumo a una cuenta corriente

**Resultado:** Error "CTACTE_SALDO_INSUFICIENTE"

**Excepciones:** Ninguna

**Destino .NET:** `CuentaCorriente.AplicarConsumo`

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Domain/Entities/Maestros/CuentaCorriente.cs`

---

### BR-CTACTE-002
**Nombre:** Estado '03' = documento ctacte pendiente de cobro

**Origen:** Legacy/frmCuentaCobrar.frm

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmCuentaCobrar.frm`

**Procedimiento/Función:** `cmdProcesa_Click` — filtra `vDocumentoGrilla WHERE tEstadoDocumento='03'`

**Descripción:** Los documentos con estado '03' representan deuda por cuenta corriente pendiente de cobro.

**Condición:** Al consultar documentos pendientes de cobro en `FrmCuentaCobrar`

**Resultado:** Lista de documentos en estado '03' en el rango de fechas indicado

**Excepciones:** FechaFin no puede ser anterior a FechaInicio → error "CTACTE_FECHA_INVALIDA"

**Destino .NET:** `ObtenerDocumentosPendientesCobroHandler` + `CuentaCorrienteRepository`

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Application/Maestros/CuentaCorrienteHandlers.cs`, `modern-net8/src/Inforest.Infrastructure/Maestros/CuentaCorrienteRepository.cs`

---

### BR-CTACTE-003
**Nombre:** Solo clientes con ctacte habilitada aparecen en correlativo

**Origen:** Legacy/frmCtaCte.frm + vCompania

**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmCtaCte.frm`

**Procedimiento/Función:** `Form_Load` — carga `vCompania` (TDELIVERY WHERE lActivo=1 AND lClienteCtaCte=1)

**Descripción:** El correlativo de cuentas corrientes solo muestra registros activos con ctacte habilitada.

**Condición:** Al cargar `FrmCtaCte`

**Resultado:** Solo se retornan registros de TDELIVERY con lActivo=1 y lClienteCtaCte=1

**Excepciones:** Ninguna

**Destino .NET:** `ObtenerCuentasCorrientesHandler` + `CuentaCorrienteRepository`

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Application/Maestros/CuentaCorrienteHandlers.cs`, `modern-net8/src/Inforest.Infrastructure/Maestros/CuentaCorrienteRepository.cs`

---

## BR-RESERVA-001
**Nombre:** Estado inicial de reserva = Pendiente

**Origen:** Legacy/frmReservaDetalle.frm

**Archivo:** legacy-restaurant/restaurant-vb6/Formularios/frmReservaDetalle.frm

**Procedimiento/Función:** cmdOpcion_Click Case 1 (Sw=True)

**Descripción:** Al crear una reserva, el campo tEstadoReserva se inicializa a '01' (Pendiente).

**Condición:** INSERT INTO TRESERVA con tEstadoReserva='01'

**Resultado:** Reserva creada con estado Pendiente

**Excepciones:** Ninguna

**Destino .NET:** `Reserva.Crear()` — `EstadoReserva.Pendiente`

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Domain/Entities/Reservas/Reserva.cs`

---

## BR-RESERVA-002
**Nombre:** Cliente requerido en reserva

**Origen:** Legacy/frmReservaDetalle.frm

**Archivo:** legacy-restaurant/restaurant-vb6/Formularios/frmReservaDetalle.frm

**Procedimiento/Función:** cmdOpcion_Click Case 1 — validación previa

**Descripción:** No se puede crear ni modificar una reserva sin indicar el cliente.

**Condición:** `If txtCliente.Text = "" Then MsgBox "Ingrese el Cliente"`

**Resultado:** Error si el cliente no fue indicado

**Excepciones:** Ninguna

**Destino .NET:** `Reserva.Crear()` — DomainException RESERVA_CLIENTE_REQUERIDO

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Domain/Entities/Reservas/Reserva.cs`

---

## BR-RESERVA-003
**Nombre:** PAX requerido y mayor a cero

**Origen:** Legacy/frmReservaDetalle.frm

**Archivo:** legacy-restaurant/restaurant-vb6/Formularios/frmReservaDetalle.frm

**Procedimiento/Función:** cmdOpcion_Click Case 1 — validación previa

**Descripción:** El número de PAX debe ser mayor a cero al crear o modificar una reserva.

**Condición:** `If txtPax.Text = "" Then MsgBox "Ingrese el N° de Pax"`

**Resultado:** Error si PAX es vacío o cero

**Excepciones:** Ninguna

**Destino .NET:** `Reserva.Crear()` / `Reserva.Modificar()` — DomainException RESERVA_PAX_INVALIDO

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Domain/Entities/Reservas/Reserva.cs`

---

## BR-RESERVA-004
**Nombre:** Estado controla modificación/anulación/atención de reservas

**Origen:** Legacy/frmReservaDetalle.frm

**Archivo:** legacy-restaurant/restaurant-vb6/Formularios/frmReservaDetalle.frm

**Procedimiento/Función:** cmdOpcion_Click Case 1 (verificación EstadoReserva2), Case 2, Case 4

**Descripción:** No se puede modificar ni atender una reserva que ya fue Atendida ('02') o Anulada ('03'). No se puede anular una reserva ya Anulada. Solo se puede convertir a pedido una reserva Pendiente ('01').

**Condición:** EstadoReserva='02' → bloqueado para modificar y atender. EstadoReserva='03' → bloqueado para todo. EstadoReserva='01' → permitido atender via spIns_MPEDIDO_RESERVA.

**Resultado:** DomainException con código según la transición inválida

**Excepciones:** Ninguna

**Destino .NET:** `Reserva.Modificar()`, `Reserva.Anular()`, `Reserva.MarcarAtendida()` — DomainException RESERVA_YA_ATENDIDA / RESERVA_YA_ANULADA / RESERVA_NO_PENDIENTE

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Domain/Entities/Reservas/Reserva.cs`, `modern-net8/src/Inforest.Application/Reservas/ReservaHandlers.cs`

---

## BR-INSUMO-001
**Nombre:** Descripción requerida y no duplicada para insumo

**Origen:** Legacy/frmInsumoDetalle.frm

**Archivo:** legacy-restaurant/restaurant-vb6/Formularios/frmInsumoDetalle.frm

**Procedimiento/Función:** cmdOpcion_Click Case 1 (Grabar)

**Descripción:** Al agregar o modificar un insumo, la descripción no puede estar vacía. Al agregar, no puede existir ya un insumo con la misma descripción (comparación con UPPER).

**Condición:** txtMensaje.Text = "" → error. Descripcion duplicada → error.

**Resultado:** MsgBox de error, no se ejecuta el SP.

**Excepciones:** Ninguna

**Destino .NET:** `Insumo.Crear()` — DomainException INSUMO_DESCRIPCION_REQUERIDA. `AgregarInsumoHandler` — Result.Fail INSUMO_DESCRIPCION_DUPLICADA.

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Domain/Entities/Maestros/Insumo.cs`, `modern-net8/src/Inforest.Application/Maestros/InsumoHandlers.cs`

---

## BR-INSUMO-002
**Nombre:** Descripción de insumo en mayúsculas

**Origen:** Legacy/frmInsumoDetalle.frm

**Archivo:** legacy-restaurant/restaurant-vb6/Formularios/frmInsumoDetalle.frm

**Procedimiento/Función:** cmdOpcion_Click Case 1 (txtMensaje.Text = UCase(txtMensaje.Text))

**Descripción:** La descripción del insumo se normaliza a mayúsculas antes de guardar.

**Condición:** Siempre al agregar.

**Resultado:** La descripción queda almacenada en mayúsculas en TINSUMO.

**Excepciones:** Ninguna

**Destino .NET:** `Inforest.Crear()` llama `.ToUpperInvariant()` en la descripción.

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Domain/Entities/Maestros/Insumo.cs`

---

## BR-INSUMO-003
**Nombre:** Modificación de insumo actualiza descripción, stock, tipo, activo

**Origen:** Legacy/frmInsumoDetalle.frm

**Archivo:** legacy-restaurant/restaurant-vb6/Formularios/frmInsumoDetalle.frm

**Procedimiento/Función:** cmdOpcion_Click Case 1 (Sw = False → USP_MODIFICARINSUMOS)

**Descripción:** Al modificar, se actualizan: descripción, usuario, caja, activo, nstock (solo si modulo=INFOREST), LINSUMO y fecha de modificación.

**Condición:** Insumo ya existe. Sw=False (edición).

**Resultado:** SP USP_MODIFICARINSUMOS ejecutado con los nuevos valores.

**Excepciones:** Ninguna

**Destino .NET:** `Insumo.Actualizar()` + `ModificarInsumoHandler` → `InsumoRepository.ModificarAsync` → USP_MODIFICARINSUMOS.

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Domain/Entities/Maestros/Insumo.cs`, `modern-net8/src/Inforest.Application/Maestros/InsumoHandlers.cs`, `modern-net8/src/Inforest.Infrastructure/Maestros/InsumoRepository.cs`

---

## BR-INSUMO-004
**Nombre:** Eliminación de insumo previa confirmación

**Origen:** Legacy/frmInsumoDetalle.frm

**Archivo:** legacy-restaurant/restaurant-vb6/Formularios/frmInsumoDetalle.frm

**Procedimiento/Función:** cmdOpcion_Click Case 2 (USP_ELIMINARINSUMOS)

**Descripción:** Se solicita confirmación del usuario antes de eliminar un insumo mediante USP_ELIMINARINSUMOS.

**Condición:** Registro existe.

**Resultado:** SP USP_ELIMINARINSUMOS ejecutado. Grilla recargada.

**Excepciones:** Si no existe, no se elimina.

**Destino .NET:** `EliminarInsumoHandler` → `IInsumoRepository.EliminarAsync` → USP_ELIMINARINSUMOS. WinForm muestra MessageBox de confirmación.

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Application/Maestros/InsumoHandlers.cs`, `modern-net8/src/Inforest.Infrastructure/Maestros/InsumoRepository.cs`, `modern-net8/src/Inforest.Desktop/Maestros/FrmInsumo.cs`

---

## BR-IMPORT-001

**Nombre:** Solo requerimientos aprobados de áreas habilitadas pueden importarse

**Origen:** Legacy/frmImportacionRequerimientos.frm

**Archivo:** legacy-restaurant/restaurant-vb6/Formularios/frmImportacionRequerimientos.frm

**Procedimiento/Función:** cmdProcesa_Click

**Descripción:** Solo se muestran/importan requerimientos con `CodEstado='02'` (aprobado) cuya área esté habilitada para el punto de venta (`TRUTAAREA.lImportarPV=1`).

**Condición:** vRequerimiento.CodEstado = '02' AND TRUTAAREA.lImportarPV = 1

**Resultado:** Lista de requerimientos disponibles para importación.

**Excepciones:** Requerimientos en otros estados o de áreas no habilitadas no aparecen.

**Destino .NET:** `ObtenerRequerimientosPendientesHandler` → `IRequerimientoAlmacenRepository.ObtenerPendientesAsync`. Entidad `RequerimientoAlmacen.PuedeImportarse()`.

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Application/Almacen/ImportacionRequerimientoHandlers.cs`, `modern-net8/src/Inforest.Infrastructure/Almacen/RequerimientoAlmacenRepository.cs`, `modern-net8/src/Inforest.Domain/Entities/Almacen/RequerimientoAlmacen.cs`

---

## BR-IMPORT-002

**Nombre:** Un requerimiento no puede importarse dos veces

**Origen:** Legacy/frmImportacionRequerimientos.frm

**Archivo:** legacy-restaurant/restaurant-vb6/Formularios/frmImportacionRequerimientos.frm

**Procedimiento/Función:** cmdOpcion_Click Case 2

**Descripción:** Antes de crear el pedido, se verifica que `MREQUERIMIENTO.tPedido` esté vacío. Si ya tiene un código de pedido, se muestra error y se cancela la operación.

**Condición:** ISNULL(tPedido,'') = ''

**Resultado:** Si ya existe pedido: error y cancelación. Si no: continúa el flujo de importación.

**Excepciones:** Ninguna.

**Destino .NET:** `ImportarRequerimientoHandler` verifica `ObtenerPedidoAsociadoAsync` antes de crear el pedido. Retorna `REQ_YA_IMPORTADO` si ya existe.

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Application/Almacen/ImportacionRequerimientoHandlers.cs`

---

## BR-IMPORT-003

**Nombre:** Todos los artículos del requerimiento deben tener enlace con INFOREST

**Origen:** Legacy/frmImportacionRequerimientos.frm

**Archivo:** legacy-restaurant/restaurant-vb6/Formularios/frmImportacionRequerimientos.frm

**Procedimiento/Función:** InsertaProducto / cmdOpcion_Click Case 2

**Descripción:** Para cada artículo del requerimiento (CodArt), se busca el código de producto INFOREST mediante `TPRODUCTO.tEnlace`. Si un artículo no tiene enlace (`codProducto = ""` o `"0"`), el pedido ya creado se cancela (estado '03') y se notifica el error.

**Condición:** codProducto de vProducto WHERE tEnlace = CodArt

**Resultado:** Si todos tienen enlace: pedido creado con todos los ítems. Si alguno no tiene: pedido cancelado.

**Excepciones:** Si el pedido ya fue creado (spIns_MPEDIDO ejecutado), se cancela con estado '03'.

**Destino .NET:** `ImportarRequerimientoHandler` valida enlaces antes de crear el pedido (retorna `REQ_PRODUCTO_SIN_ENLACE`). `ImportacionPedidoGateway` cancela el pedido si el producto no existe en BD.

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Application/Almacen/ImportacionRequerimientoHandlers.cs`, `modern-net8/src/Inforest.Infrastructure/Almacen/ImportacionPedidoGateway.cs`

---

## BR-IMPORT-004

**Nombre:** El requerimiento debe marcarse como importado tras generar el pedido

**Origen:** Legacy/frmImportacionRequerimientos.frm

**Archivo:** legacy-restaurant/restaurant-vb6/Formularios/frmImportacionRequerimientos.frm

**Procedimiento/Función:** cmdOpcion_Click Case 2 (tras éxito de inserción)

**Descripción:** Tras crear exitosamente el pedido y todos sus ítems, se actualiza `MREQUERIMIENTO` con `lPedido=1` y `tPedido=<código pedido>` para bloquear re-importaciones.

**Condición:** Importación exitosa de todos los ítems.

**Resultado:** `MREQUERIMIENTO.lPedido = 1`, `MREQUERIMIENTO.tPedido = <código>`.

**Excepciones:** Si la creación de ítems falla, no se marca (el pedido se cancela y se retorna error).

**Destino .NET:** `ImportarRequerimientoHandler` llama `IRequerimientoAlmacenRepository.MarcarImportadoAsync` solo si el gateway retorna éxito.

**Estado:** MIGRATED

**Evidencia:** `modern-net8/src/Inforest.Application/Almacen/ImportacionRequerimientoHandlers.cs`, `modern-net8/src/Inforest.Infrastructure/Almacen/RequerimientoAlmacenRepository.cs`

---

## BR-RECIBO-001
**Nombre:** Correlativo de ingreso (formato año + secuencia)
**Origen:** Legacy/frmReciboIngresoDetalle.frm
**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmReciboIngresoDetalle.frm`
**Procedimiento/Función:** Botón Guardar / generación de tCodigoIngreso
**Descripción:** El correlativo de ingreso se forma concatenando el año de 2 dígitos + número de secuencia en 8 dígitos con padding de ceros (ej: `2600000001`).
**Condición:** Al registrar un nuevo ingreso.
**Resultado:** `tCodigoIngreso` = Right(Year(Date),2) + Format(secuencia,"00000000")
**Excepciones:** Ninguna.
**Destino .NET:** `ReciboIngreso.Registrar()` + `ReciboIngresoRepository.ObtenerSiguienteSecuenciaAsync()`
**Estado:** MIGRATED

---

## BR-RECIBO-002
**Nombre:** Estado inicial de ingreso = EMITIDO
**Origen:** Legacy/frmReciboIngresoDetalle.frm
**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmReciboIngresoDetalle.frm`
**Procedimiento/Función:** INSERT INTO MINGRESO
**Descripción:** Todo ingreso se crea con estado `01` (EMITIDO).
**Condición:** Al crear el registro.
**Resultado:** `tEstado = '01'`
**Excepciones:** Ninguna.
**Destino .NET:** `ReciboIngreso.Registrar()` fija `Estado = "01"`
**Estado:** MIGRATED

---

## BR-RECIBO-003
**Nombre:** Fecha y hora de registro = ahora
**Origen:** Legacy/frmReciboIngresoDetalle.frm
**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmReciboIngresoDetalle.frm`
**Procedimiento/Función:** INSERT INTO MINGRESO
**Descripción:** `dFechaIngreso` y `tHoraIngreso` se toman del momento de la operación.
**Condición:** Siempre al registrar.
**Resultado:** Campos de fecha/hora = DateTime.Now del sistema.
**Excepciones:** Ninguna.
**Destino .NET:** `ReciboIngreso.Registrar()` asigna `FechaIngreso = DateTime.Now`
**Estado:** MIGRATED

---

## BR-RECIBO-004
**Nombre:** Monto de ingreso > 0
**Origen:** Legacy/frmReciboIngresoDetalle.frm
**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmReciboIngresoDetalle.frm`
**Procedimiento/Función:** Validación antes del INSERT
**Descripción:** No se permite registrar un ingreso con monto cero o negativo.
**Condición:** `dMontoIngreso <= 0`
**Resultado:** Error de validación, registro no persiste.
**Excepciones:** Ninguna.
**Destino .NET:** `ReciboIngreso.Registrar()` lanza `DomainException` si monto ≤ 0
**Estado:** MIGRATED

---

## BR-RECIBO-005
**Nombre:** Anulación de ingreso cambia estado a 04
**Origen:** Legacy/frmReciboIngreso.frm
**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmReciboIngreso.frm`
**Procedimiento/Función:** Botón Anular
**Descripción:** Al anular un ingreso, su estado cambia de `01` (EMITIDO) a `04` (ANULADO). No se permite anular si ya está en `04`.
**Condición:** `tEstado = '01'` → puede anular; `tEstado = '04'` → ya anulado.
**Resultado:** `UPDATE MINGRESO SET tEstado='04'`
**Excepciones:** Ninguna.
**Destino .NET:** `ReciboIngreso.Anular()` + `AnularIngresoHandler`
**Estado:** MIGRATED

---

## BR-RECIBO-006
**Nombre:** Filtro multi-caja por usuario (lMCPV)
**Origen:** Legacy/frmReciboIngreso.frm
**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmReciboIngreso.frm`
**Procedimiento/Función:** Carga de grilla (SELECT FROM MINGRESO)
**Descripción:** Cuando el flag `lMCPV` (multi-caja punto de venta) está activo, la consulta filtra ingresos por el usuario de la sesión (`tUsuario`). Sin el flag se muestran todos.
**Condición:** `lMCPV = True` → añade `WHERE tUsuario = @usuario`
**Resultado:** Lista de ingresos filtrada.
**Excepciones:** Ninguna.
**Destino .NET:** `ObtenerIngresosQuery.CodigoUsuario` opcional en `ReciboIngresoRepository.ObtenerAsync()`
**Estado:** MIGRATED

---

## BR-RECIBO-007
**Nombre:** Correlativo de egreso (formato año + secuencia)
**Origen:** Legacy/frmReciboEgresoDetalle.frm
**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmReciboEgresoDetalle.frm`
**Procedimiento/Función:** Botón Guardar / generación de tCodigoEgreso
**Descripción:** El correlativo de egreso se forma igual que el de ingreso: año 2 dígitos + secuencia 8 dígitos.
**Condición:** Al registrar un nuevo egreso.
**Resultado:** `tCodigoEgreso` = Right(Year(Date),2) + Format(secuencia,"00000000")
**Excepciones:** Ninguna.
**Destino .NET:** `ReciboEgreso.Registrar()` + `ReciboEgresoRepository.ObtenerSiguienteSecuenciaAsync()`
**Estado:** MIGRATED

---

## BR-RECIBO-008
**Nombre:** Estado inicial de egreso = EMITIDO
**Origen:** Legacy/frmReciboEgresoDetalle.frm
**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmReciboEgresoDetalle.frm`
**Procedimiento/Función:** INSERT INTO MEGRESO
**Descripción:** Todo egreso se crea con estado `01` (EMITIDO). No existe estado PAGADO para egresos.
**Condición:** Al crear el registro.
**Resultado:** `tEstado = '01'`
**Excepciones:** Ninguna.
**Destino .NET:** `ReciboEgreso.Registrar()` fija `Estado = "01"`
**Estado:** MIGRATED

---

## BR-RECIBO-009
**Nombre:** Solicitud de clave supervisor para egresos
**Origen:** Legacy/frmReciboEgreso.frm
**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmReciboEgreso.frm`
**Procedimiento/Función:** Botón Nuevo/Editar → `lSolicitaClaveEgreso`
**Descripción:** Si el parámetro `lSolicitaClaveEgreso` está activo, antes de abrir el formulario de detalle de egreso se solicita contraseña de supervisor.
**Condición:** `lSolicitaClaveEgreso = True`
**Resultado:** Se solicita contraseña; si es correcta, se abre el detalle.
**Excepciones:** Si la contraseña es incorrecta, se cancela la acción.
**Destino .NET:** Verificación delegada a la capa de presentación (`FrmReciboEgreso`) antes de abrir `FrmReciboEgresoDetalle`.
**Estado:** MIGRATED

---

## BR-RECIBO-010
**Nombre:** Modo de egreso: VARIOS / NOTACREDITO / MERCADERIA
**Origen:** Legacy/frmReciboEgresoDetalle.frm
**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmReciboEgresoDetalle.frm`
**Procedimiento/Función:** Combo `cboModoEgreso`
**Descripción:** El egreso puede ser de tipo `VARIOS` (general), `NOTACREDITO` (requiere referencia) o `MERCADERIA` (compra de mercadería).
**Condición:** `tModoEgreso IN ('VARIOS','NOTACREDITO','MERCADERIA')`
**Resultado:** Permite registrar el egreso con el modo seleccionado.
**Excepciones:** Valor fuera de rango → error de validación.
**Destino .NET:** `ModoEgreso` enum + validación en `ReciboEgreso.Registrar()`
**Estado:** MIGRATED

---

## BR-RECIBO-011
**Nombre:** Anulación de egreso cambia estado a 04
**Origen:** Legacy/frmReciboEgreso.frm
**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmReciboEgreso.frm`
**Procedimiento/Función:** Botón Anular
**Descripción:** Al anular un egreso su estado pasa a `04` (ANULADO). No se permite anular un egreso ya anulado.
**Condición:** `tEstado = '01'` → puede anular.
**Resultado:** `UPDATE MEGRESO SET tEstado='04'`
**Excepciones:** Ninguna.
**Destino .NET:** `ReciboEgreso.Anular()` + `AnularEgresoHandler`
**Estado:** MIGRATED

---

## BR-RECIBO-012
**Nombre:** NotaCredito requiere campo dReferencia
**Origen:** Legacy/frmReciboEgresoDetalle.frm
**Archivo:** `legacy-restaurant/restaurant-vb6/Formularios/frmReciboEgresoDetalle.frm`
**Procedimiento/Función:** Validación antes del INSERT cuando tModoEgreso = 'NOTACREDITO'
**Descripción:** Cuando el modo de egreso es `NOTACREDITO`, el campo `dReferencia` (número de la nota de crédito) es obligatorio.
**Condición:** `tModoEgreso = 'NOTACREDITO' AND IsEmpty(dReferencia)`
**Resultado:** Error de validación; no se persiste.
**Excepciones:** Ninguna.
**Destino .NET:** Validación en `ReciboEgreso.Registrar()` — lanza `DomainException` si modo es NOTACREDITO y referencia vacía.
**Estado:** MIGRATED
