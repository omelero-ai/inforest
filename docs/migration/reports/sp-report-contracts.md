# Contratos SP de Reportes — INFOREST

> Última actualización: 2026-08-12
>
> Referencia: ADR-002 (SQL Server + Dapper), ADR-007 (FastReport .NET)
>
> **Decisión de retención:** Todos los `spRep_*` se mantienen en SQL Server sin modificación.
> Son envueltos por `IReporteRepository` / `ReporteRepository` en la capa Infrastructure.

---

## spRep_Comanda

**Origen Legacy:** `frmRepComanda.frm`
**Archivo SQL:** `legacy-restaurant/database-sql-server/5. SP.sql`
**Reglas:** BR-REP-001

### Parámetros de entrada

| Parámetro | Tipo SQL | Descripción |
|---|---|---|
| `@flagTipo` | `bit` | 1 = por ítem DPEDIDO, 0 = cabecera |
| `@SOrden` | `nvarchar(50)` | Columna de ordenamiento |
| `@fInicio` | `smalldatetime` | Fecha inicio |
| `@fFinal` | `smalldatetime` | Fecha fin |
| `@sCriterio` | `nvarchar(400)` | Cláusula WHERE dinámica adicional |

### Dataset retornado

| Columna | Tipo | Descripción |
|---|---|---|
| `tCodigoPedido` | `nvarchar(10)` | Código pedido |
| `tComanda` | `nvarchar(10/50)` | Número de comanda |
| `NombreProducto` | `nvarchar(60)` | Nombre producto |
| `Mozo` | `nvarchar(30)` | Nombre mozo |
| `nCantidad` | `float` | Cantidad |
| `PrecioUnitario` | `float` | Precio unitario |
| `PrecioTotal` | `float` | Precio total |
| `fFecha` | `smalldatetime` | Fecha |
| `Usuario` | `nvarchar(50)` | Usuario |
| `tDocumento` | `nvarchar(20)` | Documento |
| `Estado` | `nvarchar(1)` | Estado ítem |
| `tObservacion` | `nvarchar(250)` | Observación anulado |

### Tablas involucradas
`MPEDIDO`, `DPEDIDO`, `TPRODUCTO`, `vMozo`

### Crystal Report asociado
`dsrComandaD.Dsr`, `dsrComandaR.Dsr`

### FastReport Target
`RepComanda.frx`, `RepComandaDetallado.frx`

---

## spRep_Propina

**Origen Legacy:** `frmRepPropina.frm`
**Archivo SQL:** `legacy-restaurant/database-sql-server/5. SP.sql`
**Reglas:** BR-REP-002

### Parámetros de entrada

| Parámetro | Tipo SQL | Descripción |
|---|---|---|
| `@fInicio` | `smalldatetime` | Fecha inicio |
| `@fFinal` | `smalldatetime` | Fecha fin |
| `@sCondicion` | `nvarchar(400)` | Condición WHERE adicional dinámica |

### Dataset retornado

| Columna | Tipo | Descripción |
|---|---|---|
| `tMozo` | `nvarchar` | Código mozo |
| `tmotorizado` | `nvarchar` | Código motorizado |
| `tDetallado` | `nvarchar` | Tipo tarjeta |
| `Propina` | `float` | Monto propina |
| `tDocumento` | `nvarchar(20)` | Documento |
| `fRegistro` | `nvarchar` | Fecha formateada dd/mm/yyyy |
| `tCorrelativo` | `nvarchar` | Correlativo |
| `ttipopedido` | `nvarchar` | Tipo pedido |
| `Trabajador` | `nvarchar` | Nombre trabajador (mozo o motorizado) |
| `TipoPedido` | `nvarchar` | Descripción tipo pedido |
| `nfactorretencion` | `float` | Factor retención tarjeta |

### Tablas involucradas
`DPAGODOCUMENTO`, `MPEDIDO`, `DPEDIDO`, `TTARJETACREDITO`, `vMozo`, `vMotorizado`, `vTipoPedido`

### Crystal Report asociado
`dsrPropinaD.Dsr`, `dsrPropinaR.Dsr`

### FastReport Target
`RepPropina.frx`

---

## spRep_PrincipalCliente

**Origen Legacy:** `frmRepPrincipal.frm`
**Archivo SQL:** `legacy-restaurant/database-sql-server/5. SP.sql`
**Reglas:** BR-REP-003

### Parámetros de entrada

| Parámetro | Tipo SQL | Descripción |
|---|---|---|
| `@flagTipo` | `bit` | 1 = detalle por documento, 0 = resumen por cliente |
| `@sMonto` | `float` | Monto mínimo de consumo acumulado |
| `@SCliente` | `nvarchar(20)` | Filtro por código cliente (vacío = todos) |
| `@fInicio` | `smalldatetime` | Fecha inicio |
| `@fFinal` | `smalldatetime` | Fecha fin |

### Dataset retornado (modo detalle `@flagTipo=1`)

| Columna | Tipo | Descripción |
|---|---|---|
| `tCodigoCliente` | `nvarchar` | Código cliente |
| `tEmpresa` | `nvarchar` | Nombre empresa |
| `tDocumento` | `nvarchar` | Número documento |
| `fFecha` | `smalldatetime` | Fecha |
| `nNeto` | `float` | Neto |
| `nPrecioImpuesto1/2/3` | `float` | Impuestos |
| `nVenta` | `float` | Venta total |

### Dataset retornado (modo resumen `@flagTipo=0`)

| Columna | Tipo | Descripción |
|---|---|---|
| `tCodigoCliente` | `nvarchar` | Código cliente |
| `tEmpresa` | `nvarchar` | Nombre empresa |
| `Cantidad` | `int` | Cantidad documentos |
| `nNeto` | `float` | Suma neto |
| `nVenta` | `float` | Suma venta |

### Tablas involucradas
`MDOCUMENTO`, `TCLIENTE`

### Crystal Report asociado
`dsrPrincipalD.Dsr`, `dsrPrincipalR.Dsr`

### FastReport Target
`RepPrincipalCliente.frx`

---

## spRep_CtaCteIntegrado

**Origen Legacy:** `frmRepCtaCteIntegrado.frm`
**Archivo SQL:** `legacy-restaurant/database-sql-server/5. SP.sql`
**Reglas:** BR-REP-004

### Parámetros de entrada

| Parámetro | Tipo SQL | Descripción |
|---|---|---|
| `@flagTipo` | `nvarchar(2)` | `'1'`=consolidado, `'2'`=detallado, `'3'`=estado |
| `@fInicio` | `smalldatetime` | Fecha inicio |
| `@fFinal` | `smalldatetime` | Fecha fin |

### Tablas involucradas
`MPEDIDO`, `DPEDIDO`, `MDOCUMENTO`, `TCLIENTE`, `vCompania`, `vSalon`

### Crystal Report asociado
`dsrRepCtaCteIntegradoConsolidado.Dsr`, `dsrRepCtaCteIntegradoDetallado.Dsr`

### FastReport Target
`RepCtaCteIntegrado.frx`

---

## spRep_PaloteoComparativo

**Origen Legacy:** `frmRepPaloteoComparativo.frm`
**Archivo SQL:** `legacy-restaurant/database-sql-server/5. SP.sql`
**Reglas:** BR-REP-005

### Parámetros de entrada

| Parámetro | Tipo SQL | Descripción |
|---|---|---|
| `@flagTurnoOFecha` | `bit` | 0=por turno, 1=por rango fecha |
| `@flagTipoValor` | `bit` | 0=cantidad, 1=valor |
| `@tTipoProducto` | `nvarchar(20)` | Filtro tipo producto |
| `@tAreaProduccion` | `nvarchar(50)` | Filtro área producción |
| `@tGrupo` | `nvarchar(50)` | Filtro grupo |
| `@tSubGrupo` | `nvarchar(50)` | Filtro subgrupo |
| `@tCodigoProducto` | `nvarchar(20)` | Filtro código producto |
| `@tTurno` | `nvarchar(20)` | Turno (si aplica) |
| `@finicio` | `smalldatetime` | Fecha inicio |
| `@ffinal` | `smalldatetime` | Fecha fin |
| `@flagNFacturado` | `bit` | 1=solo facturados |

### Dataset retornado

| Columna | Tipo | Descripción |
|---|---|---|
| `Codigo` | `nvarchar(7)` | Código producto |
| `Grupo` | `nvarchar(50)` | Grupo |
| `SubGrupo` | `nvarchar(50)` | Subgrupo |
| `Producto` | `nvarchar(100)` | Nombre producto |
| `Valor` | `float` | Valor referencia |
| `Produccion` | `float` | Producción |
| `Venta` | `float` | Venta |
| `Cortesia` | `float` | Cortesía |
| `CtaCte` | `float` | Cuenta corriente |
| `Canal1`–`Canal5` | `float` | Canales (salon/delivery/llevar/canal4/canal5) |

### Tablas involucradas
`MPEDIDO`, `DPEDIDO`, `TPRODUCTO`, `TGRUPO`, `TSUBGRUPO`, `MDOCUMENTO`

### Crystal Report asociado
`dsrPaloteoComparativo.Dsr`

### FastReport Target
`RepPaloteoComparativo.frx`

---

## spRep_PaloteoSubProd

**Origen Legacy:** `frmRepPaloteoSubProd.frm`
**Archivo SQL:** `legacy-restaurant/database-sql-server/5. SP.sql`
**Reglas:** BR-REP-006

### Parámetros de entrada

| Parámetro | Tipo SQL | Descripción |
|---|---|---|
| `@flagTurno` | `bit` | 0=por turno, 1=por fecha |
| `@tTurno` | `nvarchar(30)` | Turno |
| `@tLocal` | `nvarchar(30)` | Local |
| `@tSalon` | `nvarchar(30)` | Salón |
| `@tipoProd` | `nvarchar(30)` | Tipo producto |
| `@tmozo` | `nvarchar(30)` | Mozo |
| `@ttipoPedido` | `nvarchar(30)` | Tipo pedido |
| `@tarea` | `nvarchar(30)` | Área |
| `@tcaja` | `nvarchar(30)` | Caja |
| `@tcodigoproducto` | `nvarchar(30)` | Código producto |
| `@tcodigocliente` | `nvarchar(30)` | Código cliente |
| `@sBoton2`–`@sBoton5` | `nvarchar(30)` | Filtros boton personalizados |
| `@fInicio` | `smalldatetime` | Fecha inicio |
| `@fFinal` | `smalldatetime` | Fecha fin |

### Dataset retornado

Columnas: `tLocal`, `Local`, `Salon`, `tMesa`, `TipoProducto`, `Grupo`, `SubGrupo`, `Producto`, `Cantidad`, `Venta`, `Pedido`, `Documento`, `Fecha`, `tTipoPedido`, `Area`, `subProducto`, `cantprod`

### Tablas involucradas
`MPEDIDO`, `DPEDIDO`, `vSalon`, `vPRODUCTOXPRODUCTO`, `MDOCUMENTO`

### Crystal Report asociado
`dsrPaloteoSPD.Dsr`

### FastReport Target
`RepPaloteoSubProd.frx`

---

## spRep_PaloteoVentaIntegrado

**Origen Legacy:** `frmRepPaloteoVentaIntegrado.frm`
**Archivo SQL:** `legacy-restaurant/database-sql-server/5. SP.sql`
**Reglas:** BR-REP-007

### Parámetros de entrada (seleccionados)

| Parámetro | Tipo SQL | Descripción |
|---|---|---|
| `@flagPropiedades` | `bit` | Incluir propiedades |
| `@flagProduccion` | `bit` | Incluir producción |
| `@flagVenta` | `bit` | Incluir venta |
| `@flagCortesia` | `bit` | Incluir cortesía |
| `@flagCuentaCte` | `bit` | Incluir cuenta corriente |
| `@flagPedidosFacturados` | `bit` | Solo facturados |
| `@flagCombinacion` | `bit` | Combinación |
| `@flagCargo` | `bit` | Incluir cargo |
| `@flagOpcion` | `bit` | 1=detallado, 0=resumido |
| `@flagVVenta` / `@flagVNeto` / `@flagVCosto` | `bit` | Modo precio |
| `@tTipoProducto` / `@tTipoPedido` / `@tGrupo` / `@tSubGrupo` | `nvarchar(30)` | Filtros |
| `@sPrecio` / `@sNeto` / `@sCosto` | `nvarchar(250)` | Expresión precio dinámica |
| `@sOrden` | `nvarchar(200)` | Orden |
| `@sBoton2`–`@sBoton5` | `nvarchar(30)` | Filtros botón |
| `@finicio` / `@ffinal` | `smalldatetime` | Rango fecha |
| (continúa con más parámetros de local/turno/salon...) | | |

### Tablas involucradas
`MPEDIDO`, `DPEDIDO`, `TPRODUCTO`, `TGRUPO`, `TSUBGRUPO`, `MDOCUMENTO`, `vSalon`

### Crystal Report asociado
`dsrPaloteoVentaIntegradol.Dsr`, `dsrPaloteoVentaIntegradoS.Dsr`, `dsrPaloteoVentaIntegradoComparativo.Dsr`

### FastReport Target
`RepPaloteoVentaIntegrado.frx`

---

## spRep_RankingIntegrado

**Origen Legacy:** `frmRepRankingIntegrado.frm`
**Archivo SQL:** `legacy-restaurant/database-sql-server/5. SP.sql`
**Reglas:** BR-REP-008

### Parámetros de entrada (seleccionados)

Misma estructura de flags que `spRep_PaloteoVentaIntegrado` más:

| Parámetro | Tipo SQL | Descripción |
|---|---|---|
| `@finicio` / `@ffinal` | `smalldatetime` | Rango fecha |
| `@flagProduccion`/`@flagVenta`/`@flagCortesia`/`@flagCuentaCte` | `bit` | Incluir categorías |
| `@flagPedidosFacturados` | `bit` | Solo facturados |
| `@tGrupo` / `@tSubGrupo` / `@tCodigoProducto` | `nvarchar` | Filtros |
| `@sPrecio` / `@sNeto` / `@sCosto` | `nvarchar(300)` | Expresión precio |

### Tablas involucradas
`MPEDIDO`, `DPEDIDO`, `TPRODUCTO`, `TGRUPO`, `TSUBGRUPO`, `MDOCUMENTO`, `vSalon`

### Crystal Report asociado
`dsrRankingIntegradoComparativo.Dsr`, `dsrRankingIntegradol.Dsr`

### FastReport Target
`RepRankingIntegrado.frx`

---

## spRep_VentaMensualIntegrado

**Origen Legacy:** `frmRepVentaMensualIntegrado.frm`, `frmRepDiarioVentaIntegrado.frm`
**Archivo SQL:** `legacy-restaurant/database-sql-server/5. SP.sql`
**Reglas:** BR-REP-009

### Parámetros de entrada

| Parámetro | Tipo SQL | Descripción |
|---|---|---|
| `@sPrecio` | `nvarchar(150)` | Expresión precio dinámica SQL |
| `@sAno` | `nvarchar(4)` | Año |
| `@sMes` | `nvarchar(4)` | Mes |
| `@sFecha` | `nvarchar(400)` | Condición fecha adicional |
| `@dHour` | `float` | Hora de corte del día (día contable) |
| `@tipo` | `nvarchar(1)` | `'D'`=diario, `'M'`=mensual |
| `@metaMensual` | `float` | Meta mensual de venta |
| `@diames` | `float` | Días del mes |

### Dataset retornado

| Columna | Tipo | Descripción |
|---|---|---|
| `Dia` | `int` | Día del mes |
| `Salon` | `float` | Venta salón |
| `Delivery` | `float` | Venta delivery |
| `Llevar` | `float` | Venta llevar |
| `Canal4` / `Canal5` | `float` | Otros canales |
| `Venta` | `float` | Venta total |
| `Cantidad` | `int` | Cantidad pedidos |
| `Pax` | `int` | Comensales |
| `TipoPedido` | `nvarchar(3)` | Tipo pedido |
| `costo` | `float` | Costo |

### Tablas involucradas
`MPEDIDO`, `DPEDIDO`, `MDOCUMENTO`

### Crystal Report asociado
`dsrVentaMensualIntegrada.Dsr`, `dsrVentaDiariaIntegrada.Dsr`

### FastReport Target
`RepVentaMensualIntegrado.frx`, `RepVentaDiariaIntegrado.frx`

---

## spRep_AnaliticoMotorizadoIntegrado

**Origen Legacy:** `frmRepAnaliticoMotorizadoIntegrado.frm`
**Archivo SQL:** `legacy-restaurant/database-sql-server/5. SP.sql`
**Reglas:** BR-REP-010

### Parámetros de entrada

| Parámetro | Tipo SQL | Descripción |
|---|---|---|
| `@flagProduccion` | `bit` | Incluir producción |
| `@flagVenta` | `bit` | Incluir venta |
| `@flagCortesia` | `bit` | Incluir cortesía |
| `@flagCuentaCte` | `bit` | Incluir cuenta corriente |
| `@flagPedidosFacturados` | `bit` | Solo facturados |
| `@flagCombinacion` | `bit` | Combinación |
| `@flagCargo` | `bit` | Cargo |
| `@sPrecio` | `nvarchar(150)` | Expresión precio dinámica |
| `@finicio` | `smalldatetime` | Fecha inicio |
| `@ffinal` | `smalldatetime` | Fecha fin |

### Dataset retornado

Columnas: `Motorizado`, `TipoProducto`, `Grupo`, `SubGrupo`, `Producto`, `Cantidad`, `Venta`, `nPedidos`, `tcodigopedido`, `Comision`

### Tablas involucradas
`MPEDIDO`, `DPEDIDO`, `TPRODUCTO`, `TGRUPO`, `TSUBGRUPO`, `MDOCUMENTO`, `vMotorizado`

### Crystal Report asociado
`dsrAnaliticoMotorizadoIntegradoConsolidado.Dsr`, `dsrAnaliticoMotorizadoIntegradoDetallado.Dsr`

### FastReport Target
`RepAnaliticoMotorizadoIntegrado.frx`

---

## USP_KDS_ResporteTiempoPedido

**Origen Legacy:** KDS monitor
**Archivo SQL:** `legacy-restaurant/database-sql-server/5. SP.sql`
**Reglas:** BR-REP-011

### Parámetros de entrada

| Parámetro | Tipo SQL | Descripción |
|---|---|---|
| `@fInicio` | `smalldatetime` | Fecha inicio |
| `@fFinal` | `smalldatetime` | Fecha fin |

### Dataset retornado

| Columna | Tipo | Descripción |
|---|---|---|
| `Pedido` | `nvarchar` | Código pedido |
| `Producto` | `nvarchar` | Nombre producto |
| `TiempoCorto` | `varchar(8)` | Tiempo mínimo preparación (HH:mm:ss) |
| `TiempoLargo` | `varchar(8)` | Tiempo máximo preparación |
| `TiempoPromedio` | `varchar(8)` | Tiempo promedio preparación |

### Tablas involucradas
`DPEDIDOKDS`, `DPEDIDO`, `TPRODUCTO`

### Crystal Report asociado
`dsrTiempoKdsP.Dsr`, `dsrTiempoKdsR.Dsr`

### FastReport Target
`RepTiempoKdsPedido.frx`

---

## USP_KDS_ResporteTiempoProducto

**Origen Legacy:** KDS monitor
**Archivo SQL:** `legacy-restaurant/database-sql-server/5. SP.sql`
**Reglas:** BR-REP-012

### Parámetros de entrada

| Parámetro | Tipo SQL | Descripción |
|---|---|---|
| `@fInicio` | `smalldatetime` | Fecha inicio |
| `@fFinal` | `smalldatetime` | Fecha fin |
| `@tGrupo` | `nvarchar(10)` | Filtro grupo |
| `@tSubGrupo` | `nvarchar(10)` | Filtro subgrupo |
| `@tProducto` | `nvarchar(30)` | Filtro producto |

### Dataset retornado

| Columna | Tipo | Descripción |
|---|---|---|
| `Grupo` | `nvarchar` | Nombre grupo |
| `SubGrupo` | `nvarchar` | Nombre subgrupo |
| `Producto` | `nvarchar` | Nombre producto |
| `TiempoCorto` | `varchar(8)` | Tiempo mínimo |
| `TiempoLargo` | `varchar(8)` | Tiempo máximo |
| `TiempoPromedio` | `varchar(8)` | Tiempo promedio |

### Tablas involucradas
`DPEDIDOKDS`, `DPEDIDO`, `TPRODUCTO`, `TGRUPO`, `TSUBGRUPO`

### Crystal Report asociado
`dsrTiempoKdsP.Dsr`, `dsrTiempoKdsR.Dsr`

### FastReport Target
`RepTiempoKdsProducto.frx`

---

## Nota sobre SPs con SQL Dinámico

Los SPs `spRep_Propina`, `spRep_PaloteoVentaIntegrado`, `spRep_RankingIntegrado`, `spRep_VentaMensualIntegrado` y `spRep_AnaliticoMotorizadoIntegrado` usan `EXEC sp_executesql` con SQL construido dinámicamente. El parámetro `@sPrecio` (y similares) contiene una expresión SQL como `'nVenta'` o `'nNeto'` que se inyecta en el SELECT. Esta lógica de construcción de expresión debe replicarse en el formulario VB6 → Formulario .NET equivalente antes de llamar al SP.

**Riesgo:** validar que el valor de `@sPrecio` no pueda ser manipulado (SQL injection) desde la UI. El formulario .NET debe usar un `enum` o constantes controladas, no texto libre del usuario.
