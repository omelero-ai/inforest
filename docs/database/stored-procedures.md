# Stored Procedures — INFOREST

> Status: CONFIRMED — inventario basado en análisis de `legacy-restaurant/database-sql-server/5. SP.sql`

---

## Resumen

| Métrica | Valor |
|---|---|
| Total SPs detectados | 105+ |
| SPs de reporte (`spRep_*`) | ~45 |
| SPs KDS (`USP_KDS_*`) | 5 |
| SPs de insumos (`USP_*`) | 5 |
| SPs de auditoría (`usp_Aud_*`) | ~3 |
| SPs operativos (`spIns_`, `spUpd_`, `sp_`) | ~15 |
| Funciones escalares | 2 |

---

## Inventario por Categoría

### Operativos — Pedidos y Documentos

| SP | Tablas Involucradas | Uso | Equivalente .NET | Estado |
|---|---|---|---|---|
| `spIns_MPEDIDO` | MPEDIDO, DPEDIDO | Inserción de pedido | — | NOT_STARTED |
| `spIns_TipoCambio` | TTIPOCAMBIO | Inserción tipo de cambio | — | NOT_STARTED |
| `spUpd_MPEDIDO` | MPEDIDO | Actualización de pedido | — | NOT_STARTED |

### Reportes — `spRep_*`

| SP | Módulo | Descripción | Equivalente .NET | Estado |
|---|---|---|---|---|
| `spRep_Anulacion` | Auditoría | Reporte de anulaciones | — | NOT_STARTED |
| `spRep_AnulacionDocumentoIntegrado` | Auditoría | Anulaciones integradas | — | NOT_STARTED |
| `spRep_AnulacionPedidoIntegrado` | Auditoría | Anulación pedidos integrado | — | NOT_STARTED |
| `spRep_AnaliticoMotorizado` | Delivery | Analítico de motorizados | — | NOT_STARTED |
| `spRep_AnaliticoMotorizadoIntegrado` | Delivery | Analítico motorizado integrado | — | NOT_STARTED |
| `spRep_AutorizacionAutoimpresion` | Caja | Autorización autoimpresión | — | NOT_STARTED |
| `spRep_Comanda` | Cocina | Reporte de comandas | — | NOT_STARTED |
| `spRep_ComprobantesVentas` | Ventas | Comprobantes de ventas | — | NOT_STARTED |
| `spRep_ControlDocumentos` | Ventas | Control de documentos | — | NOT_STARTED |
| `spRep_Cortesia` | Ventas | Reporte de cortesías | — | NOT_STARTED |
| `spRep_CtaCteN` | Clientes | Cuenta corriente | — | NOT_STARTED |
| `spRep_CtaCteIntegrado` | Clientes | Cta. corriente integrada | — | NOT_STARTED |
| `spRep_CuentasCobrar` | Clientes | Cuentas por cobrar | — | NOT_STARTED |
| `spRep_Descuento` | Ventas | Reporte de descuentos | — | NOT_STARTED |
| `spRep_Diferencia` | Caja | Diferencias de caja | — | NOT_STARTED |
| `spRep_Entregas` | Delivery | Reporte de entregas | — | NOT_STARTED |
| `spRep_FormaPagoIntegrado` | Caja | Formas de pago integradas | — | NOT_STARTED |
| `spRep_LiquidacionOutPut` | Caja | Liquidación output | — | NOT_STARTED |
| `spRep_LiquidacionSuma` | Caja | Liquidación suma | — | NOT_STARTED |
| `spRep_MensajeUsuario` | Usuarios | Mensajes a usuario | — | NOT_STARTED |
| `spRep_Ocupabilidad` | Mesas | Ocupabilidad de mesas | — | NOT_STARTED |
| `spRep_PaloteoComparativo` | Ventas | Paloteo comparativo | — | NOT_STARTED |
| `spRep_PaloteoInsumo` | Almacén | Paloteo de insumos | — | NOT_STARTED |
| `spRep_PaloteoInsumoIntegrado` | Almacén | Paloteo insumos integrado | — | NOT_STARTED |
| `spRep_PaloteoOferta` | Ventas | Paloteo de ofertas | — | NOT_STARTED |
| `spRep_PaloteoProduccion` | Cocina | Paloteo de producción | — | NOT_STARTED |
| `spRep_PaloteoProduccionPorMes` | Cocina | Paloteo producción por mes | — | NOT_STARTED |
| `spRep_PaloteoPropiedad` | Ventas | Paloteo de propiedades | — | NOT_STARTED |
| `spRep_PaloteoSubProd` | Almacén | Paloteo sub-productos | — | NOT_STARTED |
| `spRep_PaloteoVentaIntegrado` | Ventas | Paloteo ventas integrado | — | NOT_STARTED |
| `spRep_Pedido` | Pedidos | Reporte de pedidos | — | NOT_STARTED |
| `spRep_PlanillaMovilidad` | RRHH | Planilla de movilidad | — | NOT_STARTED |
| `spRep_PlanillaMovilidadGeneral` | RRHH | Planilla movilidad general | — | NOT_STARTED |
| `spRep_PrincipalCliente` | Clientes | Reporte principal clientes | — | NOT_STARTED |
| `spRep_Propina` | Caja | Reporte de propinas | — | NOT_STARTED |
| `spRep_Ranking` | Ventas | Ranking de productos | — | NOT_STARTED |
| `spRep_RankingIntegrado` | Ventas | Ranking integrado | — | NOT_STARTED |
| `spRep_RegVenta` | Ventas | Registro de ventas | — | NOT_STARTED |
| `spRep_RegVentaIntegrado` | Ventas | Registro ventas integrado | — | NOT_STARTED |
| `spRep_RegVentaSunat` | FE Perú | Registro ventas SUNAT | — | NOT_STARTED |
| `spRep_RegVentaSunatAD` | FE Perú | Reg. ventas SUNAT adicional | — | NOT_STARTED |
| `spRep_RepClieFrecuentes` | Clientes | Clientes frecuentes | — | NOT_STARTED |
| `spRep_ResultadoOperativo` | Gerencial | Resultado operativo | — | NOT_STARTED |
| `spRep_Rotacion` | Mesas | Rotación de mesas | — | NOT_STARTED |
| `spRep_TiempoDelivery` | Delivery | Tiempos de delivery | — | NOT_STARTED |
| `spRep_TiempoDeliveryIntegrado` | Delivery | Tiempos delivery integrado | — | NOT_STARTED |
| `spRep_TiempoSalon` | Mesas | Tiempos en salón | — | NOT_STARTED |
| `spRep_TipoProductoVentaIntegrado` | Ventas | Ventas por tipo producto | — | NOT_STARTED |
| `spRep_VentaFecha` | Ventas | Ventas por fecha | — | NOT_STARTED |
| `spRep_VentaIntervaloIntegrado` | Ventas | Ventas por intervalo | — | NOT_STARTED |
| `spRep_VentaMensualCanalesIntegrado` | Ventas | Ventas mensual por canales | — | NOT_STARTED |
| `spRep_VentaMensualIntegrado` | Ventas | Ventas mensual integrado | — | NOT_STARTED |

### KDS — `USP_KDS_*`

| SP | Módulo | Descripción | Equivalente .NET | Estado |
|---|---|---|---|---|
| `USP_KDS_ObtenerProductoPedido` | KDS | Obtiene productos del pedido para KDS | — | NOT_STARTED |
| `USP_KDS_ObtenerProductoPedidoDeCombo` | KDS | Productos de combo para KDS | — | NOT_STARTED |
| `USP_KDS_ResporteTiempoPedido` | KDS | Reporte tiempo de pedido KDS | — | NOT_STARTED |
| `USP_KDS_ResporteTiempoProducto` | KDS | Reporte tiempo de producto KDS | — | NOT_STARTED |

### Insumos — `USP_*`

| SP | Módulo | Descripción | Equivalente .NET | Estado |
|---|---|---|---|---|
| `USP_MODIFICARINSUMOS` | Almacén | Modifica insumos | — | NOT_STARTED |
| `USP_LISTARINSUMOS` | Almacén | Lista insumos | — | NOT_STARTED |
| `USP_AGREGARINSUMOS` | Almacén | Agrega insumos | — | NOT_STARTED |
| `USP_ELIMINARINSUMOS` | Almacén | Elimina insumos | — | NOT_STARTED |
| `USP_actualizaStockInsumo` | Almacén | Actualiza stock de insumo | — | NOT_STARTED |

### Auditoría — `usp_Aud_*`

| SP | Módulo | Descripción | Equivalente .NET | Estado |
|---|---|---|---|---|
| `usp_Aud_ObtenerSalones` | Auditoría | Obtiene salones para auditoría | — | NOT_STARTED |

### Operativos Misceláneos — `sp_*`

| SP | Módulo | Descripción | Equivalente .NET | Estado |
|---|---|---|---|---|
| `sp_CopiaArchivosRemotos` | Sistema | Copia archivos remotos | — | NOT_STARTED |
| `sp_TraeDatosVentas_SyBase` | Integración | Trae datos ventas (SyBase) | — | NOT_STARTED |
| `sp_TraeDatosPagos_SyBase` | Integración | Trae datos pagos (SyBase) | — | NOT_STARTED |
| `sp_InsUptSocioDelivery_SyBASE` | Delivery | Ins/Upd socio delivery (SyBase) | — | NOT_STARTED |
| `sp_CreaTemporalSocio_SyBASE` | Delivery | Crea temporal socio (SyBase) | — | NOT_STARTED |
| `sp_UpdFotoDelivery` | Delivery | Actualiza foto delivery | `ActualizarFotoClienteDeliveryHandler` + `ClienteDeliveryRepository.ActualizarFotoAsync` | IN_PROGRESS |
| `sp_UpdFotoProducto` | Productos | Actualiza foto producto | — | NOT_STARTED |
| `sp_UpdImagenCaja` | Caja | Actualiza imagen caja | — | NOT_STARTED |
| `usp_Inforest_ObtenerClientesFrecuente` | Clientes | Clientes frecuentes | — | NOT_STARTED |

---

## Observaciones

- Los SPs `sp_*SyBase*` sugieren una integración histórica con Sybase que puede estar obsoleta. Estado: UNKNOWN.
- Los SPs de reporte (`spRep_*`) son reutilizables en el Target independientemente del motor de reportes elegido.
- Los SPs KDS pueden reutilizarse si se mantiene SQL Server como base Target.

---

## Referencias

- `legacy-restaurant/database-sql-server/5. SP.sql` — fuente
- [Legacy Database](legacy-database.md)
- [Trazabilidad](../migration/traceability-matrix.md)
