# Plantillas FastReport — INFOREST

> ADR-007: FastReport .NET reemplaza Crystal Reports
> 
> Convención de nombres: `Rep{NombreReporte}.frx`

## Grupo A — Core Operativo (pendiente crear .frx)

| Plantilla | Reporte | SP Legacy | Crystal Legacy |
|---|---|---|---|
| `RepComanda.frx` | Comanda (resumido) | `spRep_Comanda` | `dsrComandaR.Dsr` |
| `RepComandaDetallado.frx` | Comanda (detallado) | `spRep_Comanda` | `dsrComandaD.Dsr` |
| `RepPropina.frx` | Propinas | `spRep_Propina` | `dsrPropinaD.Dsr` |
| `RepPrincipalClienteDetalle.frx` | Clientes Principales (detalle) | `spRep_PrincipalCliente` | `dsrPrincipalD.Dsr` |
| `RepPrincipalClienteResumen.frx` | Clientes Principales (resumen) | `spRep_PrincipalCliente` | `dsrPrincipalR.Dsr` |
| `RepPaloteoComparativo.frx` | Paloteo Comparativo | `spRep_PaloteoComparativo` | `dsrPaloteoComparativo.Dsr` |
| `RepPaloteoSubProd.frx` | Paloteo Sub-Productos | `spRep_PaloteoSubProd` | `dsrPaloteoSPD.Dsr` |

## Grupo B — Integrados Multi-Local (pendiente crear .frx)

| Plantilla | Reporte | SP Legacy | Crystal Legacy |
|---|---|---|---|
| `RepCtaCteIntegradoConsolidado.frx` | Cta.Cte. Integrado (consolidado) | `spRep_CtaCteIntegrado` | `dsrRepCtaCteIntegradoConsolidado.Dsr` |
| `RepCtaCteIntegradoDetallado.frx` | Cta.Cte. Integrado (detallado) | `spRep_CtaCteIntegrado` | `dsrRepCtaCteIntegradoDetallado.Dsr` |
| `RepPaloteoVentaIntegradoDetallado.frx` | Paloteo Venta Integrado (det.) | `spRep_PaloteoVentaIntegrado` | `dsrPaloteoVentaIntegradol.Dsr` |
| `RepPaloteoVentaIntegradoResumido.frx` | Paloteo Venta Integrado (res.) | `spRep_PaloteoVentaIntegrado` | `dsrPaloteoVentaIntegradoS.Dsr` |
| `RepRankingIntegradoDetallado.frx` | Ranking Integrado (detallado) | `spRep_RankingIntegrado` | `dsrRankingIntegradoComparativo.Dsr` |
| `RepRankingIntegradoResumido.frx` | Ranking Integrado (resumido) | `spRep_RankingIntegrado` | `dsrRankingIntegradol.Dsr` |
| `RepVentaMensualIntegrado.frx` | Venta Mensual Integrada | `spRep_VentaMensualIntegrado` | `dsrVentaMensualIntegrada.Dsr` |
| `RepVentaDiariaIntegrado.frx` | Venta Diaria Integrada | `spRep_VentaMensualIntegrado` | `dsrVentaDiariaIntegrada.Dsr` |

## Grupo C — Delivery / Motorizado (pendiente crear .frx)

| Plantilla | Reporte | SP Legacy | Crystal Legacy |
|---|---|---|---|
| `RepAnaliticoMotorizadoIntegrado.frx` | Analítico Motorizado Integrado | `spRep_AnaliticoMotorizadoIntegrado` | `dsrAnaliticoMotorizadoIntegradoConsolidado.Dsr` |
| `RepTiempoKdsPedido.frx` | Tiempo KDS por Pedido | `USP_KDS_ResporteTiempoPedido` | `dsrTiempoKdsP.Dsr` |
| `RepTiempoKdsProducto.frx` | Tiempo KDS por Producto | `USP_KDS_ResporteTiempoProducto` | `dsrTiempoKdsR.Dsr` |

## Notas

- Los archivos `.frx` son plantillas XML de FastReport y se crean con el diseñador FastReport.
- Cada plantilla debe tener un datasource con el nombre indicado en el handler correspondiente.
- Los archivos `.frx` NO se incluyen en este commit inicial; se crearán cuando FastReport esté instalado.
- Para crear una plantilla: FastReport Designer → New Report → Add DataSource → nombre exacto del handler.
