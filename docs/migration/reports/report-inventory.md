# Inventario de Reportes — INFOREST

> Última actualización: 2026-08-12
>
> Total reportes Crystal Legacy: 206 `.Dsr`
> Total clasificados: 206
> Estado general: IN_PROGRESS — Etapa 10

---

## Criterio de Clasificación por Grupo

| Grupo | Descripción | Prioridad |
|---|---|---|
| A | Core operativo — emitidos desde módulo local (single-tenant) | Alta |
| B | Integrados multi-local — requieren acceso centralizado | Media |
| C | Delivery / Motorizado | Media (paralelo Etapa 9) |
| D | Almacén / Compras / Inventario | Baja (posponer hasta módulo Almacén) |
| E | Otros / Administrativos / Fiscales | Baja |

---

## Grupo A — Core Operativo

| Crystal Legacy (`.Dsr`) | SP(s) Legacy | Formulario VB6 | Tablas principales | Prioridad | Estado |
|---|---|---|---|---|---|
| `dsrLiquidacion` / `dsrLiquidacionVenta` | `spRep_Liquidacion`, `spRep_LiquidacionSuma`, `spRep_LiquidacionOutPut`, `spRep_LiquidacionOrigenVenta`, `spRep_Liquidacion_NC`, `spRep_LiquidacionSuma_NC`, `spRep_LiquidacionOutPut_NC` | `frmRepLiquidacion.frm` | MPEDIDO, DPEDIDO, MDOCUMENTO, MTURNO | Alta | NOT_STARTED |
| `dsrVentaFecha` | `spRep_VentaFecha` | `frmRepVentaFecha.frm`, `frmRepProductoMes.frm` | MPEDIDO, DPEDIDO, MDOCUMENTO | Alta | NOT_STARTED |
| `dsrFormaPagoDiarioIntegradoLocales` / `dsrFormaPagoDiarioIntegradoSector` | `spRep_FormaPagoIntegrado` | `frmRepFormaPagoIntegrado.frm` | DPAGODOCUMENTO, MDOCUMENTO, TTARJETACREDITO | Alta | NOT_STARTED |
| `dsrPaloteoA` / `dsrPaloteoC` / `dsrPaloteoD` / `dsrPaloteoP` / `dsrPaloteoR` / `dsrPaloteoItemCombo` | `spRep_PaloteoProduccion` | `frmRepPaloteo.frm` | MPEDIDO, DPEDIDO, TPRODUCTO, vGrupo, vSubGrupo | Alta | NOT_STARTED |
| `dsrPaloteoSPD` | `spRep_PaloteoSubProd` | `frmRepPaloteoSubProd.frm` | MPEDIDO, DPEDIDO, vPRODUCTOXPRODUCTO | Alta | IN_PROGRESS |
| `dsrRegistroVentaC` / `dsrRegistroVentaD` / `dsrRegistroVentaE` / `dsrRegistroVentaF` / `dsrRegistroVentaR` | `spRep_RegVenta` | `frmRepRegistroVenta.frm` | MDOCUMENTO, DPEDIDO, MPEDIDO | Alta | NOT_STARTED |
| `dsrRankingD` / `dsrRankingR` | `spRep_Ranking` | `frmRepRanking.frm` | MPEDIDO, DPEDIDO, TPRODUCTO | Alta | NOT_STARTED |
| `dsrPropinaD` / `dsrPropinaR` | `spRep_Propina` | `frmRepPropina.frm` | DPAGODOCUMENTO, TTARJETACREDITO, MPEDIDO | Alta | IN_PROGRESS |
| `dsrComandaD` / `dsrComandaR` | `spRep_Comanda` | `frmRepComanda.frm` | MPEDIDO, DPEDIDO, TPRODUCTO | Alta | IN_PROGRESS |
| `dsrDescuento` / `dsrDescuentoD` | `spRep_Descuento` | `frmRepDescuento.frm` | MDOCUMENTO, DPEDIDO | Media | NOT_STARTED |
| `dsrCortesia` / `dsrCortesiaD` / `dsrCortesiaR` | `spRep_Cortesia` | `frmRepCortesia.frm` | MPEDIDO, DPEDIDO | Media | NOT_STARTED |
| `dsrAnuladoD` / `dsrAnuladoR` | `spRep_Anulacion` | `frmRepAnulado.frm` | MPEDIDO, DPEDIDO | Media | NOT_STARTED |
| `dsrCancelacionD` / `dsrCancelacionR` | `spRep_Cancelacion` | `frmRepCancelacion.frm` | MPEDIDO, DPEDIDO | Media | NOT_STARTED |
| `dsrPaloteoComparativo` | `spRep_PaloteoComparativo` | `frmRepPaloteoComparativo.frm` | MPEDIDO, DPEDIDO, TPRODUCTO | Media | IN_PROGRESS |
| `dsrPaloteoPropiedadA` / `D` / `G` / `P` | `spRep_PaloteoPropiedad` | `frmRepPaloteoPropiedades.frm` | MPEDIDO, DPEDIDO, TPROPIEDAD | Media | NOT_STARTED |
| `dsrPaloteoOfertaD` / `P` / `R` | `spRep_Paloteooferta` | `frmRepPaloteoOfertas.frm` | MPEDIDO, DPEDIDO, TOFERTA | Baja | NOT_STARTED |
| `dsrCobranzaFecha` | `spRep_CobranzaFecha` | `frmRepCobranzaFecha.frm` | DPAGODOCUMENTO, MDOCUMENTO | Media | NOT_STARTED |
| `dsrOcupabilidadD` / `dsrOcupabilidadR` | `spRep_Ocupabilidad` | `frmRepOcupabilidad.frm` | MPEDIDO, MESA | Media | NOT_STARTED |
| `dsrRotacionMesa` | `spRep_Rotacion` | `frmRepRotacionMesa.frm` | MPEDIDO, MESA | Baja | NOT_STARTED |
| `dsrDiferencia` | `spRep_Diferencia` | `frmRepDiferencia.frm` | MPEDIDO, MDOCUMENTO | Media | NOT_STARTED |
| `dsrResultadosOperativos` | `spRep_ResultadoOperativo` | `frmRepResultadoOperativo.frm` | MDOCUMENTO, MTURNO | Media | NOT_STARTED |
| `dsrReciboEgreso` | `spRep_ReciboEgreso` | `frmReciboEgreso.frm` | MEGRESO | Baja | NOT_STARTED |
| `dsrPedidosD` / `dsrPedidosFF` / `dsrPedidosFFyKDS` / `dsrPedidosR` / `dsrPedidosX` | `spRep_Pedido` | `frmRepPedido.frm` | MPEDIDO, DPEDIDO | Baja | NOT_STARTED |
| `dsrMotivoDescuento` | `spRep_Descuento` | `frmRepDescuento.frm` | MPEDIDO, TDESCUENTO | Baja | NOT_STARTED |
| `dsrInsumosCD` / `dsrInsumosD` / `dsrInsumosR` | `spRep_PaloteoInsumo` | `frmRepInsumoVentas.frm` | MPEDIDO, DPEDIDO, TINSUMO | Media | NOT_STARTED |
| `dsrRepInsumos` / `dsrRepInsumosPtoVenta` | `spRep_PaloteoInsumo` | `frmRepInsumoVentas.frm` | DPEDIDO, TINSUMO | Media | NOT_STARTED |
| `dsrRegistroVentaSunat` / `dsrRegistroVentaSunatAD` / `dsrRegistroVentaSunatPago` | `spRep_RegVentaSunat` | `frmRepRegistroVenta.frm` | MDOCUMENTO, SUNAT tables | Media | NOT_STARTED |
| `dsrTiempoSalon` | `spRep_TiempoSalon` | `frmRepTiempoSalon.frm` | MPEDIDO | Baja | NOT_STARTED |
| `dsrAsistencia` / `dsrAsistencia2` | `spRep_Asistencia` | `frmRepAsistencia.frm` | MASISTENCIA | Baja | NOT_STARTED |
| `dsrProducto` | — | — | TPRODUCTO | Baja | NOT_STARTED |
| `dsrPlato` / `dsrPlatoCombo` | — | — | TPRODUCTO, TPRODUCTOXPRODUCTO | Baja | NOT_STARTED |
| `dsrPrecios` | — | — | TPRODUCTO, TPRECIO | Baja | NOT_STARTED |
| `dsrSubGrupo` | — | — | TSUBGRUPO | Baja | NOT_STARTED |
| `dsrArea` | — | — | TAREA | Baja | NOT_STARTED |
| `dsrImpresora` | — | — | TIMPRESORA | Baja | NOT_STARTED |
| `dsrMensajeUsuario` | `spRep_MensajeUsuario` | `xfrmMozoUsuario.frm` | TMENSAJEUSUARIO | Baja | NOT_STARTED |
| `dsrProductoNoEnlazados` | `spRep_ProductosNoEnlazados` | `frmRepProductosNoEnlazados.frm` | TPRODUCTO | Baja | NOT_STARTED |
| `dsrProductoXMes` / `dsrProductoXMesMonto` | `spRep_PaloteoProduccionPorMes`, `spRep_VentaFecha` | `frmRepProductoMes.frm` | MPEDIDO, DPEDIDO | Baja | NOT_STARTED |
| `dsrSemana` | — | — | MTURNO | Baja | NOT_STARTED |
| `dsrVentaTurno` | — | — | MPEDIDO, MTURNO | Baja | NOT_STARTED |
| `dsrVentaAcumulada` | — | — | MPEDIDO, MDOCUMENTO | Baja | NOT_STARTED |
| `dsrVentaAnualGrafico` | — | — | MPEDIDO, MDOCUMENTO | Baja | NOT_STARTED |
| `dsrVentaCompAnual` | `spRep_VentaCompMensual` | `frmRepVentaCompMensual.frm` | MPEDIDO, MDOCUMENTO | Baja | NOT_STARTED |
| `dsrVentaCompMensualDia` | `spRep_VentaCompMensual` | `frmRepVentaCompMensual.frm` | MPEDIDO, MDOCUMENTO | Baja | NOT_STARTED |
| `dsrReporte` / `dsrReporte4` / `dsrReporte5` / `dsrReporte6` / `dsrReporte7` | — | — | Varios | Baja | NOT_STARTED |
| `dsrRepDescargo` / `dsrRepDescargoResumido` | — | — | MDESCARGO | Baja | NOT_STARTED |
| `dsrRepNotaCredito` / `dsrNotaCredito` / `dsrNotaCreditoObservacion` | — | — | MDOCUMENTO (nota crédito) | Baja | NOT_STARTED |
| `dsrCliente` / `dsrClienteFactura` | — | — | TCLIENTE | Baja | NOT_STARTED |
| `dsrContacto` | — | — | TCLIENTE | Baja | NOT_STARTED |
| `dsrBoleta` / `dsrBoletaC` | — | — | MDOCUMENTO | Baja | NOT_STARTED |
| `dsrAnaliticoMozo` / `dsrAnaliticoMozoC` | `spRep_AnaliticoMozo` | `frmRepAnaliticoMozo.frm` | MPEDIDO, DPEDIDO | Media | NOT_STARTED |
| `dsrMozoProduccion` / `dsrProduccionMozo` | — | — | MPEDIDO, DPEDIDO | Baja | NOT_STARTED |
| `frmRepCtaCte` (modos consolidado/resumido/detallado) | `spRep_CtaCteN` | `frmRepCtaCte.frm` | MPEDIDO, DPEDIDO, vCompania, vTipoCtaCte, vSubTipoCtaCte | Media | MIGRATED |
| `dsrRepCtaCteIntegradoConsolidado` / `dsrRepCtaCteIntegradoDetallado` (variantes XC) | `spRep_CtaCteIntegrado` | `frmRepCtaCteIntegrado.frm` | MDOCUMENTO, TCLIENTE, VCOMPANIA | Media | IN_PROGRESS |
| `dsrClienteCtaCte` / `dsrClienteDelivery` / `dsrClienteDeudaD` / `dsrClienteDeudaR` / `dsrClienteDeudaHistorico` | `spRep_CuentasCobrar` | `frmRepClienteDeuda.frm` | MDOCUMENTO, TCLIENTE | Media | NOT_STARTED |
| `dsrCuentaCobrar` | `spRep_CuentasCobrar` | `frmRepClienteDeuda.frm` | MDOCUMENTO, TCLIENTE | Media | NOT_STARTED |
| `dsrPrincipalD` / `dsrPrincipalR` | `spRep_PrincipalCliente` | `frmRepPrincipal.frm` | MDOCUMENTO, TCLIENTE | Alta | IN_PROGRESS |
| `dsrReciboEgreso` / `dsrReciboEgresoCompra` | `spRep_ReciboEgreso` | `frmReciboEgreso.frm` | MEGRESO | Baja | NOT_STARTED |
| `dsrRepCtaCteIntegradoConsolidadoXM` / `dsrRepCtaCteIntegradoDetalladoXM` (variantes) | `spRep_CtaCteIntegrado` | `frmRepCtaCteIntegrado.frm` | MDOCUMENTO, TCLIENTE | Media | NOT_STARTED |
| `dsrReserva` / `dsrReservas` | — | — | TRESERVA | Baja | NOT_STARTED |
| `dsrGuia` / `dsrGuiaCorrelativo` / `dsrNewFormatoGuia` | — | — | MGUIA | Baja | NOT_STARTED |
| `dsrLiquidacion3` | `spRep_Liquidacion` | `frmRepLiquidacion.frm` | MTURNO, MDOCUMENTO | Baja | NOT_STARTED |
| `dsrRepresentación` (reprecargoconsumo / det) | — | — | MPEDIDO, DPEDIDO | Baja | NOT_STARTED |
| `dsrRepresentación` (dsrReprecargoconsumo / dsrReprecargoconsumodet) | — | — | MPEDIDO | Baja | NOT_STARTED |
| `DataReport1` | — | — | — | Baja | NOT_STARTED |
| `xxx` | — | — | — | Baja | NOT_STARTED |
| `dsrMovimiento` | — | — | MMOVIMIENTO | Baja | NOT_STARTED |

---

## Grupo B — Integrados Multi-Local

| Crystal Legacy (`.Dsr`) | SP(s) Legacy | Formulario VB6 | Tablas principales | Prioridad | Estado |
|---|---|---|---|---|---|
| `dsrPaloteoVentaIntegradol` / `dsrPaloteoVentaIntegradoS` / `dsrPaloteoVentaIntegradoComparativo` | `spRep_PaloteoVentaIntegrado` | `frmRepPaloteoVentaIntegrado.frm` | MPEDIDO, DPEDIDO (multi-local) | Alta | IN_PROGRESS |
| `dsrRankingIntegradoComparativo` / `dsrRankingIntegradol` | `spRep_RankingIntegrado` | `frmRepRankingIntegrado.frm` | MPEDIDO, DPEDIDO (multi-local) | Alta | IN_PROGRESS |
| `dsrVentaMensualIntegrada` / `dsrVentaMensualIntegradaS` | `spRep_VentaMensualIntegrado` | `frmRepVentaMensualIntegrado.frm` | MPEDIDO, MDOCUMENTO | Alta | IN_PROGRESS |
| `dsrVentaDiariaIntegrada` / `dsrVentaDiariaIntegradaS` / `dsrVentaDiariaIntegradaMetas` | `spRep_VentaMensualIntegrado` | `frmRepDiarioVentaIntegrado.frm` | MPEDIDO, MDOCUMENTO | Alta | IN_PROGRESS |
| `dsrFormaPagoIntegrado` | `spRep_FormaPagoIntegrado` | `frmRepFormaPagoIntegrado.frm` | DPAGODOCUMENTO, MDOCUMENTO | Alta | IN_PROGRESS |
| `dsrRegistroVentaIntegradoE` / `dsrRegistroVentaIntegradoL` / `dsrRegistroVentaIntegradoS` | `spRep_RegVentaIntegrado` | `frmRepRegistroVentaIntegrado.frm` | MDOCUMENTO | Media | NOT_STARTED |
| `dsrInsumosIntegrados` | `spRep_PaloteoInsumoIntegrado` | `frmRepInsumoVentasIntegrado.frm` | DPEDIDO, TINSUMO | Media | NOT_STARTED |
| `dsrTipoProductoVentaIntegrado` | `spRep_TipoProductoVentaIntegrado` | `frmRepVentaTipoProductoIntegrado.frm` | MPEDIDO, TPRODUCTO | Media | NOT_STARTED |
| `dsrVentaMensualCanalesIntegrada` | `spRep_VentaMensualCanalesIntegrado` | `frmRepVentaMensualCanalesIntegrado.frm` | MPEDIDO | Media | NOT_STARTED |
| `dsrCompConsumoInsumos` / `dsrCompConsumoInsumosValor` / `dsrComparativoConsumoInsumos` / `dsrComparativoConsumoInsumosValor` | `sp_ComparativoConsumo` | — | MPEDIDO, DPEDIDO | Baja | NOT_STARTED |

---

## Grupo C — Delivery / Motorizado

| Crystal Legacy (`.Dsr`) | SP(s) Legacy | Formulario VB6 | Tablas principales | Prioridad | Estado |
|---|---|---|---|---|---|
| `dsrAnaliticoMotorizado` / `dsrAnaliticoMotorizadoC` | `spRep_AnaliticoMotorizado` | `frmRepAnaliticoMotorizado.frm` | MPEDIDO, TMOTORIZADODATOS | Alta | NOT_STARTED |
| `dsrAnaliticoMotorizadoIntegradoConsolidado` / variantes | `spRep_AnaliticoMotorizadoIntegrado` | `frmRepAnaliticoMotorizadoIntegrado.frm` | MPEDIDO (multi-local) | Alta | IN_PROGRESS |
| `dsrDeliveryDetallado` / `dsrDeliveryDetalladoxDia` / `dsrDeliveryGrupo` | `spRep_TiempoDelivery` | `frmRepTiempoDelivery.frm` | MPEDIDO, TMOTORIZADODATOS | Alta | NOT_STARTED |
| `dsrDeliveryTiempo` / `dsrDeliveryTiempoCD` | `spRep_TiempoDelivery` | `frmRepTiempoDelivery.frm` | MPEDIDO | Alta | NOT_STARTED |
| `dsrTiempoDeliveryIntegrado` | `spRep_TiempoDeliveryIntegrado` | `frmRepTiempoDeliveryIntegrado.frm` | MPEDIDO | Media | NOT_STARTED |
| `dsrPlanillaMovilidadGeneral` / `dsrPlanillaMovilidadPorDiaDetallado` / `dsrPlanillaMovilidadPorDiaResumido` | `spRep_PlanillaMovilidadGeneral`, `spRep_PlanillaMovilidad` | `frmRepPlanillaMovilidadMotorizado.frm` | TMOTORIZADODATOS, MPEDIDO | Alta | NOT_STARTED |
| `dsrControlMotorizado` | `spRep_ControlMotorizado` | `frmRepControlMotorizado.frm` | MPEDIDO, TMOTORIZADODATOS | Alta | NOT_STARTED |
| `dsrControlEnviosAutorizados` | `spRep_ControlEnviosAutorizados` | `frmRepControlEnviosAutorizados.frm` | MPEDIDO | Media | NOT_STARTED |
| `dsrReporteEntrega` / `dsrReporteEntregaFormato2` / `dsrReporteEntregaRegistro` / `dsrReporteEntregaResumidoProd` | — | — | MPEDIDO, DPEDIDO | Baja | NOT_STARTED |
| `dsrReporteEmpacador` | — | — | MPEDIDO | Baja | NOT_STARTED |
| `dsrTiempoKdsP` / `dsrTiempoKdsR` | `USP_KDS_ResporteTiempoPedido`, `USP_KDS_ResporteTiempoProducto` | — (KDS monitor) | DPEDIDOKDS, TPRODUCTO | Alta | IN_PROGRESS |
| `dsrTiempoChefDeta` / `dsrTiempoChefResumido` | — | — | DPEDIDOKDS | Baja | NOT_STARTED |

---

## Grupo D — Almacén / Compras / Inventario

| Crystal Legacy | SP(s) Legacy | Formulario VB6 | Tablas principales | Prioridad | Estado |
|---|---|---|---|---|---|
| `CrtComprasxTipoDocumentoIntegrado` | `spRep_RegCompraTipoDocumentoIntegrado` | `frmRepRegistroComprasIntegrado.frm` | MCOMPRA | Baja | NOT_STARTED |
| `CrtComprasxTipoProductoIntegrado` | `spRep_RegCompraTipoProductoIntegrado` | `frmRepRegistroComprasIntegrado.frm` | MCOMPRA | Baja | NOT_STARTED |
| `CrtInsSolicitadosVsEntregadoDetallado` | `spRep_RequerimientosPorAreaIntegrado` | `frmRepSolicitadoVsEntregado.frm` | MINSUMO | Baja | NOT_STARTED |
| `CrtInsSolicitadosVsEntregadoResumido` | `spRep_RequerimientosPorAreaIntegrado` | `frmRepSolicitadoVsEntregado.frm` | MINSUMO | Baja | NOT_STARTED |
| `CrtRegComprasIntegrado` | `spRep_RegCompraIntegrado` | `frmRepRegistroComprasIntegrado.frm` | MCOMPRA | Baja | NOT_STARTED |
| `CrtRepInventarioIntegrado` | `spRep_InventarioDiaIntegrado` | `FrmRepInventariosIntegrado.frm` | TINVENTARIO | Baja | NOT_STARTED |
| `CrtRepMovimientosIntegrado` | `spRep_MovimientosArticuloIntegrado` | `FrmRepMovimientosIntegrado.frm` | MMOVIMIENTO | Baja | NOT_STARTED |
| `CrtRepMovimientosDetalladoIntegrado` | `spRep_MovimientosArticuloDetalladoIntegrado` | `FrmRepMovimientosIntegrado.frm` | MMOVIMIENTO | Baja | NOT_STARTED |
| `dsrRepInsumos` (almacén) | — | — | TINSUMO | Baja | NOT_STARTED |
| `dsrInventarioConsolidado` | — | — | TINVENTARIO | Baja | NOT_STARTED |

---

## Grupo E — Otros / Administrativos / Fiscales

| Crystal Legacy | SP(s) Legacy | Formulario VB6 | Tablas principales | Prioridad | Estado |
|---|---|---|---|---|---|
| `dsrRegistroVentaSunat` (variantes) | `spRep_RegVentaSunat`, `spRep_RegVentaSunatAD` | `frmRepRegistroVenta.frm` | MDOCUMENTO (Sunat) | Media | NOT_STARTED |
| `dsrRepCtaCteIntegrado*` (variantes XC, XM) | `spRep_CtaCteIntegrado` | `frmRepCtaCteIntegrado.frm` | MDOCUMENTO, TCLIENTE | Media | NOT_STARTED |
| `dsrCaja` | — | — | MTURNO, MCAJA | Media | NOT_STARTED |
| `dsrRepresentación recargo` (2 variantes) | — | — | MPEDIDO | Baja | NOT_STARTED |

---

## Resumen de Cobertura

| Grupo | Total reportes | Migrados | Cobertura |
|---|---|---|---|
| A — Core Operativo | ~90 | 0 | 0% |
| B — Integrados Multi-Local | ~30 | 0 | 0% |
| C — Delivery / Motorizado | ~20 | 0 | 0% |
| D — Almacén / Compras | ~10 | 0 | 0% |
| E — Otros / Fiscales | ~10 | 0 | 0% |
| **Total** | **~160** | **0** | **0%** |

> Nota: Los 206 Crystal Legacy incluyen múltiples variantes (consolidado/detallado/resumido) del mismo reporte funcional. El inventario unifica variantes bajo el mismo nombre funcional.

---

## Convención de Nombres FastReport

| Patrón Legacy | Patrón FastReport .NET |
|---|---|
| `dsr{Nombre}.Dsr` | `Rep{Nombre}.frx` |
| `dsr{Nombre}D.Dsr` (detallado) | `Rep{Nombre}Detallado.frx` |
| `dsr{Nombre}R.Dsr` (resumido) | `Rep{Nombre}Resumido.frx` |

Ubicación plantillas: `modern-net8/src/Inforest.Desktop/reports/templates/`
