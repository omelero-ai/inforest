# Módulo Consultas — INFOREST

> Ejecutable Legacy: `Consulta.exe`
> Etapa de migración: 10 — Reportes, Consultas y Salidas Operativas
> Estado: ANALYSIS

---

## Descripción

El módulo `Consulta.exe` concentra todos los formularios de consulta y reporte del sistema INFOREST.
Es el punto de acceso para supervisores, gerentes y administradores que necesitan obtener
información operativa, estadísticas de venta, análisis de productos, control de motorizados,
reportes de propinas, etc.

---

## Formularios VB6 (124 formularios)

Los formularios se clasifican en tres tipos funcionales:

### Tipo 1 — Reporte Crystal Reports (→ FastReport .NET)

Formularios que cargan un `.Dsr` y lo muestran/imprimen.

| Formulario VB6 | SP Legacy | Crystal | Prioridad |
|---|---|---|---|
| `frmRepComanda.frm` | `spRep_Comanda` | `dsrComandaD/R` | Alta |
| `frmRepPropina.frm` | `spRep_Propina` | `dsrPropinaD/R` | Alta |
| `frmRepPrincipal.frm` | `spRep_PrincipalCliente` | `dsrPrincipalD/R` | Alta |
| `frmRepLiquidacion.frm` | `spRep_Liquidacion` (múltiples) | `dsrLiquidacion` | Alta |
| `frmRepVentaFecha.frm` | `spRep_VentaFecha` | `dsrVentaFecha` | Alta |
| `frmRepRegistroVenta.frm` | `spRep_RegVenta` (múltiples) | `dsrRegistroVentaC/D/E/R` | Alta |
| `frmRepRanking.frm` | `spRep_Ranking` | `dsrRankingD/R` | Alta |
| `frmRepPaloteo.frm` | `spRep_PaloteoProduccion` | `dsrPaloteoA/C/D/P/R` | Alta |
| `frmRepPaloteoSubProd.frm` | `spRep_PaloteoSubProd` | `dsrPaloteoSPD` | Alta |
| `frmRepPaloteoComparativo.frm` | `spRep_PaloteoComparativo` | `dsrPaloteoComparativo` | Media |
| `frmRepDescuento.frm` | `spRep_Descuento` | `dsrDescuento/D` | Media |
| `frmRepCortesia.frm` | `spRep_Cortesia` | `dsrCortesia/D/R` | Media |
| `frmRepAnulado.frm` | `spRep_Anulacion` | `dsrAnuladoD/R` | Media |
| `frmRepOcupabilidad.frm` | `spRep_Ocupabilidad` | `dsrOcupabilidadD/R` | Media |
| `frmRepCtaCteIntegrado.frm` | `spRep_CtaCteIntegrado` | `dsrRepCtaCteIntegrado*` | Media |
| `frmRepPaloteoVentaIntegrado.frm` | `spRep_PaloteoVentaIntegrado` | `dsrPaloteoVentaIntegrado*` | Alta |
| `frmRepRankingIntegrado.frm` | `spRep_RankingIntegrado` | `dsrRankingIntegrado*` | Alta |
| `frmRepVentaMensualIntegrado.frm` | `spRep_VentaMensualIntegrado` | `dsrVentaMensualIntegrada*` | Alta |
| `frmRepDiarioVentaIntegrado.frm` | `spRep_VentaMensualIntegrado` | `dsrVentaDiariaIntegrada*` | Alta |
| `frmRepFormaPagoIntegrado.frm` | `spRep_FormaPagoIntegrado` | `dsrFormaPagoIntegrado*` | Alta |
| `frmRepAnaliticoMotorizado.frm` | `spRep_AnaliticoMotorizado` | `dsrAnaliticoMotorizado/C` | Alta |
| `frmRepAnaliticoMotorizadoIntegrado.frm` | `spRep_AnaliticoMotorizadoIntegrado` | `dsrAnaliticoMotorizadoIntegrado*` | Alta |
| `frmRepTiempoDelivery.frm` | `spRep_TiempoDelivery` | `dsrDelivery*` | Alta |
| `frmRepPlanillaMovilidadMotorizado.frm` | `spRep_PlanillaMovilidadGeneral` | `dsrPlanillaMovilidad*` | Alta |
| `frmRepControlMotorizado.frm` | `spRep_ControlMotorizado` | `dsrControlMotorizado` | Alta |
| `frmRepCobranzaFecha.frm` | `spRep_CobranzaFecha` | `dsrCobranzaFecha` | Media |
| `frmRepResultadoOperativo.frm` | `spRep_ResultadoOperativo` | `dsrResultadosOperativos` | Media |
| `frmRepRotacionMesa.frm` | `spRep_Rotacion` | `dsrRotacionMesa` | Baja |

### Tipo 2 — Grilla de Consulta (→ DataGridView + usp_*/SELECT directo)

Formularios que muestran datos en grilla sin generar reporte imprimible.

| Formulario VB6 | Origen datos | Prioridad |
|---|---|---|
| `frmRepClieFrecuentes.frm` | `spRep_RepClieFrecuentes` | Media |
| `frmRepClienteDeuda.frm` | `spRep_CuentasCobrar` | Media |
| `frmRepProductosNoEnlazados.frm` | `spRep_ProductosNoEnlazados` | Baja |

### Tipo 3 — Exportación a Excel / Texto

Formularios que generan archivos de salida.

| Formulario VB6 | Origen datos | Prioridad |
|---|---|---|
| `frmGenerarsunat.frm` | `spRep_Sunattxt` | Media (legal) |
| `frmRepRegistroVenta.frm` (modo SUNAT) | `spRep_RegVentaSunat*` | Media (legal) |

---

## Dependencias

- **Seguridad:** RBAC sobre `TACCESO` (módulo = `CONSULTA`)
- **SPs:** ver `docs/migration/reports/sp-report-contracts.md`
- **Plantillas:** `modern-net8/src/Inforest.Desktop/reports/templates/`
- **Parámetros del sistema:** `TPARAMETRO` — algunos filtros dependen de `nHoraCorte`, `nDiasInventario`

---

## Inventario de los 10 Formularios Prioritarios

Según frecuencia de referencia a SPs de reporte en VB6:

| # | Formulario | SP más referenciado | Frecuencia |
|---|---|---|---|
| 1 | `frmRepLiquidacion.frm` | `spRep_Liquidacion` | 6 veces |
| 2 | `frmRepRegistroVenta.frm` | `spRep_RegVenta*` (varios) | 5 veces |
| 3 | `frmRepVentaMensualIntegrado.frm` | `spRep_VentaMensualIntegrado` | 3 veces |
| 4 | `frmRepPaloteoVentaIntegrado.frm` | `spRep_PaloteoVentaIntegrado` | 2 veces |
| 5 | `frmRepFormaPagoIntegrado.frm` | `spRep_FormaPagoIntegrado` | 2 veces |
| 6 | `frmRepInsumoVentas.frm` | `spRep_PaloteoInsumo` | 2 veces |
| 7 | `frmRepVentaFecha.frm` | `spRep_VentaFecha` | 2 veces |
| 8 | `frmRepControlMotorizado.frm` | `spRep_PlanillaMovilidadGeneral` | 2 veces |
| 9 | `frmRepPlanillaMovilidadMotorizado.frm` | `spRep_PlanillaMovilidad` | 2 veces |
| 10 | `frmRepComanda.frm` | `spRep_Comanda` | 1 vez |

---

## Estado de Migración

| Componente | Estado |
|---|---|
| Inventario formularios Consulta.exe | COMPLETED |
| Análisis SP contracts | COMPLETED |
| IReporteRepository + ReporteRepository | COMPLETED |
| Handlers Grupo A + B + C core | COMPLETED |
| Formularios WinForms Grupo A (5/10) | IN_PROGRESS |
| Plantillas FastReport .frx | NOT_STARTED |
| Formularios WinForms Grupo B | NOT_STARTED |
| Exportación SUNAT | NOT_STARTED |

---

## Referencias

- `docs/migration/reports/report-inventory.md`
- `docs/migration/reports/sp-report-contracts.md`
- `modern-net8/src/Inforest.Application/Reportes/`
- `modern-net8/src/Inforest.Infrastructure/Reportes/`
- `modern-net8/src/Inforest.Desktop/Forms/Reportes/`
