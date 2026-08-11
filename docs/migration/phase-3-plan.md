# Plan de Fase 3 — Migración controlada y secuencial

> Estado del documento: `ANALYSIS`
>
> Última actualización: 2026-08-11
>
> Propósito: definir las etapas de ejecución de la Fase 3 sin iniciar todavía ninguna implementación.

---

## 1. Base de análisis

Este plan se construye con evidencia técnica de:

- `/home/runner/work/inforest/inforest/docs/README.md`
- `/home/runner/work/inforest/inforest/docs/migration/migration-status.md`
- `/home/runner/work/inforest/inforest/docs/migration/migration-strategy.md`
- `/home/runner/work/inforest/inforest/docs/migration/migration-order.md`
- `/home/runner/work/inforest/inforest/docs/migration/business-rules.md`
- `/home/runner/work/inforest/inforest/docs/migration/known-gaps.md`
- `/home/runner/work/inforest/inforest/docs/migration/legacy-inventory.md`
- `/home/runner/work/inforest/inforest/docs/migration/database/sql-inventory.md`
- `/home/runner/work/inforest/inforest/docs/migration/traceability/business-logic-matrix.md`
- `/home/runner/work/inforest/inforest/docs/architecture/legacy-architecture.md`
- `/home/runner/work/inforest/inforest/docs/architecture/architecture-decisions.md`
- `/home/runner/work/inforest/inforest/docs/architecture/technology-selection.md`
- `/home/runner/work/inforest/inforest/docs/modules/restaurante/README.md`
- `/home/runner/work/inforest/inforest/legacy-restaurant/README.md`
- `/home/runner/work/inforest/inforest/legacy-restaurant/database-sql-server/README.md`
- `/home/runner/work/inforest/inforest/legacy-restaurant/restaurant-vb6/Clases/clsComando.cls`
- `/home/runner/work/inforest/inforest/legacy-restaurant/restaurant-vb6/Modulos/modPuntoVenta.bas`
- `/home/runner/work/inforest/inforest/legacy-restaurant/restaurant-vb6/Modulos/modKDS.bas`
- `/home/runner/work/inforest/inforest/legacy-restaurant/restaurant-vb6/Modulos/modImpresoraFiscal.bas`
- `/home/runner/work/inforest/inforest/legacy-restaurant/restaurant-vb6/Modulos/modProcedimientoNuevo.bas`
- `/home/runner/work/inforest/inforest/legacy-restaurant/database-sql-server/1. Estructura.sql`
- `/home/runner/work/inforest/inforest/legacy-restaurant/database-sql-server/5. SP.sql`

---

## 2. Criterios que definen el orden recomendado

1. **La Fase 3 no debe comenzar por formularios**. El Legacy depende de estado global, SQL Server, SPs, configuración por caja, impresión, licencias e integraciones transversales.
2. **Infrastructure Base por sí sola no es suficiente como primera etapa cerrada**. Antes de cualquier módulo funcional deben quedar separados:
   - conexión,
   - ejecución de comandos,
   - transacciones,
   - configuración,
   - sesión,
   - auditoría,
   - logging,
   - errores,
   - contratos SQL.
3. **`clsComando.cls` no puede mapearse automáticamente a un único componente .NET**. Debe descomponerse en responsabilidades distintas antes de diseñar equivalentes.
4. **`TPARAMETRO` y `TCAJA` son prerrequisitos operativos**, no simples catálogos. Alteran venta, impuestos, impresión, delivery, KDS, hardware y fecha operativa.
5. **Hardware e impresión no deben quedar completamente al final**. El POS no puede considerarse funcional sin impresión base, correlativos y periféricos críticos de caja.
6. **Reportes y consultas deben ir después del núcleo transaccional mínimo**, reutilizando SPs y vistas ya estabilizados.
7. **Delivery, KDS y centralización dependen de pedidos, configuración, áreas, impresión y caja**.

---

## 3. Orden recomendado de ejecución

1. P3-01 — Re-baseline de arquitectura ejecutable y gobierno técnico
2. P3-02 — Contrato SQL y ejecución de datos Legacy
3. P3-03 — Seguridad, sesión, auditoría y licenciamiento
4. P3-04 — Configuración operativa global y por caja
5. P3-05 — Catálogos y maestros operativos
6. P3-06 — Turno, día contable y pedido base
7. P3-07 — Venta, documento, impresión base y pagos core
8. P3-08 — Cocina, comanda, áreas y KDS
9. P3-09 — Delivery, despacho, motorizados y centralización
10. P3-10 — Reportes, consultas y salidas operativas
11. P3-11 — Periféricos especializados e integraciones por país
12. P3-12 — Validación integral, hardening y cierre de Fase 3

---

## 4. Etapas propuestas

### Etapa P3-01

- **ID:** P3-01
- **Nombre:** Re-baseline de arquitectura ejecutable y gobierno técnico
- **Objetivo:** confirmar que la base actual de `modern-net8/` respeta ADRs, reglas de dependencia y restricciones de migración antes de agregar funcionalidad operativa.
- **Alcance:** solución, proyectos, DI, configuración base, logging estructurado, convenciones de trazabilidad, políticas de errores, revisión del código preliminar ya existente.
- **Componentes Legacy involucrados:**
  - `Modulos/modPuntoVenta.bas`
  - `Modulos/modDeclaracion.bas`
  - `Modulos/modProcedimiento.bas`
  - `Clases/clsComando.cls`
- **Scripts SQL involucrados:**
  - `1. Estructura.sql`
  - `5. SP.sql`
  - `opcionales/Seguridad.sql`
- **Dependencias:** ADR-001 a ADR-008 aceptados; estructura existente en `modern-net8/`.
- **Componentes .NET que se crearán:**
  - composición de solución revisada,
  - bootstrap de Desktop,
  - módulos base de observabilidad,
  - convenciones de trazabilidad,
  - contratos base de errores y resultado,
  - baseline de pruebas.
- **Módulos afectados:** transversal a todos.
- **Prerrequisitos:** ninguno dentro de Fase 3.
- **Riesgos:**
  - reutilizar sin revisión componentes prematuros del baseline actual;
  - contaminar Domain/Application con decisiones tomadas sin evidencia Legacy.
- **Criterios de entrada:**
  - ADRs aceptados;
  - inventarios y matrices de Fase 1/Fase 2 disponibles.
- **Criterios de salida:**
  - baseline .NET validado o corregido;
  - reglas de dependencia `Desktop → Application → Domain ← Infrastructure` confirmadas;
  - convenciones de trazabilidad y validación definidas para el resto de etapas.
- **Pruebas requeridas:**
  - build de solución,
  - pruebas base existentes,
  - validación de arranque de host/DI/configuración.
- **Dependencias con otras etapas:** bloquea P3-02 a P3-12.
- **Estado:** NOT_STARTED

### Etapa P3-02

- **ID:** P3-02
- **Nombre:** Contrato SQL y ejecución de datos Legacy
- **Objetivo:** establecer el modelo seguro y trazable de acceso a SQL Server y SPs Legacy sin asumir equivalencias simplistas desde `clsComando.cls`.
- **Alcance:** conexiones, comandos, parámetros, transacciones, timeouts, manejo de errores SQL, lectura de recordsets, mapeo de resultados y observabilidad de acceso a datos.
- **Componentes Legacy involucrados:**
  - `Clases/clsComando.cls`
  - formularios que invocan `spIns_MPEDIDO`, `spUpd_MPEDIDO`, `usp_GenObtieneDiaContable`
  - `Modulos/modPuntoVenta.bas`
- **Scripts SQL involucrados:**
  - `5. SP.sql`
  - `1. Estructura.sql`
  - `4. Vistas.sql`
- **Dependencias:** P3-01.
- **Componentes .NET que se crearán:**
  - `IDbConnectionFactory` si la evidencia lo justifica,
  - ejecutor de comandos/SPs,
  - binder de parámetros,
  - política de transacciones,
  - mapper de lectura Dapper/ADO,
  - traducción de errores de infraestructura.
- **Módulos afectados:** transversal a POS, Caja, Administración, Consulta, Delivery, Motorizados.
- **Prerrequisitos:** baseline técnico aprobado.
- **Riesgos:**
  - reducir `clsComando` a “solo conexión” y perder responsabilidades reales;
  - romper semántica de SPs con output params, recordsets o timeouts.
- **Criterios de entrada:**
  - responsabilidades de `clsComando.cls` catalogadas;
  - SPs críticos priorizados.
- **Criterios de salida:**
  - separación explícita entre conexión, comando, transacción, errores y logging;
  - patrón estándar para invocar SPs Legacy desde .NET;
  - contratos para lectura de vistas y tablas consolidados.
- **Pruebas requeridas:**
  - pruebas unitarias de binder/mapeo,
  - pruebas de integración contra SQL Server para SPs piloto,
  - pruebas de error/timeout/transacción.
- **Dependencias con otras etapas:** depende de P3-01; bloquea P3-03 a P3-12.
- **Estado:** NOT_STARTED

### Etapa P3-03

- **ID:** P3-03
- **Nombre:** Seguridad, sesión, auditoría y licenciamiento
- **Objetivo:** migrar el armazón transversal de autenticación, autorización, sesión operativa, auditoría y validación de licencia.
- **Alcance:** acceso de usuarios, RBAC, sesión por caja/terminal, auditoría hacia `INFSEGURIDAD`, transición desde cifrado débil y análisis del dongle/licenciamiento.
- **Componentes Legacy involucrados:**
  - `Formularios/frmAcceso.frm`
  - `Clases/ClsSeguridad.cls`
  - `Clases/License.cls`
  - `Modulos/modSeguridadInfhotel.bas`
  - `Modulos/modAuditoriaIntegral.bas`
- **Scripts SQL involucrados:**
  - `opcionales/Seguridad.sql`
  - tablas `TUSUARIO`, `TGRUPOUSUARIO`, `TACCESO`, `TGRUPOACCESO`, `THARDKEY`, `MMOVIMIENTO`, `MMOVIMIENTOACCESO`, `TPARAMETROVERSION`
- **Dependencias:** P3-02.
- **Componentes .NET que se crearán:**
  - servicios de autenticación,
  - RBAC,
  - contexto de usuario/sesión,
  - auditoría de acceso y cambios,
  - adaptador de licenciamiento,
  - estrategia de transición de passwords.
- **Módulos afectados:** todos los ejecutables.
- **Prerrequisitos:** acceso a datos Legacy estabilizado.
- **Riesgos:**
  - replicar cifrado inseguro del Legacy;
  - subestimar dependencia del dongle/licencia en el arranque real.
- **Criterios de entrada:**
  - infraestructura de comandos y consultas operativa;
  - reglas BR-006 y BR-014 en análisis detallado.
- **Criterios de salida:**
  - login y permisos funcionales con trazabilidad;
  - auditoría mínima equivalente disponible;
  - decisión documentada para convivencia con licencias Legacy.
- **Pruebas requeridas:**
  - unit tests de permisos y sesiones,
  - integration tests con tablas de seguridad,
  - validación comparativa de permisos con Legacy.
- **Dependencias con otras etapas:** depende de P3-02; bloquea P3-04 a P3-12.
- **Estado:** NOT_STARTED

### Etapa P3-04

- **ID:** P3-04
- **Nombre:** Configuración operativa global y por caja
- **Objetivo:** modelar correctamente `TPARAMETRO`, `TCAJA`, `TCAJACANALVENTA`, `TTIPODOCUMENTOIMPRESORA` y configuración de terminal/periféricos sin reducirlos a simples lecturas CRUD.
- **Alcance:** parámetros globales, perfil por caja, correlativos, flags por canal, impresión, facturación electrónica, día contable, KDS, delivery, periféricos y centralización.
- **Componentes Legacy involucrados:**
  - `Modulos/modPuntoVenta.bas`
  - `Modulos/modProcedimiento.bas`
  - `Formularios/frmParametro.frm`
  - `Formularios/frmConfiguracionParametro.frm`
  - `Formularios/frmTerminalDetalle.frm`
- **Scripts SQL involucrados:**
  - `1. Estructura.sql`
  - `5. SP.sql`
  - `6. Actualiza.sql`
  - `opcionales/scriptPeruAlIniciar.sql`
  - `opcionales/scriptChileAlIniciar.sql`
  - `opcionales/scriptBoliviaAlIniciar.sql`
  - `opcionales/scriptEcuadorAlIniciar.sql`
  - `opcionales/scriptArgentinaAlIniciar.sql`
  - `opcionales/scriptEspanaAlIniciar.sql`
- **Dependencias:** P3-03.
- **Componentes .NET que se crearán:**
  - modelo de configuración runtime,
  - perfiles de caja/terminal,
  - políticas de feature flags por país/local,
  - configuración de emisión e impresión,
  - proveedores de parámetros con cache invalidation cuando la evidencia lo requiera.
- **Módulos afectados:** POS, Caja Rápida, Administración, Consulta, Despacho, Motorizados.
- **Prerrequisitos:** seguridad y sesión definidas.
- **Riesgos:**
  - crear `IParametroService` sin analizar usos y reglas incrustadas en `TPARAMETRO`;
  - perder diferencias por país, caja, canal o terminal.
- **Criterios de entrada:**
  - inventario de columnas críticas de `TPARAMETRO` y `TCAJA`;
  - usos VB6 y SQL catalogados para flags prioritarios.
- **Criterios de salida:**
  - clasificación explícita de parámetros por responsabilidad;
  - configuración global, por caja y por terminal disponible para etapas operativas;
  - correlativos y perfiles de impresión modelados.
- **Pruebas requeridas:**
  - unit tests de políticas de configuración,
  - integration tests de lectura y persistencia controlada,
  - matrices de prueba por país/local/caja/canal.
- **Dependencias con otras etapas:** depende de P3-03; bloquea P3-05 a P3-12.
- **Estado:** NOT_STARTED

### Etapa P3-05

- **ID:** P3-05
- **Nombre:** Catálogos y maestros operativos
- **Objetivo:** migrar los catálogos y maestros que alimentan venta, caja, cocina, delivery e impresión.
- **Alcance:** productos, grupos, subgrupos, clientes, mesas, áreas, tipos de pedido, tipos de pago, operadores, tarjetas, unidades de negocio y catálogos auxiliares.
- **Componentes Legacy involucrados:**
  - `Formularios/frmProducto.frm`
  - `Formularios/frmProductoBoton.frm`
  - `Formularios/frmCliente.frm`
  - `Formularios/frmMesa.frm`
  - `Formularios/frmArea.frm`
  - `Modulos/modProcedimiento.bas`
- **Scripts SQL involucrados:**
  - `1. Estructura.sql`
  - `4. Vistas.sql`
  - `5. SP.sql`
- **Dependencias:** P3-04.
- **Componentes .NET que se crearán:**
  - entidades de maestros,
  - queries de catálogos,
  - servicios de precio/visibilidad por canal,
  - repositorios/queries de lectura operativa.
- **Módulos afectados:** Administración, POS, Caja Rápida, Adición, Delivery.
- **Prerrequisitos:** configuración global y de caja resuelta.
- **Riesgos:**
  - migrar maestros sin preservar visibilidad por canal y unidad de negocio;
  - omitir catálogos que condicionan reglas en SPs y formularios.
- **Criterios de entrada:**
  - catálogos prioritarios identificados;
  - BR-001, BR-002, BR-005, BR-011, BR-014 refinadas.
- **Criterios de salida:**
  - maestros base disponibles para flujos operativos;
  - precios/canales/áreas alineados con el Legacy;
  - catálogos críticos trazados.
- **Pruebas requeridas:**
  - unit tests de reglas de precio e impuestos,
  - integration tests de consultas y vistas,
  - validación funcional con pantallas Legacy relevantes.
- **Dependencias con otras etapas:** depende de P3-04; bloquea P3-06 a P3-12.
- **Estado:** NOT_STARTED

### Etapa P3-06

- **ID:** P3-06
- **Nombre:** Turno, día contable y pedido base
- **Objetivo:** establecer el mínimo operativo transaccional para abrir sesión de caja y registrar/editar pedidos equivalentes al Legacy.
- **Alcance:** `MTURNO`, `TDIACONTABLE`, `MPEDIDO`, `DPEDIDO`, `CPEDIDO`, reglas de mesa/salón, correlativos, día contable y estado de pedido.
- **Componentes Legacy involucrados:**
  - `Modulos/modPuntoVenta.bas`
  - `Clases/clsDiaContable.cls`
  - `Formularios/frmVenta.frm`
  - `Formularios/frmCajaRapida.frm`
  - `Formularios/frmAdicion.frm`
- **Scripts SQL involucrados:**
  - `1. Estructura.sql`
  - `5. SP.sql` (`spIns_MPEDIDO`, `spUpd_MPEDIDO`, `usp_GenObtieneDiaContable` y relacionados)
- **Dependencias:** P3-05.
- **Componentes .NET que se crearán:**
  - aggregate de pedido,
  - servicios de turno y día contable,
  - commands de creación/edición de pedido,
  - políticas de consistencia de mesa/salón/canal.
- **Módulos afectados:** Punto de Venta, Caja Rápida, Adición.
- **Prerrequisitos:** maestros y configuración disponibles.
- **Riesgos:**
  - perder side effects SQL como correlativos y actualización de `TPEDIDOMESA`;
  - subestimar reglas de fecha operativa y estado de turno.
- **Criterios de entrada:**
  - BR-003, BR-007, BR-009, BR-SQL-001 a BR-SQL-008 analizadas;
  - SPs de pedido clasificados.
- **Criterios de salida:**
  - apertura/cierre de turno base disponible;
  - pedido simple y edición equivalentes al Legacy;
  - comportamiento comparado contra VB6 y SQL.
- **Pruebas requeridas:**
  - unit tests de dominio de pedido/turno,
  - integration tests con SPs de pedido,
  - escenarios comparativos mesa/salón/canal/día contable.
- **Dependencias con otras etapas:** depende de P3-05; bloquea P3-07 a P3-12.
- **Estado:** NOT_STARTED

### Etapa P3-07

- **ID:** P3-07
- **Nombre:** Venta, documento, impresión base y pagos core
- **Objetivo:** cerrar el circuito mínimo de emisión y cobro con correlativos, documento, pagos y salidas impresas base.
- **Alcance:** `MDOCUMENTO`, `DDOCUMENTO`, `DPAGODOCUMENTO`, `DPAGOTARJETA`, impresión de documento/precuenta/comanda básica, correlativos de `TTIPODOCUMENTOIMPRESORA`, reglas de caja para cobro antes de impresión.
- **Componentes Legacy involucrados:**
  - `Formularios/frmVenta.frm`
  - `Formularios/frmPago.frm`
  - `Formularios/frmCaja.frm`
  - `Modulos/modProcedimiento.bas`
  - `Formularios/frmPrecuenta.frm`
  - `Formularios/frmPrecuentaImpresora.frm`
- **Scripts SQL involucrados:**
  - `1. Estructura.sql`
  - `5. SP.sql`
  - vistas `vDocumento`, `vDocumentoImpresora`, `vPreCuenta`
- **Dependencias:** P3-06.
- **Componentes .NET que se crearán:**
  - commands de emisión,
  - servicios de pago base,
  - adaptadores de impresión base,
  - queries de dataset de impresión,
  - políticas de correlativo y numeración documental.
- **Módulos afectados:** Punto de Venta, Caja Rápida, Consulta operacional.
- **Prerrequisitos:** pedido y turno operativos.
- **Riesgos:**
  - dejar impresión para después y bloquear validación real del POS;
  - mezclar impresión, reglas de caja y numeración en una sola capa.
- **Criterios de entrada:**
  - pedido base validado;
  - `TTIPODOCUMENTOIMPRESORA` y reglas de caja trazadas.
- **Criterios de salida:**
  - pedido → documento → pago base → impresión funcional;
  - correlativos persistidos correctamente;
  - regla de pago antes de impresión controlada.
- **Pruebas requeridas:**
  - unit tests de políticas de cobro y numeración,
  - integration tests de persistencia documental,
  - pruebas operativas de impresión y reimpresión base.
- **Dependencias con otras etapas:** depende de P3-06; bloquea P3-08 a P3-12.
- **Estado:** NOT_STARTED

### Etapa P3-08

- **ID:** P3-08
- **Nombre:** Cocina, comanda, áreas y KDS
- **Objetivo:** migrar el flujo de producción que deriva pedidos hacia comandas impresas y/o KDS según áreas y configuración.
- **Alcance:** routing por área, impresión de cocina, `TPRODUCTOAREA`, `TAREAIMPRESORA`, `DPEDIDOKDS`, XML KDS y variantes multi-modelo.
- **Componentes Legacy involucrados:**
  - `Modulos/modKDS.bas`
  - `Formularios/frmCheffControl.frm`
  - `Formularios/frmKDSConfiguracion.frm`
  - `Formularios/frmMensajeCocina.frm`
  - `Formularios/frmVenta.frm`
- **Scripts SQL involucrados:**
  - `1. Estructura.sql`
  - `5. SP.sql` (`USP_KDS_*`)
  - `4. Vistas.sql`
- **Dependencias:** P3-07.
- **Componentes .NET que se crearán:**
  - router de producción,
  - adaptador KDS,
  - adaptador de comanda por área,
  - políticas de envío a cocina y reenvío.
- **Módulos afectados:** POS, Caja Rápida, Cocina/KDS.
- **Prerrequisitos:** pedido/documento/impresión base disponibles.
- **Riesgos:**
  - no preservar diferencias entre comanda impresa y envío KDS;
  - romper tratamiento de combos, condimentos y estaciones.
- **Criterios de entrada:**
  - BR-010 y BR-011 refinadas;
  - áreas, impresoras y estaciones trazadas.
- **Criterios de salida:**
  - generación de salida a cocina equivalente;
  - KDS y comanda base comparados contra Legacy.
- **Pruebas requeridas:**
  - unit tests de routing por área,
  - integration tests de consultas KDS,
  - validación de XML y prueba operativa con directorios/estaciones.
- **Dependencias con otras etapas:** depende de P3-07; bloquea P3-09 a P3-12.
- **Estado:** NOT_STARTED

### Etapa P3-09

- **ID:** P3-09
- **Nombre:** Delivery, despacho, motorizados y centralización
- **Objetivo:** migrar los flujos externos y multi-local que dependen del núcleo operativo ya estabilizado.
- **Alcance:** central pedidos, despacho, motorizados, administración centralizada, sincronización multi-local y canales externos.
- **Componentes Legacy involucrados:**
  - `Modulos/modDespachador.bas`
  - `Modulos/modMotorizado.bas`
  - `Modulos/modPuntoVenta.bas`
  - `Formularios/frmDespachador.frm`
  - formularios de delivery y motorizados
- **Scripts SQL involucrados:**
  - `1. Estructura.sql`
  - `5. SP.sql` (`usp_AdmCen_*`, `spRep_TiempoDelivery`, integraciones delivery)
  - tablas `TLOCAL`, `TTIENDA`, `TMOTORIZADODATOS`, `TDELIVERYCLIENTE`, `TDELIVERYINVITADO`
- **Dependencias:** P3-08.
- **Componentes .NET que se crearán:**
  - servicios de despacho,
  - sincronización/consulta centralizada,
  - workflows de motorizados,
  - políticas de origen de venta y canal centralizado.
- **Módulos afectados:** Despachador, Motorizados, POS, Administración centralizada.
- **Prerrequisitos:** pedidos, pagos e impresión estabilizados; cocina/KDS resueltos si intervienen en producción.
- **Riesgos:**
  - comportamiento divergente por local y por canal central;
  - dependencias ocultas con configuración y origen de venta.
- **Criterios de entrada:**
  - BR-012 y reglas de origen/canal analizadas;
  - tablas y SPs de delivery clasificadas.
- **Criterios de salida:**
  - flujo delivery/despacho/motorizado base funcional;
  - sincronización centralizada documentada y validada.
- **Pruebas requeridas:**
  - unit tests de políticas de canal/origen,
  - integration tests de sincronización y consultas multi-local,
  - validación comparativa de tiempos/estados de delivery.
- **Dependencias con otras etapas:** depende de P3-08; bloquea P3-10 a P3-12.
- **Estado:** NOT_STARTED

### Etapa P3-10

- **ID:** P3-10
- **Nombre:** Reportes, consultas y salidas operativas
- **Objetivo:** migrar la capa de consulta y reporting reutilizando contratos SQL Legacy ya estabilizados.
- **Alcance:** `spRep_*`, vistas de consulta, formularios `frmRep*`, datasets de impresión operativa y estrategia FastReport .NET.
- **Componentes Legacy involucrados:**
  - `Consulta.vbp`
  - `Administracion.vbp`
  - `Reportes/*.Dsr`
  - formularios `frmRep*`
- **Scripts SQL involucrados:**
  - `4. Vistas.sql`
  - `5. SP.sql` (`spRep_*`, `usp_Inforest_Impresion`, `usp_RepInforest_DescargoVenta`)
  - `8. InfoFact.sql`
- **Dependencias:** P3-09.
- **Componentes .NET que se crearán:**
  - queries de reportes,
  - adaptadores FastReport,
  - datasets de consulta,
  - capa de exportación/salida operativa.
- **Módulos afectados:** Consulta, Administración, POS, Caja.
- **Prerrequisitos:** núcleo transaccional y delivery estabilizados.
- **Riesgos:**
  - migrar reportes antes de cerrar semántica de ventas/caja;
  - subestimar diferencias entre reportes contables, operativos y fiscales.
- **Criterios de entrada:**
  - SPs de reporte priorizados;
  - fuentes operativas validadas en etapas anteriores.
- **Criterios de salida:**
  - primer lote de reportes críticos equivalente;
  - estrategia repetible para Crystal → FastReport;
  - consultas operativas desacopladas del Legacy UI.
- **Pruebas requeridas:**
  - integration tests de `spRep_*`,
  - comparación de datasets Legacy vs .NET,
  - validación visual/operativa de salidas.
- **Dependencias con otras etapas:** depende de P3-09; bloquea P3-11 y P3-12.
- **Estado:** NOT_STARTED

### Etapa P3-11

- **ID:** P3-11
- **Nombre:** Periféricos especializados e integraciones por país
- **Objetivo:** cerrar las dependencias físicas y regulatorias que no deben bloquear el diseño del core, pero sí deben validarse antes del cierre real de la fase.
- **Alcance:** PinPad, CashDro, BlueVision, biometría, impresora fiscal Epson, FE por país, QR y particularidades fiscales/regulatorias.
- **Componentes Legacy involucrados:**
  - `Modulos/DLL3500.bas`
  - `Modulos/modProcedimientoNuevo.bas`
  - `Modulos/modBlueVision.bas`
  - `Modulos/modImpresoraFiscal.bas`
  - `Clases/clsTrama.cls`
  - `Clases/clsxml.cls`
  - `Formularios/frmPago.frm`
  - `Formularios/frmPagoPinPad.frm`
- **Scripts SQL involucrados:**
  - `5. SP.sql`
  - `8. InfoFact.sql`
  - scripts opcionales por país
  - tablas `TCONFIGURAPERIFERICO`, `TCAJATERMINAL`, `TTERMINAL`, `TESTADOBIZLINK`, `TINTEGRACIONES`
- **Dependencias:** P3-10.
- **Componentes .NET que se crearán:**
  - adaptadores de hardware especializado,
  - gateway de FE,
  - políticas por país,
  - drivers/wrappers de integración.
- **Módulos afectados:** POS, Caja, Administración, Consulta, países específicos.
- **Prerrequisitos:** documentos, pagos, impresión base y reportes estabilizados.
- **Riesgos:**
  - dependencias Win32/OCX sin equivalencia directa;
  - variaciones regulatorias por país no cubiertas por una sola implementación.
- **Criterios de entrada:**
  - catálogo de periféricos y flujos por país priorizado;
  - matriz país × local × integración definida.
- **Criterios de salida:**
  - adaptadores críticos validados en ambiente representativo;
  - integraciones por país documentadas y trazadas.
- **Pruebas requeridas:**
  - pruebas de integración end-to-end con hardware,
  - smoke tests por país,
  - validación operativa en ambientes reales o equivalentes.
- **Dependencias con otras etapas:** depende de P3-10; bloquea P3-12.
- **Estado:** NOT_STARTED

### Etapa P3-12

- **ID:** P3-12
- **Nombre:** Validación integral, hardening y cierre de Fase 3
- **Objetivo:** consolidar la equivalencia funcional, registrar gaps y cerrar la fase con trazabilidad completa.
- **Alcance:** validación cruzada contra Legacy, corrección de regresiones, documentación, trazabilidad, matrices de negocio, riesgos remanentes y readiness para ejecución controlada de fases posteriores.
- **Componentes Legacy involucrados:** todos los componentes tocados en P3-01 a P3-11.
- **Scripts SQL involucrados:** todos los scripts y objetos tocados en P3-01 a P3-11.
- **Dependencias:** P3-11.
- **Componentes .NET que se crearán:**
  - suites de validación,
  - utilitarios de comparación,
  - documentación de gaps y cierre,
  - evidencias de trazabilidad final de fase.
- **Módulos afectados:** transversal a todos.
- **Prerrequisitos:** etapas previas ejecutadas y documentadas.
- **Riesgos:**
  - declarar completitud sin comparación suficiente con Legacy;
  - dejar gaps sin documentar en comportamiento, país, hardware o reportes.
- **Criterios de entrada:**
  - entregables de P3-01 a P3-11 disponibles;
  - matrices de reglas y trazabilidad actualizadas por etapa.
- **Criterios de salida:**
  - regresiones corregidas o registradas;
  - trazabilidad actualizada;
  - documentación de fase cerrada;
  - backlog de bloqueantes y siguientes pasos aprobado.
- **Pruebas requeridas:**
  - test suite completa,
  - validación comparativa con Legacy,
  - pruebas integrales por flujo,
  - revisión de seguridad y observabilidad.
- **Dependencias con otras etapas:** depende de P3-11; cierre de Fase 3.
- **Estado:** NOT_STARTED

---

## 5. Dependencias entre etapas

```text
P3-01 → P3-02
P3-02 → P3-03
P3-03 → P3-04
P3-04 → P3-05
P3-05 → P3-06
P3-06 → P3-07
P3-07 → P3-08
P3-08 → P3-09
P3-09 → P3-10
P3-10 → P3-11
P3-11 → P3-12
```

Dependencias críticas adicionales:

- P3-02 depende de descomponer `clsComando.cls` antes de diseñar componentes .NET definitivos.
- P3-04 depende de analizar `TPARAMETRO` y `TCAJA` completos antes de fijar servicios o modelos.
- P3-07 no debe comenzar sin correlativos, impresión base y reglas de pago validadas.
- P3-11 no sustituye P3-07; solo completa periféricos e integraciones especializadas que requieren core estable.

---

## 6. Justificación del orden frente al esquema inicial sugerido

El orden sugerido inicialmente (`Infrastructure Base → Shared/Core → Catálogos → Procesos → Integraciones → Reportes → Hardware → Validación`) no refleja completamente la evidencia del repositorio por estas razones:

1. **`Shared/Core` no puede aislarse correctamente antes del contrato SQL, seguridad y configuración**, porque el core Legacy está gobernado por `clsComando`, `TPARAMETRO`, `TCAJA`, `INFSEGURIDAD` y SPs.
2. **Hardware no puede quedar íntegramente después de reportes**, porque impresión, correlativos de documento y periféricos de caja forman parte del flujo transaccional mínimo.
3. **Configuración merece una etapa propia antes de procesos**, porque `TPARAMETRO` y `TCAJA` gobiernan impuestos, día contable, KDS, delivery, impresión y validaciones.
4. **Delivery/KDS/centralización no deben mezclarse con el primer núcleo transaccional**, ya que dependen de pedidos, caja, áreas, canales y configuración previa.
5. **Reportes deben migrarse después de estabilizar ventas, caja y delivery**, para reutilizar correctamente `spRep_*` y vistas sin fijar semánticas incompletas.

---

## 7. Regla de ejecución de la fase

La Fase 3 debe ejecutarse estrictamente así:

```text
Plan aprobado
  ↓
Ejecutar una sola etapa
  ↓
Validar y documentar
  ↓
Detener
  ↓
Esperar aprobación
  ↓
Continuar con la siguiente etapa
```

No se debe avanzar automáticamente a la siguiente etapa sin aprobación explícita.
