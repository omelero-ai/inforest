# Glosario — INFOREST

> Términos y conceptos del sistema INFOREST. Español/Técnico.

---

## Términos del Negocio

| Término | Definición |
|---|---|
| **Adición** | Proceso de agregar ítems a un pedido ya existente desde un terminal independiente (`Adicion.exe`) |
| **Área de impresión** | Zona de cocina/barra a la que se envía la comanda de un ítem específico (configurada en `TPRODUCTOAREA`) |
| **Caja** | Terminal de cobro. Cada caja tiene configuración independiente en `TCAJA` |
| **Canal de venta** | Tipo de atención: 01=Local (salón), 02=Delivery, 03=Llevar, 04=Canal4, 05=Canal5 |
| **Central de pedidos** | Módulo que centraliza todos los pedidos de delivery (`Despachador.exe`) |
| **Cierre de turno** | Proceso de liquidación de la caja al finalizar el turno de trabajo |
| **Comanda** | Impresión de la orden enviada a cocina para preparación |
| **Día contable** | Período contable activo. Puede ser automático (hora de corte) o manual |
| **Descargo** | Reducción automática de stock de insumos al registrar una venta |
| **Delivery** | Modalidad de venta a domicilio |
| **Insumo** | Ingrediente o materia prima que se descarga del inventario cuando se vende un producto |
| **KDS** | Kitchen Display System — pantalla de cocina que muestra las órdenes pendientes |
| **Llevar** | Modalidad de venta para consumo fuera del local (take away) |
| **Mozo** | Mesero/camarero que atiende las mesas |
| **Motorizado** | Repartidor de delivery a domicilio |
| **Paloteo** | Reporte de conteo por producto (cuántas veces se vendió cada producto) |
| **Precuenta** | Pre-factura o cuenta previa al cobro definitivo |
| **Propina** | Gratificación adicional al servicio incluida en el pago |
| **Salón** | Área del restaurante donde se atiende a los comensales |
| **Sub-producto** | Componente de un combo o producto compuesto |
| **Turno** | Período de trabajo de un cajero/operador. Se registra en `MTURNO` |
| **Vale de consumo** | Voucher de consumo como medio de pago |

---

## Términos Técnicos del Legacy

| Término | Definición |
|---|---|
| **ADO** | ActiveX Data Objects — tecnología de acceso a datos de VB6 |
| **ADODB.Command** | Objeto ADO para ejecutar stored procedures parametrizados |
| **ADODB.Connection** | Objeto ADO para conexión a SQL Server |
| **BAS** | Extensión de módulo VB6 (`.bas`) — contiene código compartido |
| **CLS** | Extensión de clase VB6 (`.cls`) — define objetos con propiedades y métodos |
| **clsComando** | Clase VB6 wrapper de ADODB.Command para invocar SPs con timeout de 600s |
| **Crystal Reports** | Motor de reportes usado en el Legacy (versiones 6 y 9) |
| **DLL3500** | Terminal de pago con tarjeta (PinPad). Interfaz vía `CAJA_PINPAD.dll` |
| **FRM** | Extensión de formulario VB6 (`.frm`) — contiene UI y lógica mezcladas |
| **FRX** | Extensión de recursos binarios de formulario VB6 (`.frx`) |
| **INI** | Archivo de configuración Windows (`INFOREST.INI`, `ALMACEN.INI`, etc.) |
| **modDeclaracion** | Módulo VB6 con 543 variables globales — estado de sesión del sistema |
| **modProcedimiento** | Módulo VB6 con utilidades núcleo del sistema |
| **OCX** | Componente ActiveX/COM usado en VB6 |
| **SQLOLEDB** | Provider ADO para SQL Server |
| **Sub Main** | Punto de entrada de cada ejecutable VB6 |
| **VBP** | Archivo de proyecto VB6 (`.vbp`) |

---

## Términos Técnicos de BD

| Término | Definición |
|---|---|
| **APEDIDO** | Tabla de auditoría de ítems de pedido anulados |
| **CPEDIDO** | Tabla de detalle de combos en pedido |
| **DDOCUMENTO** | Tabla de detalle de documento de venta |
| **DPEDIDO** | Tabla de detalle de pedido |
| **DPAGODOCUMENTO** | Tabla de pagos registrados por documento |
| **INFOREST** | Base de datos principal del sistema |
| **INFSEGURIDAD** | Base de datos de seguridad y auditoría |
| **MCIERRE** | Tabla de cierre contable |
| **MDOCUMENTO** | Tabla de cabecera de documento de venta |
| **MPEDIDO** | Tabla de cabecera de pedido |
| **MTURNO** | Tabla de turno de caja |
| **TPARAMETRO** | Tabla maestra de parámetros del sistema (100+ flags de configuración) |
| **TCAJA** | Tabla de configuración de caja (30+ flags de comportamiento) |
| **TDIACONTABLE** | Tabla del día contable activo |
| **TPRODUCTO** | Tabla maestra de productos |

---

## Términos de la Migración

| Término | Definición |
|---|---|
| **ADR** | Architecture Decision Record — documento de decisión arquitectónica |
| **Big Bang** | Estrategia de migración donde todo se reescribe y despliega a la vez |
| **BR** | Business Rule — regla de negocio |
| **GAP** | Brecha entre el Legacy y el Target identificada |
| **Legacy** | El sistema existente en VB6 que se está migrando |
| **Strangler Fig** | Estrategia de migración donde los módulos se reemplazan gradualmente |
| **Target** | La nueva implementación .NET 8 |
| **Trazabilidad** | Documentación que vincula cada componente Legacy con su equivalente Target |

---

## Siglas

| Sigla | Significado |
|---|---|
| ADO | ActiveX Data Objects |
| AFIP | Administración Federal de Ingresos Públicos (Argentina) |
| COM | Component Object Model (Microsoft) |
| FE | Facturación Electrónica |
| IGV | Impuesto General a las Ventas (Perú) |
| INI | Initialization file (archivo de configuración) |
| IVA | Impuesto al Valor Agregado |
| KDS | Kitchen Display System |
| OCX | OLE Control Extension (control ActiveX) |
| POS | Point Of Sale (Punto de Venta) |
| RBAC | Role-Based Access Control |
| RUC | Registro Único de Contribuyentes (Perú, Ecuador) |
| SIN | Servicio de Impuestos Nacionales (Bolivia) |
| SII | Servicio de Impuestos Internos (Chile) |
| SP | Stored Procedure |
| SRI | Servicio de Rentas Internas (Ecuador) |
| SUNAT | Superintendencia Nacional de Aduanas y de Administración Tributaria (Perú) |
| TVS | Table Value System (BlueVision) |
| VB6 | Visual Basic 6 |
| VBP | Visual Basic Project |

---

*Glosario en construcción. Se añaden términos a medida que se profundiza el análisis.*
