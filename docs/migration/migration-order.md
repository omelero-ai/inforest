# Orden de Migración Recomendado — INFOREST

## Justificación del orden

1. **Infrastructure primero**: no existe código .NET 8; se necesita el armazón base para encapsular ADO/SQL Server, hardware Windows y configuración.
2. **Database temprano**: la base Legacy contiene 150 SP con lógica real de negocio; el contrato SQL condiciona Domain/Application.
3. **Shared antes de módulos funcionales**: tipos comunes, errores, configuración, auditoría y catálogos deben estabilizarse una sola vez.
4. **Seguridad antes de operación**: login, RBAC y sesión son prerrequisitos para cualquier pantalla operativa.
5. **Maestros antes de transacciones**: productos, grupos, clientes, mesas y catálogos alimentan pedidos, venta y caja.
6. **Configuración (`TPARAMETRO`/`TCAJA`) antes del POS**: los flags alteran el flujo de venta, caja, impresión, delivery y hardware.
7. **Turno y día contable antes de cobrar/vender**: el POS depende del turno abierto y de la fecha operativa correcta.
8. **Pedidos simples antes de venta/documentos**: primero se estabiliza la creación/edición de pedidos, luego la emisión documental.
9. **Caja/pagos después de documentos**: los pagos requieren pedido/documento/turno ya modelados.
10. **Reportes después del core transaccional**: conviene reutilizar SPs y validar primero la consistencia operativa.
11. **Integraciones hardware al final del core**: la UI y el dominio deben existir antes de cablear impresoras, PinPad, fiscal o biometría.
12. **KDS/Delivery al final**: dependen de pedidos, áreas de producción, caja, configuración y sincronización multi-local.

## Tabla de workstreams y prioridad

| Orden | Módulo / Workstream | Complejidad | Dependencias | Prioridad | Justificación |
|---:|---|---|---|---|---|
| 1 | Infrastructure | Alta | ADR-001, ADR-002, ADR-003 | Crítica | Base técnica para SQL Server, WinForms, hardware y configuración. |
| 2 | Database | Alta | Infrastructure | Crítica | Define acceso a 126 tablas, 116 vistas y 150 SP. |
| 3 | Shared | Media | Infrastructure, Database | Alta | Tipos comunes, utilitarios, errores, auditoría y contratos transversales. |
| 4 | Seguridad | Media | Shared, Database | Alta | Login, hash BCrypt, RBAC y sesión habilitan toda operación. |
| 5 | Maestros (Productos/Grupos/Clientes) | Media | Shared, Seguridad, Database | Alta | Alimenta pedidos, caja, reportes e inventario. |
| 6 | Configuración (`TPARAMETRO`/`TCAJA`) | Alta | Shared, Database, Maestros | Alta | Los flags cambian el comportamiento del sistema por local/caja/país. |
| 7 | Turno / Día Contable | Alta | Seguridad, Configuración | Alta | Condición previa para caja, venta y reportes contables. |
| 8 | Pedidos Simples | Alta | Maestros, Configuración, Turno | Alta | Núcleo operativo mínimo para crear/editar pedidos. |
| 9 | Venta / Documentos | Muy alta | Pedidos Simples, Turno, Configuración | Crítica | Emisión de comprobantes, impuestos, descuentos y numeración. |
| 10 | Caja / Pagos | Muy alta | Venta / Documentos, Turno | Crítica | Cobranza multimedio, propinas, anticipos y cuadre. |
| 11 | Reportes | Media | Venta / Documentos, Caja / Pagos | Media | Reutiliza SPs `spRep_*` y consolida validación operativa. |
| 12 | Integraciones Hardware | Muy alta | WinForms base, Venta, Caja | Alta | Impresoras, cajón, fiscal, PinPad, biometría, balanza. |
| 13 | KDS / Delivery / Centralizados | Muy alta | Pedidos, Configuración, Hardware, Reportes | Alta | Requiere pedidos estables, áreas de producción y sincronización. |

## Orden recomendado detallado

1. **Infrastructure**
2. **Database**
3. **Shared**
4. **Seguridad**
5. **Maestros (Productos / Grupos / Clientes)**
6. **Configuración (`TPARAMETRO` / `TCAJA`)**
7. **Turno / DíaContable**
8. **PedidosSimples**
9. **Venta / Documentos**
10. **Caja / Pagos**
11. **Reportes**
12. **Integraciones Hardware**
13. **KDS / Delivery**

## Dependencias críticas a respetar

- `Seguridad` depende de `Infrastructure + Database + Shared`.
- `Maestros` depende de `Seguridad` para permisos y de `Database` para contratos SQL.
- `Configuración` depende de `Maestros` porque varios catálogos/flags gobiernan visibilidad y comportamiento.
- `Turno / DíaContable` depende de `Configuración` (`TPARAMETRO`, `TCAJA`) y de `Seguridad` (usuario/caja).
- `PedidosSimples` depende de `Maestros + Configuración + Turno`.
- `Venta / Documentos` depende de `PedidosSimples` y del modelo de impuestos/correlativos.
- `Caja / Pagos` depende de `Venta / Documentos`, `Turno` y hardware de cobro.
- `Reportes` depende del cierre semántico de `Venta / Documentos` y `Caja / Pagos`.
- `Integraciones Hardware` y `KDS / Delivery` deben cerrarse después del núcleo transaccional para evitar re-trabajo.
