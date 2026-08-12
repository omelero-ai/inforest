# Decisiones Arquitectónicas — INFOREST (ADR)

> Registro formal de decisiones arquitectónicas aceptadas para la migración INFOREST VB6 → .NET 8.

---

## Índice de ADRs

| ADR | Título | Estado |
|---|---|---|
| [ADR-001](#adr-001) | WinForms .NET 8 como cliente POS principal | Accepted |
| [ADR-002](#adr-002) | SQL Server mantenido como motor y contrato de datos | Accepted |
| [ADR-003](#adr-003) | Clean Architecture + CQRS | Accepted |
| [ADR-004](#adr-004) | Strangler Fig Pattern para migración gradual | Accepted |
| [ADR-005](#adr-005) | `appsettings.json` + User Secrets para configuración | Accepted |
| [ADR-006](#adr-006) | BCrypt + RBAC en memoria sobre tablas Legacy | Accepted |
| [ADR-007](#adr-007) | FastReport .NET para reportes | Accepted |
| [ADR-008](#adr-008) | Feature flags por país + tenant ID | Accepted |

---

## ADR-001

**Título:** WinForms .NET 8 como cliente POS principal

**Estado:** Accepted

**Contexto:**
El sistema Legacy es Windows-only, opera en terminales POS con dependencias intensivas de hardware (impresoras térmicas, cajón, PinPad, biometría, impresora fiscal) y concentra 400 formularios `.frm` con flujos altamente interactivos y latencia crítica.

**Problema:**
Definir el tipo de cliente Target que minimice riesgo de migración y preserve compatibilidad con hardware POS y operación offline/local.

**Decisión:**
El cliente principal del nuevo sistema será **WinForms sobre .NET 8**. Se mantendrá un enfoque desktop nativo para los ejecutables operativos (POS, Caja, Administración, Consulta, Despacho, Motorizados), encapsulando el hardware detrás de adaptadores de Infrastructure.

**Alternativas consideradas:**
- Aplicación Web: descartada para el core POS por dependencia intensiva de hardware y sensibilidad a latencia.
- WPF: viable, pero con menor alineación con el modelo de formularios/eventos VB6 y mayor costo de reentrenamiento.
- .NET MAUI: insuficiente para el ecosistema Win32/COM heredado del POS.
- Híbrido web + desktop: se reserva para etapas futuras, no para el primer corte del core operativo.

**Consecuencias:**
- La migración de UI puede hacerse formulario a formulario conservando patrones de interacción cercanos al Legacy.
- Las integraciones Win32/COM siguen siendo factibles durante la transición.
- El sistema Target continúa siendo Windows-first para operación POS.
- La portabilidad multiplataforma se posterga explícitamente.

**Fecha:** 2026-08-11

**Evidencia:**
- `legacy-restaurant/restaurant-vb6/Formularios/` — 400 formularios
- `IFEpson.ocx`, `CAJA_PINPAD.dll`, `sgfplibx.ocx`
- `modImpresoraFiscal.bas`, `DLL3500.bas`, `modBlueVision.bas`, `modKDS.bas`

---

## ADR-002

**Título:** SQL Server mantenido como motor y contrato de datos

**Estado:** Accepted

**Contexto:**
El Legacy opera sobre SQL Server con 126 tablas, 116 vistas y 150 stored procedures detectados en el repositorio actual. Parte relevante de la lógica de negocio está implementada en SPs (`spIns_MPEDIDO`, `spUpd_MPEDIDO`, `spRep_*`, `USP_*`).

**Problema:**
Definir si la migración cambia de motor o preserva SQL Server para reducir riesgo funcional.

**Decisión:**
Se **mantiene SQL Server** como motor de base de datos del Target y como contrato de datos inicial de migración. Los nombres Legacy se preservan durante la fase de coexistencia y los SP críticos pueden reutilizarse temporalmente tras encapsulación en Infrastructure.

**Alternativas consideradas:**
- PostgreSQL: descartado para la primera fase por costo alto de reescritura de SPs y diferencias dialectales.
- Renombrado masivo de tablas: descartado mientras la trazabilidad no esté madura.
- Reescritura inmediata de toda la lógica SQL en C#: descartada por riesgo funcional y falta de baseline .NET.

**Consecuencias:**
- Se habilita migración gradual módulo a módulo.
- Infrastructure deberá abstraer ADO/consultas/SPs sin acoplar el Domain al esquema Legacy.
- La deuda de nombres Legacy se acepta temporalmente y se documenta en trazabilidad/mapping.

**Fecha:** 2026-08-11

**Evidencia:**
- `legacy-restaurant/database-sql-server/1. Estructura.sql`
- `legacy-restaurant/database-sql-server/4. Vistas.sql`
- `legacy-restaurant/database-sql-server/5. SP.sql`

---

## ADR-003

**Título:** Clean Architecture + CQRS

**Estado:** Accepted

**Contexto:**
El Legacy mezcla UI, lógica de negocio y datos dentro de formularios y módulos globales. La nueva solución necesita separación fuerte de responsabilidades para ser migrable, testeable y mantenible.

**Problema:**
Elegir una estructura que permita aislar reglas de negocio del legado tecnológico, sin perder velocidad de implementación.

**Decisión:**
La solución .NET 8 adoptará **Clean Architecture** con capas **Domain / Application / Infrastructure / Presentation**, y usará **CQRS** a nivel de Application para separar comandos operativos y consultas/reportes.

**Alternativas consideradas:**
- N-Layer tradicional: más simple, pero menos explícita para aislar reglas y dependencias externas.
- Monolito sin CQRS: insuficiente para la densidad de reportes/consultas vs comandos operativos.
- Microservicios: descartado por complejidad operacional temprana.
- Event sourcing: descartado para la fase inicial por sobrecosto innecesario.

**Consecuencias:**
- `modern-net8/` deberá estructurarse por capas y módulos funcionales.
- Los SPs/reportes pueden implementarse como Queries sin contaminar el Domain.
- Los workflows operativos (pedido, cobro, cierre, FE) se modelan como Commands.

**Fecha:** 2026-08-11

---

## ADR-004

**Título:** Strangler Fig Pattern para migración gradual

**Estado:** Accepted

**Contexto:**
INFOREST está orientado a operación continua y no existe todavía una base .NET 8. El riesgo de un corte total es alto por amplitud funcional, hardware e integraciones multi-país.

**Problema:**
Definir la estrategia de transición entre Legacy VB6 y el Target .NET 8.

**Decisión:**
La migración seguirá el **Strangler Fig Pattern**: coexistencia controlada de Legacy y .NET 8, reemplazando capacidades en slices verificables (seguridad, maestros, configuración, pedidos, venta, caja, reportes, hardware, KDS/delivery).

**Alternativas consideradas:**
- Big Bang: descartado por alto riesgo operativo.
- Parallel Run total: útil para validación, pero demasiado costoso como estrategia primaria.
- Reescritura por repositorio sin convivencia: no resuelve rollout progresivo.

**Consecuencias:**
- La trazabilidad Legacy→.NET se vuelve artefacto obligatorio de gobierno.
- Los componentes nuevos deben validar equivalencia contra VB6/SQL antes de desplazar funcionalidad.
- Se necesitan feature toggles y puntos de integración claros entre ambos mundos.

**Fecha:** 2026-08-11

**Evidencia:**
- Repositorio sin código .NET operativo en `modern-net8/`
- Dependencias de hardware y multi-país documentadas en `legacy-restaurant/README.md`

---

## ADR-005

**Título:** `appsettings.json` + User Secrets para configuración

**Estado:** Accepted

**Contexto:**
El Legacy usa archivos INI y credenciales embebidas en código VB6, lo que constituye un riesgo crítico de seguridad y complica despliegues controlados.

**Problema:**
Definir un mecanismo moderno de configuración que separe parámetros no sensibles de secretos.

**Decisión:**
Se utilizará **`appsettings.json`** para configuración no sensible, **`.NET User Secrets`** para desarrollo local y un **secret store del entorno** (variables de entorno / vault corporativo) para ambientes no locales. La configuración Legacy basada en INI será absorbida mediante adaptadores de migración, no replicada como estándar final.

**Alternativas consideradas:**
- Seguir con INI cifrados: descartado por deuda técnica y debilidad de seguridad.
- Variables de entorno exclusivamente: insuficiente para configuración jerárquica compleja del POS.
- Secret manager sin `appsettings`: innecesariamente rígido para parámetros no sensibles.

**Consecuencias:**
- Se elimina el hardcoding de credenciales.
- Deployment y soporte deberán contemplar bootstrap de configuración por local/país.
- Será necesario mapear los parámetros de `INFOREST.INI`, `ALMACEN.INI`, `FACTURACION.INI`, etc.

**Fecha:** 2026-08-11

**Evidencia:**
- Credenciales hardcodeadas documentadas en `legacy-restaurant/README.md`
- Archivos `Inforest.ini`, `Almacen.ini`, `Facturacion.ini`, `DLL3500.ini`, `Tiempo.ini`

---

## ADR-006

**Título:** BCrypt + RBAC en memoria sobre tablas Legacy

**Estado:** Accepted

**Contexto:**
El Legacy usa `TUSUARIO`, `TGRUPOUSUARIO`, `TACCESO`, `TGRUPOACCESO` y cifrado débil en `ClsSeguridad.cls` (XOR/César). Se necesita compatibilidad funcional sin heredar el esquema de seguridad inseguro.

**Problema:**
Definir autenticación y autorización para el Target .NET 8, manteniendo equivalencia con los permisos existentes.

**Decisión:**
La autenticación Target usará **hash BCrypt** para contraseñas. La autorización se implementará como **RBAC en memoria** cargado desde las tablas Legacy de grupos y accesos (`TGRUPOUSUARIO`, `TACCESO`, `TGRUPOACCESO`), encapsulado en Infrastructure/Application.

**Alternativas consideradas:**
- Mantener algoritmo Legacy: descartado por inseguro.
- ASP.NET Core Identity completo: posible, pero sobredimensionado para la primera fase desktop.
- Windows Authentication: insuficiente para escenarios multi-local/multi-país con usuarios propios del sistema.

**Consecuencias:**
- Se deberá diseñar estrategia de transición/rehash de contraseñas.
- El modelo RBAC Legacy puede reutilizarse sin depender del cifrado actual.
- Seguridad queda desacoplada de formularios VB6 y de variables globales.

**Fecha:** 2026-08-11

**Evidencia:**
- `legacy-restaurant/restaurant-vb6/Clases/ClsSeguridad.cls`
- Tablas `TUSUARIO`, `TGRUPOUSUARIO`, `TACCESO`, `TGRUPOACCESO`

---

## ADR-007

**Título:** FastReport .NET para reportes

**Estado:** Accepted

**Contexto:**
El Legacy contiene 206 reportes Crystal Reports `.Dsr` más artefactos `.DCA/.dsx`. Crystal Reports no ofrece una ruta moderna limpia para .NET 8 y agregaría dependencia propietaria heredada.

**Problema:**
Definir el motor de reportes del Target equilibrando compatibilidad .NET 8 y reutilización del SQL existente.

**Decisión:**
Se adopta **FastReport .NET** como motor objetivo de reportes. Los SPs/vistas Legacy de reporting (`spRep_*`, `v*`) se reutilizarán como contratos de datos mientras se migran las plantillas Crystal.

**Alternativas consideradas:**
- Crystal Reports para Visual Studio: descartado por alineación débil con .NET 8 y continuidad del stack legado.
- SSRS: viable para backoffice, menos adecuado para todos los flujos desktop POS.
- Telerik/RDLC: opciones posibles, pero FastReport ofrece mejor balance para desktop .NET 8 y rediseño gradual.

**Consecuencias:**
- La migración de los 206 reportes será manual/semi-manual, no automática.
- Los reportes se podrán validar reutilizando SPs existentes antes de refactorizar datasets.
- La capa Presentation/Reporting debe aislarse del resto del POS.

**Fecha:** 2026-08-11

**Evidencia:**
- `legacy-restaurant/restaurant-vb6/Reportes/` — 206 `.Dsr`
- `5. SP.sql` — familia `spRep_*`

---

## ADR-008

**Título:** Feature flags por país + tenant ID

**Estado:** Accepted

**Contexto:**
El Legacy soporta Perú, Chile, Bolivia, Ecuador, Argentina y España con scripts SQL específicos, configuraciones fiscales diferenciadas y administración centralizada multi-local.

**Problema:**
Definir un modelo Target que maneje variaciones por país/local sin forkear la base funcional.

**Decisión:**
El Target usará **feature flags por país** y un **tenant ID / local ID** explícito para aislar configuración, fiscalidad, capacidades y rollout por local. Las variantes de país no se codificarán como ramas de UI, sino como políticas/configuración externalizada.

**Alternativas consideradas:**
- Deployments totalmente separados por país: eleva costo de mantenimiento.
- Un solo comportamiento rígido: incompatible con normativa fiscal multi-país.
- Solo multi-local sin flags: insuficiente para diferencias regulatorias reales.

**Consecuencias:**
- `TPARAMETRO`, `TCAJA` y scripts por país deben mapearse a configuración externalizada.
- Cada integración fiscal/hardware podrá habilitarse selectivamente.
- La estrategia de pruebas debe incluir matrices país × local × feature set.

**Fecha:** 2026-08-11

**Evidencia:**
- `legacy-restaurant/database-sql-server/opcionales/scriptPeruAlIniciar.sql`
- `legacy-restaurant/database-sql-server/opcionales/scriptChileAlIniciar.sql`
- `legacy-restaurant/database-sql-server/opcionales/scriptBoliviaAlIniciar.sql`
- `legacy-restaurant/database-sql-server/opcionales/scriptEcuadorAlIniciar.sql`
- `legacy-restaurant/database-sql-server/opcionales/scriptArgentinaAlIniciar.sql`
- `legacy-restaurant/database-sql-server/opcionales/scriptEspanaAlIniciar.sql`
- `modPuntoVenta.bas` — `AdministracionCentralizada`

---

*Todos los ADR listados en este documento quedan aceptados como baseline arquitectónico de Fase 2.*

---

## ADR-009

**Título:** Estrategia de interop Win32 para OCX/DLL Legacy (PinPad, fiscal Epson, biometría)

**Estado:** Accepted

**Contexto:**
Los periféricos clave del Legacy (PinPad via `caja_pinpad.dll`, Impresora Fiscal Epson via `IFEpson.ocx`, Biometría SecuGen via `sgfplibx.ocx`) dependen de librerías Win32 de 32 bits sin SDK .NET oficial disponible.

**Problema:**
Definir cómo integrar estas librerías en un proceso .NET 8 (64 bits, WinForms) sin acoplar el dominio a tecnologías Win32 heredadas.

**Decisión:**
- Para DLL nativas con firmas conocidas (`caja_pinpad.dll`): usar **P/Invoke** con declaraciones `[DllImport]` en la capa Infrastructure. La ruta de la DLL es configurable via `appsettings.json`.
- Para OCX de 32 bits sin SDK .NET reemplazable (`IFEpson.ocx`, `sgfplibx.ocx`): implementar **proceso host de 32 bits** separado con comunicación via pipe/IPC, o esperar reemplazo de hardware. Mientras tanto, registrar como `NOT_APPLICABLE` con `Null*Service` stub.
- Todos los periféricos tienen interfaz en Application (`IPinPadService`, `IBiometriaService`, `IImpresoraFiscalService`) con implementación Null por defecto. La implementación real se activa via `appsettings.json` (`Hardware.*.Habilitado = true`).
- La capa Domain no tiene dependencia directa de ninguna librería Win32.

**Alternativas consideradas:**
- COM Interop directo desde proceso 64 bits: no funciona para OCX de 32 bits.
- Recompilar OCX a 64 bits: no factible sin código fuente del proveedor.
- Migrar a SDK moderno equivalente (ej. SecuGen .NET SDK): viable pero requiere evaluación de hardware disponible.

**Consecuencias:**
- PinPad (`caja_pinpad.dll`) puede implementarse con P/Invoke en el corto plazo.
- Fiscal Epson y biometría SecuGen quedan como `NullService` (GAP-004) hasta resolución de hardware.
- El sistema es operativo sin estos periféricos; se habilitan por configuración.

**Fecha:** 2026-08-12

**Evidencia:**
- `legacy-restaurant/restaurant-vb6/Modulos/DLL3500.bas` — `Declare Function fiOpenPort Lib "caja_pinpad.dll"`
- `legacy-restaurant/restaurant-vb6/Modulos/modImpresoraFiscal.bas` — `IFEpson.ocx`
- `legacy-restaurant/restaurant-vb6/Modulos/FpLibX_Const.bas` — `sgfplibx.ocx`

---

## ADR-010

**Título:** Estrategia de Facturación Electrónica por país

**Estado:** Accepted

**Contexto:**
El Legacy integra FE via base MDB auxiliar (`FACTURACION`) con clases VB6 (`clsTrama.cls`, `clsxml.cls`) que generan XML por país. Cada país tiene regulación diferente: SUNAT (Perú), AFIP (Argentina), SII (Chile), SRI (Ecuador), SIAT (Bolivia).

**Problema:**
Definir una arquitectura de FE que soporte múltiples países sin condicionales de país en el código de negocio.

**Decisión:**
- Se define `IFacturacionElectronicaGateway` como contrato polimórfico en Application.
- `FacturacionElectronicaFactory` en Infrastructure resuelve la implementación por `CodigoPais` (leído de `TPARAMETRO`/`appsettings.json`).
- Cada país implementa su propio gateway: `PeruFEGateway`, `ArgentinaFEGateway`, `ChileFEGateway`, `EcuadorFEGateway`, `BoliviaFEGateway`.
- **No se replica la base MDB auxiliar**: la integración es directa con la API del proveedor FE vía HTTP/SOAP.
- Las credenciales por proveedor se configuran en `appsettings.json` bajo `FE.<CodigoPais>` (ADR-005).
- `NullFEGateway` es el default cuando FE no está habilitado (`FE.Habilitado = false`).

**Alternativas consideradas:**
- Replicar el modelo MDB auxiliar Legacy: agrega complejidad innecesaria y no tiene ventajas en .NET 8.
- Un gateway único con switches por país: rompe Open/Closed; dificulta pruebas y extensión.
- Microservicio FE externo: overhead excesivo para fase de migración.

**Consecuencias:**
- Añadir soporte para un nuevo país requiere solo una nueva implementación de `IFacturacionElectronicaGateway`.
- El modelo de documento `DocumentoFE` es el contrato compartido entre todos los gateways.
- Las credenciales y endpoints FE están externalizados y no en código.

**Fecha:** 2026-08-12

**Evidencia:**
- `legacy-restaurant/restaurant-vb6/Modulos/modDeclaracion.bas` — `lFacturacionE`, `clsTramaFE`, `TipoFacturacion`
- `legacy-restaurant/database-sql-server/8. InfoFact.sql`
- `legacy-restaurant/restaurant-vb6/Clases/clsTrama.cls`
- `legacy-restaurant/restaurant-vb6/Clases/clsxml.cls`

---

## ADR-011

**Título:** CashDro — integración via process launcher

**Estado:** Accepted

**Contexto:**
El Legacy integra el cajón automático CashDro lanzando el proceso externo `MotorCashDrow.exe` con `ShellExecute`. No existe DLL de comunicación directa.

**Problema:**
Definir cómo replicar esta integración en .NET 8 manteniendo la compatibilidad con el motor CashDro existente.

**Decisión:**
- `ICashDroService` define los métodos `IniciarMotorAsync(argumentos)` y `AbrirCajonAsync()`.
- `CashDroService` implementa la integración lanzando el proceso via `System.Diagnostics.Process.Start()`.
- La ruta del ejecutable se configura en `appsettings.json` (`Hardware.CashDro.RutaMotor`).
- Si el ejecutable no existe, se lanza `InfrastructureException` con mensaje descriptivo.
- `NullCashDroService` es el default cuando CashDro está deshabilitado.

**Alternativas consideradas:**
- Comunicación IPC directa con MotorCashDrow: requiere conocer el protocolo interno del motor, no documentado.
- SDK moderno de CashDro: evaluar si el fabricante lo provee en una fase posterior.

**Consecuencias:**
- La integración es compatible con el motor CashDro Legacy sin cambios.
- El contrato de argumentos debe documentarse si se extiende la integración.

**Fecha:** 2026-08-12

**Evidencia:**
- `legacy-restaurant/restaurant-vb6/Modulos/modProcedimientoNuevo.bas` — `IniciarMotorCashDrow()`

---

## ADR-012

**Título:** BlueVision — cliente HTTP nativo (no COM)

**Estado:** Accepted

**Contexto:**
El Legacy se comunica con BlueVision TVS via SDK COM (`BlueVision_Core_TVS.dll`), leyendo credenciales de `BLUEVISION.INI`.

**Problema:**
Reemplazar la comunicación COM/VB6 con una integración moderna sin la dependencia de la DLL COM legacy.

**Decisión:**
- `IBlueVisionService` define el contrato de envío de tickets.
- `BlueVisionHttpClient` implementa la integración via `HttpClient` nativo de .NET 8.
- Las credenciales (`Login`, `ClearPassword`, `Url`) se leen de `IOptions<BlueVisionOptions>` configurado en `appsettings.json` bajo `Hardware.BlueVision` (ADR-005). No se usa `BLUEVISION.INI`.
- Los modelos de ticket (`BlueVisionTicket`, `BlueVisionLinea`) son DTOs en Domain, serializados a JSON para la API REST de BlueVision.
- `NullBlueVisionService` es el default cuando BlueVision está deshabilitado.

**Alternativas consideradas:**
- COM Interop con `BlueVision_Core_TVS.dll`: crea dependencia en DLL legacy; incompatible con 64 bits.
- Mantener `BLUEVISION.INI`: inconsistente con la estrategia de configuración centralizada (ADR-005).

**Consecuencias:**
- La integración requiere verificar el contrato HTTP/REST de la API BlueVision TVS.
- Si BlueVision expone WebSockets en lugar de REST, ajustar la implementación.

**Fecha:** 2026-08-12

**Evidencia:**
- `legacy-restaurant/restaurant-vb6/Modulos/modBlueVision.bas` — `TvsClientApi`, `CreateSession`, `SaveTicket`, `SaveTicketLine`
