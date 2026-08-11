# Decisiones Arquitectónicas — INFOREST (ADR)

> Registro de decisiones arquitectónicas importantes del proyecto de migración INFOREST VB6 → .NET 8.

---

## Índice de ADRs

| ADR | Título | Estado |
|---|---|---|
| [ADR-001](#adr-001) | Selección de tipo de aplicación Target | Proposed |
| [ADR-002](#adr-002) | Estrategia de base de datos Target | Proposed |
| [ADR-003](#adr-003) | Patrón arquitectónico para la nueva implementación | Proposed |
| [ADR-004](#adr-004) | Estrategia de migración: Big Bang vs Strangler Fig | Proposed |
| [ADR-005](#adr-005) | Gestión de credenciales y configuración | Proposed |
| [ADR-006](#adr-006) | Seguridad: autenticación y autorización | Proposed |
| [ADR-007](#adr-007) | Estrategia de reportes | Proposed |
| [ADR-008](#adr-008) | Manejo de multi-país y multi-local | Proposed |

---

## ADR-001

**Título:** Selección de tipo de aplicación Target

**Estado:** Proposed

**Contexto:**
El sistema Legacy es una aplicación desktop Windows (VB6 Forms) instalada en cada terminal POS. Existen dependencias de hardware (impresoras térmicas, cajones, PinPad, biometría) que son dependientes de Windows.

**Problema:**
¿El sistema Target debe ser una aplicación Web, Desktop nativo (.NET WinForms/WPF/MAUI) o híbrido?

**Decisión:**
UNKNOWN — pendiente de decisión del equipo y stakeholders.

**Alternativas consideradas:**
- Aplicación Web (Blazor / React + .NET API): mayor accesibilidad, pero requiere abstracción de hardware
- Desktop .NET 8 (WinForms/WPF): migración más directa, sigue siendo Windows-only
- Desktop MAUI: multiplataforma, pero limitaciones de hardware
- Híbrido: API .NET 8 + cliente Web para gestión + cliente Desktop para POS

**Consecuencias:**
La elección impacta directamente en la estrategia de integración de hardware POS.

**Fecha:** UNKNOWN

**Evidencia:**
- `legacy-restaurant/restaurant-vb6/` — hardware dependencies
- `modern-net8/README.md` — placeholder sin definición

---

## ADR-002

**Título:** Estrategia de base de datos Target

**Estado:** Proposed

**Contexto:**
El Legacy usa SQL Server con 126 tablas, 105 vistas y 105+ stored procedures. Los nombres de tablas siguen convenciones específicas (prefijos M, D, T, A).

**Problema:**
¿Se migra a SQL Server (misma plataforma) o a otro motor (PostgreSQL, etc.)? ¿Se mantienen los nombres de tablas Legacy o se renombran?

**Decisión:**
UNKNOWN — pendiente de decisión.

**Alternativas consideradas:**
- Mantener SQL Server: mínimo riesgo, reutiliza SPs existentes
- Migrar a PostgreSQL: reduce costo de licencias, mayor compatibilidad cloud
- Renombrar tablas siguiendo convenciones modernas: más legible pero requiere mapeo

**Consecuencias:**
Si se renombran tablas, se requiere mapeo completo en `docs/database/mapping.md`.
Si se mantiene SQL Server, algunos SPs pueden reutilizarse.

**Fecha:** UNKNOWN

**Evidencia:**
- `legacy-restaurant/database-sql-server/1. Estructura.sql` — 126 tablas

---

## ADR-003

**Título:** Patrón arquitectónico para la nueva implementación

**Estado:** Proposed

**Contexto:**
El Legacy es monolítico sin separación de capas. El target debe tener separación explícita para mantenibilidad.

**Problema:**
¿Qué patrón arquitectónico usar para la implementación .NET 8?

**Decisión:**
UNKNOWN — pendiente de decisión del equipo.

**Alternativas consideradas:**
- Clean Architecture (Uncle Bob): Domain, Application, Infrastructure, Presentation
- N-Layer tradicional: UI, BLL, DAL, DB
- CQRS + Event Sourcing: mayor complejidad, más trazabilidad
- Monolito modular: por módulo funcional (POS, Caja, Delivery, Admin, Cocina)
- Microservicios: mayor complejidad operacional

**Consecuencias:**
Impacta en la organización del código en `modern-net8/` y en la estrategia de pruebas.

**Fecha:** UNKNOWN

---

## ADR-004

**Título:** Estrategia de migración: Big Bang vs Strangler Fig

**Estado:** Proposed

**Contexto:**
INFOREST es un sistema en producción. No existe código .NET 8. La migración debe hacerse sin interrumpir operaciones.

**Problema:**
¿Se migra todo a la vez (Big Bang) o se migra gradualmente módulo a módulo (Strangler Fig)?

**Decisión:**
UNKNOWN — pendiente de decisión.

**Alternativas consideradas:**
- **Big Bang:** Reescribir todo y desplegar en una sola versión. Alto riesgo.
- **Strangler Fig Pattern:** Migrar módulo por módulo, coexistiendo Legacy y .NET. Bajo riesgo, mayor duración.
- **Parallel Run:** Ejecutar Legacy y .NET en paralelo para validar comportamiento. Costo alto pero validación completa.

**Consecuencias:**
Strangler Fig requiere definir la interfaz entre Legacy y .NET durante la transición.

**Fecha:** UNKNOWN

**Evidencia:**
- Sistema en uso activo (evidenciado por complejidad del código y múltiples países soportados)

---

## ADR-005

**Título:** Gestión de credenciales y configuración

**Estado:** Proposed

**Contexto:**
El Legacy tiene credenciales SQL Server hardcodeadas en `modPuntoVenta.bas` Sub Main(). Este es un riesgo crítico de seguridad.

**Problema:**
¿Cómo gestionar credenciales y configuración en el sistema Target?

**Decisión:**
Las credenciales NO deben estar en el código fuente. Se propone usar:
- `appsettings.json` para configuración no sensible
- Variables de entorno o secrets manager para credenciales

**Alternativas consideradas:**
- Environment variables: simple, disponible en todos los entornos
- Azure Key Vault / AWS Secrets Manager: más seguro para producción cloud
- .NET User Secrets (desarrollo): solo para ambiente de desarrollo
- INI files cifrados: migración mínima del enfoque actual

**Consecuencias:**
- El sistema deja de ser portable por copia de ejecutable + INI
- Requiere proceso de configuración en despliegue

**Estado:** Proposed — aceptable para proyecto

**Fecha:** UNKNOWN

**Evidencia:**
- Credenciales hardcodeadas detectadas en análisis Legacy (ver `legacy-restaurant/README.md` §7)

---

## ADR-006

**Título:** Seguridad: autenticación y autorización

**Estado:** Proposed

**Contexto:**
El Legacy usa `TUSUARIO`, `TGRUPOUSUARIO`, `TACCESO`, `TGRUPOACCESO` para control de acceso. El cifrado actual (`ClsSeguridad`) usa XOR+César que es criptográficamente débil.

**Problema:**
¿Cómo implementar autenticación y autorización en el sistema Target?

**Decisión:**
UNKNOWN — pendiente de decisión.

**Alternativas consideradas:**
- ASP.NET Core Identity con RBAC
- JWT ******
- Windows Authentication (si se mantiene desktop)
- OAuth2 / OpenID Connect

**Consecuencias:**
Los modelos de usuario y grupos (`TUSUARIO`, `TGRUPOUSUARIO`) deben migrarse.
Los passwords deben re-hashearse con BCrypt/Argon2.

**Fecha:** UNKNOWN

**Evidencia:**
- `Clases/ClsSeguridad.cls` — cifrado débil
- Tablas `TUSUARIO`, `TGRUPOUSUARIO`, `TACCESO` en BD Legacy

---

## ADR-007

**Título:** Estrategia de reportes

**Estado:** Proposed

**Contexto:**
El Legacy tiene 206 reportes Crystal Reports (.dsr/.dca). Crystal Reports requiere licencia y no es compatible con .NET 8 directamente.

**Problema:**
¿Cómo migrar 206 reportes Crystal Reports a .NET 8?

**Decisión:**
UNKNOWN — pendiente de decisión.

**Alternativas consideradas:**
- SAP Crystal Reports for Visual Studio (compatibilidad limitada con .NET 8)
- FastReport .NET (reescritura manual pero compatible)
- SSRS (SQL Server Reporting Services)
- Telerik Reporting
- rdlc (Report Definition Language Client-side)

**Consecuencias:**
206 reportes requieren migración manual o herramienta de conversión.
Los SPs de reporte (`spRep_*`) pueden reutilizarse independientemente del motor de reportes.

**Fecha:** UNKNOWN

**Evidencia:**
- `legacy-restaurant/restaurant-vb6/Reportes/` — 206 reportes .dsr

---

## ADR-008

**Título:** Manejo de multi-país y multi-local

**Estado:** Proposed

**Contexto:**
El Legacy soporta múltiples países con scripts SQL específicos y múltiples locales con administración centralizada. La configuración por país incluye impuestos, facturación electrónica y normativa fiscal.

**Problema:**
¿Cómo implementar multi-país y multi-local en .NET 8?

**Decisión:**
UNKNOWN — pendiente de decisión.

**Alternativas consideradas:**
- Multi-tenancy por país/local: un solo deployment con configuración por tenant
- Deployments independientes por país/local
- Feature flags por país
- Configuración externalizada con perfiles por país

**Consecuencias:**
La base de datos debe soportar `local_id` para mantener unicidad en arquitectura multi-local.

**Fecha:** UNKNOWN

**Evidencia:**
- `legacy-restaurant/database-sql-server/opcionales/` — scripts por país (6 países)
- Flag `CENTRALIZADA=ON` en `INFOREST.INI`

---

*Los ADRs en estado `Proposed` están pendientes de aceptación formal por el equipo de arquitectura.*
