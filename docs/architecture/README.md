# Arquitectura — INFOREST

Documentación de arquitectura del sistema INFOREST, tanto del sistema Legacy VB6 como de la arquitectura objetivo .NET 8.

---

## Documentos

| Documento | Descripción |
|---|---|
| [system-context.md](system-context.md) | Contexto del sistema — diagrama C4 nivel 1 |
| [legacy-architecture.md](legacy-architecture.md) | Arquitectura del sistema Legacy VB6 |
| [target-architecture.md](target-architecture.md) | Arquitectura objetivo .NET 8 |
| [architecture-overview.md](architecture-overview.md) | Vista general comparativa Legacy vs Target |
| [architecture-decisions.md](architecture-decisions.md) | ADRs — Decisiones arquitectónicas documentadas |

---

## Resumen

| Aspecto | Legacy | Target |
|---|---|---|
| Tecnología | Visual Basic 6 | .NET 8 (C#) |
| Estilo | Monolítico cliente-servidor | Por definir |
| UI | Windows Forms (VB6) | Por definir |
| Persistencia | SQL Server (ADO) | Por definir |
| Reportes | Crystal Reports | Por definir |
| Configuración | INI files | Por definir |
| Integraciones | COM/ActiveX/DLL | Por definir |

---

## Referencias

- [Análisis técnico Legacy](../../legacy-restaurant/README.md)
- [Decisiones arquitectónicas](architecture-decisions.md)
