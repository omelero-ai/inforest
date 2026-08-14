# Migración — INFOREST

> Documentación del proceso de migración INFOREST VB6 → .NET 8.

---

## Documentos

| Documento | Descripción |
|---|---|
| [migration-strategy.md](migration-strategy.md) | Estrategia y principios de migración |
| [migration-status.md](migration-status.md) | Estado actual de migración por módulo |
| [traceability-matrix.md](traceability-matrix.md) | Matriz de trazabilidad Legacy → .NET 8 |
| [known-gaps.md](known-gaps.md) | Brechas y pendientes identificados |
| [migration-rules.md](migration-rules.md) | Reglas del proceso de migración |
| [business-rules.md](business-rules.md) | Reglas de negocio extraídas del Legacy |

---

## Estado Actual

**Fase actual:** Análisis y Documentación

```
✓ Legacy analizado y documentado (legacy-restaurant/README.md)
✓ Estructura de documentación creada (docs/)
○ Arquitectura .NET 8 definida — PENDIENTE
○ Código .NET 8 — NO INICIADO
○ Pruebas — NO INICIADO
```

---

## Cobertura de Migración

> Los porcentajes no están disponibles porque no existe código .NET 8.

| Componente | Legacy | .NET 8 | Cobertura |
|---|---|---|---|
| Formularios | 401 | 0 | NOT_AVAILABLE |
| Módulos (lógica) | 32 | 0 | NOT_AVAILABLE |
| Clases | 10 | 0 | NOT_AVAILABLE |
| Reglas de negocio | PARTIAL | 0 | NOT_AVAILABLE |
| Stored Procedures | 105+ | 0 | NOT_AVAILABLE |
| Tablas | 126 | 0 | NOT_AVAILABLE |
| Reportes | 206 | 0 | NOT_AVAILABLE |
| Integraciones | 10+ | 0 | NOT_AVAILABLE |

---

## Referencias

- [Análisis Legacy](../../legacy-restaurant/README.md)
- [Arquitectura Target](../architecture/target-architecture.md)
- [Instrucciones Copilot](../../.github/copilot-instructions.md)
