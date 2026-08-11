# Estándares de Documentación — INFOREST

---

## Principios

1. Documentar solo lo que se puede confirmar con evidencia del repositorio
2. Indicar `UNKNOWN` si algo no es determinable
3. Indicar `PARTIAL` si la información está incompleta
4. Indicar `CONFIRMED` si hay evidencia suficiente
5. No inventar comportamiento, componentes ni reglas de negocio

---

## Indicadores de Confianza

| Indicador | Uso |
|---|---|
| `CONFIRMED` | Verificado directamente en código fuente, scripts o configuración |
| `PARTIAL` | Evidencia parcial, información incompleta |
| `UNKNOWN` | No determinable con la información disponible |

---

## Estados de Migración

| Estado | Uso |
|---|---|
| `NOT_STARTED` | No iniciado |
| `ANALYSIS` | En análisis, sin implementación |
| `IN_PROGRESS` | Implementación en curso |
| `MIGRATED` | Implementado, pendiente de validación |
| `VALIDATING` | En validación contra Legacy |
| `COMPLETED` | Validado y documentado |
| `BLOCKED` | Bloqueado por dependencia o decisión |
| `NOT_APPLICABLE` | No aplica migración |
| `UNKNOWN` | Estado no determinable |

---

## Estructura de Documentos de Módulo

Cada módulo debe tener:

```
docs/modules/<modulo>/
├── README.md              ← Descripción general, Legacy, Target, reglas, dependencias
└── migration-status.md    ← Estado de migración específico del módulo
```

---

## Formato de Reglas de Negocio

```markdown
### BR-XXX
**Nombre:** Nombre descriptivo

**Origen:** Legacy/<archivo>

**Archivo:** legacy-restaurant/restaurant-vb6/<ruta>

**Descripción:** Descripción de la regla

**Condición:** Cuándo se aplica

**Resultado:** Qué produce

**Excepciones:** Casos especiales (UNKNOWN si no determinado)

**Destino .NET:** Componente Target responsable

**Estado:** NOT_STARTED | IN_PROGRESS | MIGRATED | COMPLETED

**Evidencia:** CONFIRMED | PARTIAL | UNKNOWN
```

---

## Formato de ADR

```markdown
## ADR-XXX

**Título:** ...

**Estado:** Proposed | Accepted | Deprecated | Superseded

**Contexto:** Situación que requiere decisión

**Problema:** Pregunta a responder

**Decisión:** Qué se decidió

**Alternativas consideradas:** Opciones evaluadas

**Consecuencias:** Impacto de la decisión

**Fecha:** YYYY-MM-DD o UNKNOWN

**Evidencia:** Evidencia que soporta la decisión
```

---

## Lenguaje

- Documentación en **español** (idioma del proyecto y los usuarios)
- Código en inglés (convención técnica) o español (si el equipo lo decide)
- Términos técnicos en inglés cuando no existe traducción estándar

---

## Archivos README

Cada carpeta relevante del repositorio debe tener un `README.md` que explique:
- Qué hay en la carpeta
- Para qué sirve
- Cómo se relaciona con el resto del proyecto

---

## Actualización

La documentación debe actualizarse:
- Cuando se migra un componente
- Cuando se toma una decisión arquitectónica
- Cuando se descubre nueva información sobre el Legacy
- Cuando cambia el estado de migración

---

## Referencias

- [Instrucciones Copilot](../../.github/copilot-instructions.md)
- [Reglas de migración](../migration/migration-rules.md)
