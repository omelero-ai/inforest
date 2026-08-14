# Selección Tecnológica — INFOREST .NET 8

> Estado: **ACCEPTED** — Decisión tomada en base al análisis técnico del Legacy.
>
> Fecha: 2026-08-11
>
> Relacionado: [ADR-001](../architecture/architecture-decisions.md#adr-001)

---

## Contexto del Análisis

INFOREST es un sistema POS (Point of Sale) empresarial con:
- **401 formularios VB6** — interfaces complejas con lógica mezclada
- **Hardware intensivo:** impresoras térmicas ESC/POS, cajones de dinero (RJ11), PinPad DLL3500, biometría SecuGen OCX, impresora fiscal Epson OCX
- **Operación offline:** no requiere conectividad permanente
- **Multi-caja / multi-local:** cada terminal tiene su propio estado
- **Latencia crítica:** timers de 250ms a 10s para operación en tiempo real
- **Windows-only:** todas las dependencias de hardware son Win32/COM/OCX
- **6 países** con variaciones de impuestos y facturación electrónica

---

## Evaluación de Tecnologías

### Criterios y Pesos

| Criterio | Peso |
|---|---|
| Compatibilidad con Windows | 20% |
| Compatibilidad con POS/Hardware | 15% |
| Impresión (ESC/POS, tickets) | 10% |
| Hardware/periféricos (DLL, OCX, COM) | 10% |
| Facilidad de migración desde VB6 | 15% |
| Mantenibilidad | 10% |
| Performance | 5% |
| Escalabilidad | 5% |
| Seguridad | 5% |
| Curva de aprendizaje | 5% |

### Escala de Puntuación

```
1 = Muy malo
2 = Malo
3 = Aceptable
4 = Bueno
5 = Excelente
```

---

## Comparación de Alternativas

| Criterio | Peso | WinForms .NET 8 | WPF .NET 8 | Avalonia | Blazor Hybrid | Web App |
|---|---|---|---|---|---|---|
| Compatibilidad Windows | 20% | **5** | **5** | **4** | **4** | **2** |
| Compatibilidad POS/Hardware | 15% | **5** | **4** | **3** | **3** | **1** |
| Impresión | 10% | **5** | **4** | **3** | **3** | **2** |
| Hardware/periféricos (DLL,OCX) | 10% | **5** | **4** | **2** | **2** | **1** |
| Migración desde VB6 | 15% | **5** | **3** | **2** | **2** | **1** |
| Mantenibilidad | 10% | **3** | **4** | **4** | **4** | **5** |
| Performance | 5% | **4** | **4** | **4** | **3** | **3** |
| Escalabilidad | 5% | **3** | **3** | **4** | **4** | **5** |
| Seguridad | 5% | **3** | **3** | **3** | **4** | **4** |
| Curva de aprendizaje | 5% | **5** | **3** | **2** | **3** | **3** |
| **TOTAL PONDERADO** | 100% | **4.65** | **3.85** | **3.00** | **3.10** | **2.30** |

---

## Análisis por Tecnología

### WinForms .NET 8 — Puntuación: 4.65/5.00 ✅ RECOMENDADO

**Ventajas:**
- Máxima compatibilidad con hardware POS Win32/COM/OCX
- Migración más directa de 401 formularios VB6
- Sin necesidad de rewrites de hardware drivers
- P/Invoke nativo para DLL Win32 (PinPad DLL3500)
- Interop COM nativo para OCX (IFEpson.ocx, SecuGen)
- Menor curva de aprendizaje para el equipo
- Soporte oficial Microsoft .NET 8 LTS

**Desventajas:**
- UI más limitada que WPF en términos de gráficos/animaciones
- No es multiplataforma (Windows-only)
- DataBinding menos potente que WPF/MVVM

**Compatibilidad de hardware:**
- ✅ Impresoras ESC/POS: System.Drawing + raw printing
- ✅ Cajón de dinero: serial port / raw printing command
- ✅ PinPad DLL3500: P/Invoke directo
- ✅ Biometría SecuGen: COM Interop / wrapper
- ✅ Impresora Fiscal Epson: COM Interop / wrapper

**Riesgos:**
- OCX 32-bit pueden requerir proceso host separado (32-bit bridge)
- Algunos controles OCX pueden no tener soporte .NET directo

---

### WPF .NET 8 — Puntuación: 3.85/5.00

**Ventajas:**
- UI más moderna (MVVM, DataBinding robusto)
- Mejor para pantallas táctiles (POS táctil)
- XAML más mantenible a largo plazo

**Desventajas:**
- Curva de aprendizaje significativa
- Migración de 401 formularios a XAML es costosa
- Hardware POS: mismas capacidades que WinForms pero mayor complejidad
- No es la migración más directa desde VB6

**Conclusión:** Recomendado para una **segunda fase de modernización** después de la migración funcional.

---

### Avalonia — Puntuación: 3.00/5.00

**Ventajas:**
- Multiplataforma (Windows, Linux, macOS)
- UI moderna

**Desventajas:**
- Hardware POS intensivo en Windows no es el caso de uso principal
- Ecosistema más joven, menos documentación para POS
- COM Interop / P/Invoke más complejo
- Riesgo mayor para migración de 401 formularios

**Conclusión:** No recomendado para esta migración.

---

### Blazor Hybrid — Puntuación: 3.10/5.00

**Ventajas:**
- Puede reutilizar UI web en desktop
- Potencial para versión web futura

**Desventajas:**
- Hardware POS requiere igual un host nativo
- Mayor complejidad arquitectónica
- Rendimiento en operaciones de alta frecuencia (timers 250ms) es incierto

**Conclusión:** No recomendado para POS core.

---

### Aplicación Web — Puntuación: 2.30/5.00

**Ventajas:**
- Acceso desde cualquier dispositivo
- Sin instalación en terminales

**Desventajas:**
- Hardware POS imposible de integrar directamente
- Operación offline compleja
- Latencia de red para operaciones críticas (impresión, cajón)
- Requiere reescritura total de hardware layer

**Conclusión:** Puede ser complementario para módulo de Consultas/Reportes (sin hardware), pero NO para POS.

---

## Decisión

### Tecnología Principal: WinForms .NET 8

**Justificación:**
1. INFOREST depende de hardware Windows exclusivo (PinPad, fiscal, biometría, cajón, impresoras ESC/POS)
2. La migración de 401 formularios VB6 es más directa a WinForms
3. Sin pérdida de funcionalidad de hardware
4. Menor riesgo de regresión funcional
5. Soporte LTS .NET 8 hasta noviembre 2026 (.NET 10 LTS disponible para upgrade)
6. El equipo ya conoce el paradigma de formularios

**Nota sobre modernidad:**
La elección de WinForms no implica arquitectura legacy. La arquitectura objetivo usa Clean Architecture con Domain/Application/Infrastructure completamente separados de la UI. WinForms es solo la capa de presentación.

### Posible Evolución (Fase posterior)

```
Fase 2-4: WinForms .NET 8 (POS core)
                     ↓
Fase 5+: WPF .NET 8 (UI modernizada, misma arquitectura)
          +
          Blazor Server (Módulo de Consultas/Reportes — sin hardware)
```

---

## Decisión de Motor de Base de Datos

**SQL Server mantenido** — ver [ADR-002](../architecture/architecture-decisions.md#adr-002)

Justificación:
- 105+ stored procedures con lógica de negocio crítica
- Migración gradual: los SPs se reutilizan durante la transición
- Sin riesgo de pérdida de lógica SQL
- Compatibilidad con scripts por país ya existentes

---

## Decisión de ORM

**Dapper** para operaciones con SPs Legacy + **Entity Framework Core** para nuevas entidades.

| Caso | Herramienta |
|---|---|
| SPs existentes (Legacy) | Dapper |
| Nuevas entidades .NET | EF Core (opcional) |
| Queries complejas | Dapper |
| CRUD simple nuevas entidades | EF Core |

---

## Herramientas Seleccionadas

| Herramienta | Versión | Propósito |
|---|---|---|
| .NET 8 | 8.x LTS | Plataforma |
| C# 12 | 12 | Lenguaje |
| WinForms | .NET 8 | UI |
| Dapper | 2.x | ORM para SPs |
| Microsoft.Data.SqlClient | 5.x | Conexión SQL Server |
| Serilog | 8.x | Logging estructurado |
| BCrypt.Net-Next | 4.x | Hash de passwords (reemplaza XOR+César) |
| xUnit | 2.x | Testing |
| Moq | 4.x | Mocking en tests |
| Microsoft.Extensions.* | 8.x | DI, Configuration, Logging |

---

## Arquitectura de la Solución

```
modern-net8/
├── Inforest.sln
├── src/
│   ├── Inforest.Domain/           ← Entidades, reglas de negocio, interfaces de repositorios
│   ├── Inforest.Application/      ← Casos de uso, commands, queries, interfaces de servicios
│   ├── Inforest.Infrastructure/   ← Repositorios (Dapper), seguridad, logging, hardware
│   └── Inforest.Desktop/          ← WinForms .NET 8, formularios, DI setup
└── tests/
    ├── Inforest.Domain.Tests/
    ├── Inforest.Application.Tests/
    └── Inforest.Infrastructure.Tests/
```

---

## Referencias

- [ADR-001 — Tipo de aplicación](../architecture/architecture-decisions.md#adr-001)
- [ADR-002 — Base de datos](../architecture/architecture-decisions.md#adr-002)
- [ADR-003 — Patrón arquitectónico](../architecture/architecture-decisions.md#adr-003)
- [Arquitectura objetivo](../architecture/target-architecture.md)
- [Estado de migración](../migration/migration-status.md)
