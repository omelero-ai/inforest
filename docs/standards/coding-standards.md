# Estándares de Codificación .NET 8 — INFOREST

> Status: Proposed — pendiente de validación por el equipo.

---

## Lenguaje y Plataforma

- **Lenguaje:** C# 12+ con .NET 8
- **Nullable Reference Types:** habilitado (`<Nullable>enable</Nullable>`)
- **Implicit Usings:** habilitado
- **Target Framework:** net8.0

---

## Organización del Proyecto

```
modern-net8/
├── src/
│   ├── Inforest.Domain/           ← Entidades, value objects, reglas de negocio
│   ├── Inforest.Application/      ← Casos de uso, commands, queries
│   ├── Inforest.Infrastructure/   ← Repositorios, integraciones, persistencia
│   └── Inforest.Presentation/     ← UI o API controllers
└── tests/
    ├── Inforest.Domain.Tests/
    ├── Inforest.Application.Tests/
    └── Inforest.Integration.Tests/
```

> Esta estructura es propuesta. Sujeta a cambio según [ADR-003](../architecture/architecture-decisions.md#adr-003).

---

## Convenciones de Código C#

### Nombrado

- **Clases:** PascalCase — `OrderService`, `ProductRepository`
- **Interfaces:** IPascalCase — `IOrderService`, `IProductRepository`
- **Métodos:** PascalCase — `CreateOrder`, `GetProductById`
- **Propiedades:** PascalCase — `OrderId`, `TotalAmount`
- **Variables locales:** camelCase — `orderId`, `totalAmount`
- **Campos privados:** `_camelCase` — `_orderRepository`
- **Constantes:** UPPER_CASE — `MAX_ORDER_ITEMS`
- **Enums:** PascalCase — `OrderStatus`, `PaymentMethod`

### Organización de archivos

- Un tipo por archivo
- Nombre del archivo = nombre de la clase
- Namespace = estructura de carpetas

### Async/Await

- Todos los métodos de I/O deben ser async
- Sufijo `Async` en métodos async — `GetOrderAsync`, `CreateOrderAsync`
- Usar `CancellationToken` en métodos async públicos

### Inyección de Dependencias

- Usar constructor injection
- Registrar servicios en Program.cs o módulos de extensión
- Prohibido el uso de ServiceLocator

---

## Seguridad Obligatoria

> Estos estándares corrigen problemas críticos del Legacy.

| Estándar | Reemplaza | Razón |
|---|---|---|
| `BCrypt` para passwords | XOR+César (`ClsSeguridad`) | Cifrado seguro |
| `AES-256` para datos sensibles | XOR+César | Cifrado seguro |
| Parametrized queries / EF Core | SQL inline | SQL injection |
| `appsettings.json` + environment variables | Credenciales hardcodeadas | Seguridad |
| RBAC con roles | Control manual | Estándar moderno |

---

## Pruebas

- **Framework:** xUnit (recomendado)
- **Mocking:** Moq o NSubstitute
- **Cobertura mínima en dominio:** >80%
- Pruebas unitarias en `*.Domain.Tests`
- Pruebas de integración en `*.Integration.Tests`
- Cada regla de negocio (`BR-XXX`) debe tener al menos un test

---

## Logging

- **Framework:** Serilog (recomendado)
- Logging estructurado con contexto
- Niveles: Debug, Information, Warning, Error, Critical
- No loguear datos sensibles (passwords, tokens, credenciales)

---

## Manejo de Errores

- Excepciones de dominio específicas (no `Exception` genérica)
- Result pattern para operaciones que pueden fallar
- No suprimir excepciones sin logging

---

## Referencias

- [Naming Conventions](naming-conventions.md)
- [ADR-003](../architecture/architecture-decisions.md#adr-003)
- Microsoft C# Coding Conventions
