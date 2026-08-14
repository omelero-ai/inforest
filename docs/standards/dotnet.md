# Estándares de Código .NET 8 — INFOREST

> Estado: **Accepted**
>
> Aplica a: Toda la implementación en `modern-net8/`
>
> Última actualización: 2026-08-11

---

## 1. Plataforma y Lenguaje

- **Plataforma:** .NET 8 LTS
- **Lenguaje:** C# 12
- **Nullable Reference Types:** habilitado (`<Nullable>enable</Nullable>`) — obligatorio
- **Implicit Usings:** habilitado
- **File-scoped namespaces:** preferidos
- **Target Framework:** `net8.0` (Domain, Application, Infrastructure) / `net8.0-windows` (Desktop)

---

## 2. Estructura de la Solución

```
Inforest.Domain        ← Sin dependencias externas. Solo .NET BCL.
Inforest.Application   ← Depende de Domain. Solo abstracciones.
Inforest.Infrastructure ← Depende de Domain y Application. SQL, Hardware, Logging.
Inforest.Desktop       ← Depende de Application e Infrastructure. Solo UI.
```

### Regla de dependencias

```
Desktop → Application → Domain ← Infrastructure
```

**Infrastructure NO puede ser referenciado por Domain ni Application.**
**Domain NO puede referenciar Application ni Infrastructure.**

---

## 3. Naming Conventions

### Clases, Interfaces, Enums, Records

```csharp
// Clases: PascalCase
public class Pedido { }
public class DetallePedido { }

// Interfaces: prefijo I
public interface IPedidoRepository { }
public interface IAuthService { }

// Enums: PascalCase, valores PascalCase
public enum EstadoPedido { Abierto, Cerrado, Anulado }
public enum CanalVenta { Local = 1, Delivery = 2, Llevar = 3 }

// Records: PascalCase
public record AuthResult(bool Exitoso, string? MensajeError);
```

### Métodos

```csharp
// PascalCase, verbos descriptivos
public Pedido Crear(...) { }
public void AgregarDetalle(...) { }
public async Task<Pedido?> ObtenerPorCodigoAsync(string codigo, CancellationToken ct = default) { }
```

### Propiedades y Campos

```csharp
// Propiedades: PascalCase
public string CodigoPedido { get; private set; }
public CanalVenta CanalVenta { get; private set; }

// Campos privados: _camelCase
private readonly List<DetallePedido> _detalles = [];
private readonly ILogger<PedidoService> _logger;

// Constantes: PascalCase o UPPER_SNAKE_CASE
private const int MaximoCanalVenta = 5;
```

### Namespaces

```csharp
// Formato: Inforest.<Capa>.<Submodulo>
namespace Inforest.Domain.Entities.Ventas;
namespace Inforest.Application.Commands.Pedidos;
namespace Inforest.Infrastructure.Repositories;
namespace Inforest.Desktop.Forms.PuntoVenta;
```

---

## 4. Async / Await

**Regla:** Usar async/await únicamente cuando la operación sea I/O-bound:
- Acceso a base de datos
- Lectura/escritura de archivos
- Llamadas a red (HTTP, TCP)
- Lectura de hardware que bloquea

**No usar async para:**
- Operaciones puramente en memoria
- Cálculos simples
- Lógica de dominio

```csharp
// ✅ Correcto — I/O de base de datos
public async Task<Pedido?> ObtenerPorCodigoAsync(string codigo, CancellationToken ct = default)
{
    return await _connection.QueryFirstOrDefaultAsync<Pedido>(sql, new { codigo }, commandTimeout: 30);
}

// ✅ Correcto — dominio síncrono
public decimal CalcularTotal(IEnumerable<DetallePedido> detalles)
{
    return detalles.Sum(d => d.Venta);
}

// ❌ Incorrecto — async innecesario
public async Task<decimal> CalcularTotalAsync(...)
{
    return await Task.FromResult(detalles.Sum(d => d.Venta));
}
```

---

## 5. Excepciones

```csharp
// Domain: usar DomainException con código
throw new DomainException("El turno ya está cerrado.", "TURNO_YA_CERRADO");

// Infrastructure: capturar y loguear, re-lanzar como dominio o application exception
catch (SqlException ex) when (ex.Number == 1205) // Deadlock
{
    _logger.LogError(ex, "Deadlock al insertar pedido {CodigoPedido}", codigo);
    throw new InfrastructureException("Error de concurrencia al guardar el pedido.", ex);
}

// NUNCA usar catch genérico sin log
// ❌
catch (Exception) { }

// ✅
catch (Exception ex)
{
    _logger.LogError(ex, "Error inesperado en {Operacion}", nameof(InsertarPedidoAsync));
    throw;
}
```

---

## 6. Logging

Usar `ILogger<T>` con mensajes estructurados (Serilog).

```csharp
// ✅ Correcto — parámetros nombrados
_logger.LogInformation("Pedido {CodigoPedido} creado para caja {CodigoCaja}", codigo, caja);
_logger.LogWarning("Producto {CodigoProducto} sin stock en almacén", codigoProducto);
_logger.LogError(ex, "Error al cerrar turno {CodigoTurno}", codigoTurno);

// ❌ Incorrecto — string interpolation en logging (pérdida de structured logging)
_logger.LogInformation($"Pedido {codigo} creado");
```

---

## 7. Dependency Injection

```csharp
// Registrar en DependencyInjection.cs de cada capa
services.AddScoped<IPedidoRepository, PedidoRepository>();
services.AddScoped<IAuthService, AuthService>();
services.AddSingleton<IParametroService, ParametroService>(); // Cache de parámetros

// En formularios WinForms: resolver desde ServiceProvider
public partial class FrmPedido : Form
{
    private readonly IPedidoRepository _pedidoRepo;

    public FrmPedido(IPedidoRepository pedidoRepo)
    {
        InitializeComponent();
        _pedidoRepo = pedidoRepo;
    }
}
```

---

## 8. Repositorios con Dapper

```csharp
// Para SPs Legacy: usar Dapper
public async Task InsertarAsync(Pedido pedido, CancellationToken ct = default)
{
    using var connection = _connectionFactory.CreateConnection();
    var parameters = new DynamicParameters();
    parameters.Add("@tCodigoPedido", pedido.CodigoPedido);
    // ... otros parámetros
    
    await connection.ExecuteAsync("spIns_MPEDIDO", parameters,
        commandType: CommandType.StoredProcedure,
        commandTimeout: 30);
}

// Para queries simples: SQL inline con Dapper
public async Task<IEnumerable<ProductoDto>> ObtenerActivosAsync()
{
    using var connection = _connectionFactory.CreateConnection();
    const string sql = @"
        SELECT tCodigoProducto, tDescripcion, nPrecioVenta
        FROM TPRODUCTO
        WHERE lActivo = 1
        ORDER BY tDescripcion";
    
    return await connection.QueryAsync<ProductoDto>(sql);
}
```

---

## 9. DTOs

```csharp
// DTOs son records inmutables
public record ProductoDto(
    string CodigoProducto,
    string Descripcion,
    decimal PrecioVenta,
    decimal PrecioDelivery);

// Request/Response para comandos
public record CrearPedidoCommand(
    string CodigoCaja,
    string CodigoTurno,
    string CodigoUsuario,
    CanalVenta CanalVenta,
    int NumeroPersonas);

public record CrearPedidoResult(string CodigoPedido, bool Exitoso, string? Error);
```

---

## 10. Documentación XML

Obligatorio en:
- Toda interfaz pública (`IRepository`, `IService`)
- Toda entidad del dominio
- Toda regla de negocio implementada

```csharp
/// <summary>
/// Calcula el precio neto sin impuestos.
/// Legacy: BR-001 — frmVenta.frm, nPrecioNeto = nPrecioVenta / (1 + %IGV)
/// </summary>
/// <param name="precioVenta">Precio de venta con impuestos incluidos.</param>
/// <param name="porcentajeImpuesto1">Porcentaje del primer impuesto (ej: 18 para IGV Perú).</param>
public decimal CalcularPrecioNeto(decimal precioVenta, decimal porcentajeImpuesto1, ...)
```

---

## 11. Testing

```csharp
// Convención de nombre: Metodo_Condicion_ResultadoEsperado
[Fact]
public void Crear_SinCodigoPedido_LanzaExcepcion() { }

[Fact]
public void CalcularPrecioNeto_Con18PorcentoIGV_RetornaNetoCorreecto() { }

// Arrange / Act / Assert
[Fact]
public void Crear_ConDatosValidos_CreaConEstadoAbierto()
{
    // Arrange
    var codigo = "PED001";
    var turno = "TRN001";

    // Act
    var pedido = Pedido.Crear(codigo, null, turno, "USR001", CanalVenta.Local);

    // Assert
    Assert.Equal(EstadoPedido.Abierto, pedido.Estado);
}
```

---

## 12. Reglas para Migración de VB6

- Cada archivo .cs que migra un formulario/módulo VB6 **debe** tener comentario con la fuente Legacy
- Las reglas de negocio **deben** referenciar el `BR-XXX` correspondiente
- No inventar comportamiento — si no se puede determinar del Legacy, marcar `// TODO: verificar en Legacy`
- Los SPs se llaman por nombre exacto del Legacy (`spIns_MPEDIDO`, `spUpd_MPEDIDO`, etc.)

```csharp
/// <summary>
/// Repositorio de pedidos.
/// Legacy: spIns_MPEDIDO, spUpd_MPEDIDO, MPEDIDO, DPEDIDO
/// BR-SQL-001, BR-SQL-002
/// </summary>
public class PedidoRepository : IPedidoRepository
```

---

## 13. Configuración y Secretos

```csharp
// ✅ Correcto — desde IConfiguration
var connectionString = configuration.GetConnectionString("Inforest")
    ?? throw new InvalidOperationException("ConnectionString 'Inforest' no configurada.");

// ❌ Incorrecto — hardcodeado (problema SEC-001 del Legacy)
var connectionString = "Server=localhost;Database=INFOREST;User=sa;******;";
```

**Gestión de secretos:**
- Desarrollo: `dotnet user-secrets`
- Producción: Variables de entorno con prefijo `INFOREST_`

---

## Referencias

- [Estándares de documentación](documentation-standards.md)
- [Naming conventions](naming-conventions.md)
- [Arquitectura objetivo](../architecture/target-architecture.md)
- [Reglas de negocio](../migration/business-rules.md)
