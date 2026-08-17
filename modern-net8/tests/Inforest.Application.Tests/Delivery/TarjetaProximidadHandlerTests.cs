using Inforest.Application.Delivery;
using Inforest.Domain.Entities.Delivery;
using Inforest.Domain.Repositories;

namespace Inforest.Application.Tests.Delivery;

/// <summary>
/// Tests para handlers de TarjetaProximidad.
/// BR-RFID-001..004.
/// </summary>
public sealed class TarjetaProximidadHandlerTests
{
    [Fact]
    public async Task Crear_CodigoDuplicado_RetornaFalla()
    {
        var tarjeta = TarjetaProximidad.Crear("RFID001", "Tarjeta", 0m, "0001", "Free", "USR01");
        var handler = BuildCrearHandler([tarjeta], clienteExiste: true);

        var result = await handler.HandleAsync(new CrearTarjetaProximidadCommand("RFID001", "Duplicada", 0m, "0001", "Free", "USR01"));

        Assert.False(result.EsExitoso);
        Assert.Equal("RFID_CODIGO_DUPLICADO", result.CodigoError);
    }

    [Fact]
    public async Task Crear_ClienteInexistente_RetornaFalla()
    {
        var handler = BuildCrearHandler([], clienteExiste: false);

        var result = await handler.HandleAsync(new CrearTarjetaProximidadCommand("RFID001", "Tarjeta", 0m, "9999", "Free", "USR01"));

        Assert.False(result.EsExitoso);
        Assert.Equal("RFID_CLIENTE_NO_EXISTE", result.CodigoError);
    }

    [Fact]
    public async Task Crear_DatosValidos_RetornaExito()
    {
        var repo = new StubTarjetaRepository();
        var handler = new CrearTarjetaProximidadHandler(repo, new StubClienteDeliveryRepository(clienteExiste: true));

        var result = await handler.HandleAsync(new CrearTarjetaProximidadCommand("RFID001", "Tarjeta VIP", 10m, "0001", "Free", "USR01"));

        Assert.True(result.EsExitoso);
        Assert.Single(repo.Items);
        Assert.Equal("RFID001", repo.Items[0].CodigoRfid);
    }

    [Fact]
    public async Task Actualizar_TarjetaInexistente_RetornaFalla()
    {
        var handler = new ActualizarTarjetaProximidadHandler(new StubTarjetaRepository(), new StubClienteDeliveryRepository(clienteExiste: true));

        var result = await handler.HandleAsync(new ActualizarTarjetaProximidadCommand("RFID404", "Tarjeta", 5m, "0001", "Asignada", "USR01"));

        Assert.False(result.EsExitoso);
        Assert.Equal("RFID_NO_ENCONTRADA", result.CodigoError);
    }

    [Fact]
    public async Task Actualizar_EstadoInvalido_RetornaFallaDominio()
    {
        var repo = new StubTarjetaRepository();
        repo.Items.Add(TarjetaProximidad.Crear("RFID001", "Tarjeta", 5m, "0001", "Free", "USR01"));
        var handler = new ActualizarTarjetaProximidadHandler(repo, new StubClienteDeliveryRepository(clienteExiste: true));

        var result = await handler.HandleAsync(new ActualizarTarjetaProximidadCommand("RFID001", "Tarjeta", 5m, "0001", "Activa", "USR02"));

        Assert.False(result.EsExitoso);
        Assert.Equal("RFID_ESTADO_INVALIDO", result.CodigoError);
    }

    [Fact]
    public async Task ObtenerMovimientos_CodigoVacio_RetornaFalla()
    {
        var handler = new ObtenerMovimientosTarjetaProximidadHandler(new StubTarjetaRepository());

        var result = await handler.HandleAsync(new ObtenerMovimientosTarjetaProximidadQuery(""));

        Assert.False(result.EsExitoso);
        Assert.Equal("RFID_CODIGO_REQUERIDO", result.CodigoError);
    }

    [Fact]
    public async Task ObtenerMovimientos_RetornaUltimosRegistros()
    {
        var repo = new StubTarjetaRepository
        {
            Movimientos =
            [
                new MovimientoTarjetaProximidad(new DateTime(2026, 8, 1), 10m, 0m, 0m, 10m),
                new MovimientoTarjetaProximidad(new DateTime(2026, 8, 2), 0m, 2m, 10m, 8m)
            ]
        };
        var handler = new ObtenerMovimientosTarjetaProximidadHandler(repo);

        var result = await handler.HandleAsync(new ObtenerMovimientosTarjetaProximidadQuery("RFID001"));

        Assert.True(result.EsExitoso);
        Assert.Equal(2, result.Valor!.Count);
        Assert.Equal(8m, result.Valor[1].MontoFinal);
    }

    private static CrearTarjetaProximidadHandler BuildCrearHandler(
        IReadOnlyList<TarjetaProximidad> tarjetas,
        bool clienteExiste)
    {
        var repo = new StubTarjetaRepository();
        repo.Items.AddRange(tarjetas);
        return new CrearTarjetaProximidadHandler(repo, new StubClienteDeliveryRepository(clienteExiste));
    }

    private sealed class StubTarjetaRepository : ITarjetaProximidadRepository
    {
        public List<TarjetaProximidad> Items { get; } = [];
        public List<MovimientoTarjetaProximidad> Movimientos { get; set; } = [];

        public Task<IReadOnlyList<TarjetaProximidad>> ObtenerTodasAsync(CancellationToken ct = default)
            => Task.FromResult<IReadOnlyList<TarjetaProximidad>>(Items);

        public Task<TarjetaProximidad?> ObtenerPorCodigoAsync(string codigoRfid, CancellationToken ct = default)
            => Task.FromResult(Items.FirstOrDefault(x => x.CodigoRfid == codigoRfid));

        public Task<bool> InsertarAsync(TarjetaProximidad tarjeta, CancellationToken ct = default)
        {
            Items.Add(tarjeta);
            return Task.FromResult(true);
        }

        public Task<bool> ActualizarAsync(TarjetaProximidad tarjeta, CancellationToken ct = default)
            => Task.FromResult(true);

        public Task<IReadOnlyList<MovimientoTarjetaProximidad>> ObtenerUltimosMovimientosAsync(string codigoRfid, int cantidad = 10, CancellationToken ct = default)
            => Task.FromResult<IReadOnlyList<MovimientoTarjetaProximidad>>(Movimientos.Take(cantidad).ToList());
    }

    private sealed class StubClienteDeliveryRepository(bool clienteExiste) : IClienteDeliveryRepository
    {
        public Task<string?> ObtenerMaximoCodigoAsync(CancellationToken cancellationToken = default)
            => Task.FromResult<string?>(null);

        public Task<ClienteDelivery?> ObtenerPorCodigoAsync(string codigoDelivery, CancellationToken cancellationToken = default)
            => Task.FromResult(clienteExiste ? CrearCliente(codigoDelivery) : null);

        public Task<ClienteDelivery?> ObtenerPorTelefonoAsync(string telefono, CancellationToken cancellationToken = default)
            => Task.FromResult<ClienteDelivery?>(null);

        public Task<ClienteDelivery?> ObtenerPorIdentidadAsync(string tipoIdentidad, string numeroIdentidad, CancellationToken cancellationToken = default)
            => Task.FromResult<ClienteDelivery?>(null);

        public Task<IEnumerable<ClienteDelivery>> BuscarAsync(string? apellido, string? nombre, string? telefono, string? codigoDistrito, CancellationToken cancellationToken = default)
            => Task.FromResult<IEnumerable<ClienteDelivery>>([]);

        public Task<IEnumerable<ClienteDelivery>> ObtenerFrecuentesAsync(int top = 20, CancellationToken cancellationToken = default)
            => Task.FromResult<IEnumerable<ClienteDelivery>>([]);

        public Task InsertarAsync(ClienteDelivery cliente, CancellationToken cancellationToken = default)
            => Task.CompletedTask;

        public Task ActualizarAsync(ClienteDelivery cliente, CancellationToken cancellationToken = default)
            => Task.CompletedTask;

        public Task ActualizarFotoAsync(string codigoDelivery, byte[] foto, CancellationToken cancellationToken = default)
            => Task.CompletedTask;

        private static ClienteDelivery CrearCliente(string codigo)
            => ClienteDelivery.Crear(codigo, "01", "Cliente", "Demo", "999999999", "Direccion", "01", "01");
    }
}
