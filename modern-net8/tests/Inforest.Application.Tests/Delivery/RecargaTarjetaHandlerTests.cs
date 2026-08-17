using Inforest.Application.Delivery;
using Inforest.Domain.Entities.Delivery;
using Inforest.Domain.Repositories;

namespace Inforest.Application.Tests.Delivery;

public sealed class RecargaTarjetaHandlerTests
{
    [Fact]
    public async Task RegistrarRecarga_FallaSiTarjetaNoExiste()
    {
        var repo = new StubRecargaRepository();
        var tarjetaRepo = new StubTarjetaRepository();
        var handler = new RegistrarRecargaTarjetaHandler(tarjetaRepo, repo);

        var result = await handler.HandleAsync(new RegistrarRecargaTarjetaCommand(
            CodigoRfid: "RF-404",
            MontoRecarga: 10m,
            Usuario: "USR01"));

        Assert.False(result.EsExitoso);
        Assert.Equal("RFID_RECARGA_TARJETA_NO_EXISTE", result.CodigoError);
    }

    [Fact]
    public async Task RegistrarRecarga_FallaSiTarjetaBloqueada()
    {
        var repo = new StubRecargaRepository();
        var tarjetaRepo = new StubTarjetaRepository();
        tarjetaRepo.Items.Add(TarjetaProximidad.Crear("RF-1", "Tarjeta", 15m, "0001", "Bloqueado", "USR01"));
        var handler = new RegistrarRecargaTarjetaHandler(tarjetaRepo, repo);

        var result = await handler.HandleAsync(new RegistrarRecargaTarjetaCommand(
            CodigoRfid: "RF-1",
            MontoRecarga: 5m,
            Usuario: "USR02"));

        Assert.False(result.EsExitoso);
        Assert.Equal("RFID_RECARGA_TARJETA_BLOQUEADA", result.CodigoError);
    }

    [Fact]
    public async Task RegistrarRecarga_ActualizaSaldoYRegistraMovimiento()
    {
        var repo = new StubRecargaRepository();
        var tarjetaRepo = new StubTarjetaRepository();
        tarjetaRepo.Items.Add(TarjetaProximidad.Crear("RF-1", "Tarjeta", 20m, "0001", "Asignada", "USR01"));
        var handler = new RegistrarRecargaTarjetaHandler(tarjetaRepo, repo);

        var result = await handler.HandleAsync(new RegistrarRecargaTarjetaCommand(
            CodigoRfid: "RF-1",
            MontoRecarga: 12.5m,
            Usuario: "USR02",
            DocumentoReferencia: "DOC-1"));

        Assert.True(result.EsExitoso);
        Assert.Equal(20m, result.Valor!.MontoAnterior);
        Assert.Equal(32.5m, result.Valor!.MontoFinal);
        Assert.Equal("RF-1", repo.LastCodigoRfid);
    }

    private sealed class StubTarjetaRepository : ITarjetaProximidadRepository
    {
        public List<TarjetaProximidad> Items { get; } = [];

        public Task<IReadOnlyList<TarjetaProximidad>> ObtenerTodasAsync(CancellationToken ct = default)
            => Task.FromResult<IReadOnlyList<TarjetaProximidad>>(Items);

        public Task<TarjetaProximidad?> ObtenerPorCodigoAsync(string codigoRfid, CancellationToken ct = default)
            => Task.FromResult(Items.FirstOrDefault(x => x.CodigoRfid == codigoRfid));

        public Task<bool> InsertarAsync(TarjetaProximidad tarjeta, CancellationToken ct = default)
            => Task.FromResult(true);

        public Task<bool> ActualizarAsync(TarjetaProximidad tarjeta, CancellationToken ct = default)
            => Task.FromResult(true);

        public Task<IReadOnlyList<MovimientoTarjetaProximidad>> ObtenerUltimosMovimientosAsync(string codigoRfid, int cantidad = 10, CancellationToken ct = default)
            => Task.FromResult<IReadOnlyList<MovimientoTarjetaProximidad>>([]);
    }

    private sealed class StubRecargaRepository : IRecargaTarjetaRepository
    {
        public string? LastCodigoRfid { get; private set; }

        public Task<IReadOnlyList<RecargaTarjetaMovimiento>> ObtenerMovimientosAsync(DateTime desde, DateTime hasta, string tipoMovimiento, string? codigoRfid = null, CancellationToken ct = default)
            => Task.FromResult<IReadOnlyList<RecargaTarjetaMovimiento>>([]);

        public Task<RecargaTarjetaResultado> RegistrarRecargaAsync(string codigoRfid, decimal montoRecarga, string documentoReferencia, CancellationToken ct = default)
        {
            LastCodigoRfid = codigoRfid;
            return Task.FromResult(new RecargaTarjetaResultado(codigoRfid, documentoReferencia, 20m, montoRecarga, 20m + montoRecarga));
        }
    }
}
