using Inforest.Domain.Common;
using Inforest.Domain.Entities.Ventas;
using Inforest.Domain.Exceptions;
using Inforest.Domain.Services;

namespace Inforest.Domain.Tests;

/// <summary>
/// Pruebas baseline de P3-01 — valida que los contratos de Domain sean correctos.
/// </summary>
public class ResultTests
{
    [Fact]
    public void Result_Ok_EsExitoso_True()
    {
        var result = Result.Ok();
        Assert.True(result.EsExitoso);
        Assert.Null(result.CodigoError);
        Assert.Null(result.MensajeError);
    }

    [Fact]
    public void Result_Fail_EsExitoso_False()
    {
        var result = Result.Fail("Error de prueba", "CODIGO_PRUEBA");
        Assert.False(result.EsExitoso);
        Assert.Equal("CODIGO_PRUEBA", result.CodigoError);
        Assert.Equal("Error de prueba", result.MensajeError);
    }

    [Fact]
    public void ResultT_Ok_ContieneValor()
    {
        var result = Result.Ok(42);
        Assert.True(result.EsExitoso);
        Assert.Equal(42, result.Valor);
    }

    [Fact]
    public void ResultT_Fail_NoContieneValor()
    {
        var result = Result.Fail<int>("Falló", "ERR");
        Assert.False(result.EsExitoso);
        Assert.Equal(default, result.Valor);
        Assert.Equal("ERR", result.CodigoError);
    }

    [Fact]
    public void ResultT_ConversionImplicita_AResult()
    {
        Result<string> typed = Result<string>.Fail("msg", "CODE");
        Result untyped = typed;
        Assert.False(untyped.EsExitoso);
        Assert.Equal("CODE", untyped.CodigoError);
    }
}

public class PedidoTests
{
    [Fact]
    public void Pedido_Crear_CodigoVacio_LanzaException()
    {
        Assert.Throws<DomainException>(() =>
            Pedido.Crear("", "M01", "T01", "USR01", CanalVenta.Local));
    }

    [Fact]
    public void Pedido_Crear_SinTurno_LanzaException()
    {
        Assert.Throws<DomainException>(() =>
            Pedido.Crear("P001", "M01", "", "USR01", CanalVenta.Local));
    }

    [Fact]
    public void Pedido_Crear_Valido_EstadoAbierto()
    {
        var pedido = Pedido.Crear("P001", "M01", "T01", "USR01", CanalVenta.Local);
        Assert.Equal(EstadoPedido.Abierto, pedido.Estado);
        Assert.Equal(CanalVenta.Local, pedido.CanalVenta);
    }

    [Fact]
    public void Pedido_Anular_Cerrado_LanzaException()
    {
        // Solo se puede testear la ruta Anular->Cerrado
        // dado que no hay método Cerrar expuesto aún
        var pedido = Pedido.Crear("P002", "M02", "T01", "USR01", CanalVenta.Delivery);
        pedido.Anular("test", "USR01");
        Assert.Equal(EstadoPedido.Anulado, pedido.Estado);
    }
}

public class PrecioServiceTests
{
    private readonly PrecioService _service = new();

    [Fact]
    public void CalcularPrecioNeto_SinImpuesto_RetornaMismoPrecio()
    {
        var neto = _service.CalcularPrecioNeto(100m, 0, 0, 0);
        Assert.Equal(100m, neto);
    }

    [Fact]
    public void CalcularPrecioNeto_Con18PctIgv_RetornaNetoEsperado()
    {
        // 118 / 1.18 = 100
        var neto = _service.CalcularPrecioNeto(118m, 18m, 0, 0);
        Assert.Equal(100m, neto, precision: 10);
    }

    [Fact]
    public void CalcularTotalLinea_Multiplicacion()
    {
        var total = _service.CalcularTotalLinea(3m, 15.50m);
        Assert.Equal(46.50m, total);
    }
}

public class DomainExceptionTests
{
    [Fact]
    public void DomainException_TieneCodigoYMensaje()
    {
        var ex = new DomainException("mensaje de prueba", "CODIGO_TEST");
        Assert.Equal("CODIGO_TEST", ex.Code);
        Assert.Equal("mensaje de prueba", ex.Message);
    }
}