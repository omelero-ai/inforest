using Inforest.Domain.Entities.Motorizado;
using Inforest.Domain.Exceptions;

namespace Inforest.Domain.Tests.Delivery;

/// <summary>
/// Tests de las entidades Motorizado y AsignacionMotorizado.
/// Legacy: TMOTORIZADODATOS, vMotorizado, frmAsignacionMotorizado.frm, frmLlegadaSalida.frm
/// Reglas: BR-DEL-002, BR-DEL-006.
/// </summary>
public class MotorizadoDomainTests
{
    // ── Motorizado ─────────────────────────────────────────────────────────────

    [Fact]
    public void Motorizado_Crear_ConDatosValidos_CreaCorrectamente()
    {
        var m = Motorizado.Crear("MOT001", "12345678", 10m, 15m, 20m);

        Assert.Equal("MOT001", m.Codigo);
        Assert.Equal(10m, m.TarifaLunesViernes);
        Assert.Equal(15m, m.TarifaSabadoDomingo);
        Assert.Equal(20m, m.TarifaEspecial);
    }

    [Fact]
    public void Motorizado_Crear_SinCodigo_LanzaExcepcion()
    {
        var ex = Assert.Throws<DomainException>(() => Motorizado.Crear("", "12345678", 10, 15, 20));
        Assert.Equal("MOTORIZADO_CODIGO_REQUERIDO", ex.Code);
    }

    [Fact]
    public void Motorizado_Crear_TarifaNegativa_LanzaExcepcion()
    {
        var ex = Assert.Throws<DomainException>(() => Motorizado.Crear("MOT001", "12345678", -1, 0, 0));
        Assert.Equal("MOTORIZADO_TARIFA_NEGATIVA", ex.Code);
    }

    [Theory]
    [InlineData(TipoDiaTarifa.LunesViernes, 10)]
    [InlineData(TipoDiaTarifa.SabadoDomingo, 15)]
    [InlineData(TipoDiaTarifa.Especial, 20)]
    public void Motorizado_ObtenerTarifaPorDia_RetornaTarifaCorrecta(TipoDiaTarifa tipo, decimal esperada)
    {
        var m = Motorizado.Crear("MOT001", "12345678", 10m, 15m, 20m);
        Assert.Equal(esperada, m.ObtenerTarifaPorDia(tipo));
    }

    [Fact]
    public void Motorizado_ActualizarTarifas_CambiaValores()
    {
        var m = Motorizado.Crear("MOT001", "12345678", 10m, 15m, 20m);
        m.ActualizarTarifas(12m, 18m, 25m);

        Assert.Equal(12m, m.TarifaLunesViernes);
        Assert.Equal(18m, m.TarifaSabadoDomingo);
        Assert.Equal(25m, m.TarifaEspecial);
    }

    // ── AsignacionMotorizado ───────────────────────────────────────────────────

    [Fact]
    public void AsignacionMotorizado_Crear_ConDatosValidos_CreaCorrectamente()
    {
        var a = AsignacionMotorizado.Crear("PED001", "MOT001", "USR001");

        Assert.Equal("PED001", a.CodigoPedido);
        Assert.Equal("MOT001", a.CodigoMotorizado);
        Assert.False(a.EsReasignacion);
        Assert.Null(a.FechaSalida);
        Assert.Null(a.FechaLlegada);
    }

    [Fact]
    public void AsignacionMotorizado_RegistrarSalida_SetFechaSalida()
    {
        var a = AsignacionMotorizado.Crear("PED001", "MOT001", "USR001");
        a.RegistrarSalida();

        Assert.NotNull(a.FechaSalida);
    }

    [Fact]
    public void AsignacionMotorizado_RegistrarSalidaDoble_LanzaExcepcion()
    {
        var a = AsignacionMotorizado.Crear("PED001", "MOT001", "USR001");
        a.RegistrarSalida();

        var ex = Assert.Throws<DomainException>(() => a.RegistrarSalida());
        Assert.Equal("ASIGNACION_SALIDA_YA_REGISTRADA", ex.Code);
    }

    [Fact]
    public void AsignacionMotorizado_RegistrarLlegadaSinSalida_LanzaExcepcion()
    {
        var a = AsignacionMotorizado.Crear("PED001", "MOT001", "USR001");

        var ex = Assert.Throws<DomainException>(() => a.RegistrarLlegada());
        Assert.Equal("ASIGNACION_SALIDA_PENDIENTE", ex.Code);
    }

    [Fact]
    public void AsignacionMotorizado_RegistrarLlegadaConSalida_SetFechaLlegada()
    {
        var a = AsignacionMotorizado.Crear("PED001", "MOT001", "USR001");
        a.RegistrarSalida();
        a.RegistrarLlegada();

        Assert.NotNull(a.FechaLlegada);
    }

    [Fact]
    public void AsignacionMotorizado_Reasignar_CreaReasignacion()
    {
        var a = AsignacionMotorizado.Crear("PED001", "MOT001", "USR001");
        var r = a.Reasignar("MOT002", "SUPERVISOR");

        Assert.Equal("MOT002", r.CodigoMotorizado);
        Assert.True(r.EsReasignacion);
        Assert.Equal("SUPERVISOR", r.UsuarioReasignacion);
    }
}
