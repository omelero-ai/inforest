namespace Inforest.Infrastructure.Turno;

internal sealed class TurnoRecord
{
    public string CodigoTurno { get; init; } = string.Empty;
    public string CodigoCaja { get; init; } = string.Empty;
    public string CodigoSalon { get; init; } = string.Empty;
    public string CodigoUsuario { get; init; } = string.Empty;
    public DateTime FechaApertura { get; init; }
    public DateTime? FechaCierre { get; init; }
    public DateTime FechaDiaContable { get; init; }
    public decimal MontoInicial { get; init; }
    public decimal MontoInicialME { get; init; }
    public decimal MontoFinal { get; init; }
    public bool Cerrado { get; init; }
}

/// <summary>
/// DTO de fila para consulta del último turno (cualquier estado).
/// Legacy: RsTurno recordset de frmInicio.frm Form_Load.
/// </summary>
internal sealed class TurnoExistenteRow
{
    public string tTurno { get; init; } = string.Empty;
    public string tCaja { get; init; } = string.Empty;
    public string tUsuario { get; init; } = string.Empty;
    public DateTime? fInicial { get; init; }
    public bool lCierre { get; init; }
    public double nMontoIN { get; init; }
    public double nMontoIE { get; init; }
}
