namespace Inforest.Infrastructure.Turno;

internal sealed class TurnoRecord
{
    public string CodigoTurno { get; init; } = string.Empty;
    public string CodigoCaja { get; init; } = string.Empty;
    public string CodigoUsuario { get; init; } = string.Empty;
    public DateTime FechaApertura { get; init; }
    public DateTime? FechaCierre { get; init; }
    public DateTime FechaDiaContable { get; init; }
    public decimal MontoInicial { get; init; }
    public decimal MontoFinal { get; init; }
    public bool Cerrado { get; init; }
}
