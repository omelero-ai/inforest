using System.Reflection;
using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Application.Turno;
using Inforest.Domain.Entities.Caja;
using Inforest.Domain.Entities.Configuracion;
using TurnoEntity = Inforest.Domain.Entities.Configuracion.Turno;

namespace Inforest.Infrastructure.Turno;

internal sealed class TurnoRepository : ITurnoRepository
{
    private readonly IDbConnectionFactory _connectionFactory;
    private readonly ISpExecutor _spExecutor;

    public TurnoRepository(IDbConnectionFactory connectionFactory, ISpExecutor spExecutor)
    {
        _connectionFactory = connectionFactory;
        _spExecutor = spExecutor;
    }

    public async Task<TurnoEntity?> ObtenerTurnoActualAsync(string codigoCaja, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT TOP 1
                tTurno AS CodigoTurno,
                tCaja AS CodigoCaja,
                ISNULL(tSalon, '') AS CodigoSalon,
                ISNULL(tUsuario, '') AS CodigoUsuario,
                fInicial AS FechaApertura,
                fFinal AS FechaCierre,
                CAST(ISNULL(fInicial, GETDATE()) AS date) AS FechaDiaContable,
                CAST(ISNULL(nMontoIN, 0) AS decimal(18,2)) AS MontoInicial,
                CAST(ISNULL(nMontoIE, 0) AS decimal(18,2)) AS MontoInicialME,
                CAST(ISNULL(nMontoFN, 0) AS decimal(18,2)) AS MontoFinal,
                CAST(ISNULL(lCierre, 0) AS bit) AS Cerrado
            FROM MTURNO
            WHERE tCaja = @caja AND lCierre = 0
            ORDER BY fInicial DESC
            """;

        var record = await connection.QueryFirstOrDefaultAsync<TurnoRecord>(sql, new { caja = codigoCaja });
        return record is null ? null : Map(record);
    }

    public async Task<bool> InsertarAsync(TurnoEntity turno, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        // Legacy: INSERT INTO MTURNO(tTurno, tCaja, tSalon, fInicial, tUsuario, lCierre, nMontoIN, nMontoIE)
        // frmInicio.frm cmdOpcion_Click(0) — BR-TURNO-001.
        const string sql = """
            INSERT INTO MTURNO (tTurno, tCaja, tSalon, fInicial, tUsuario, lCierre, nMontoIN, nMontoIE)
            VALUES (@CodigoTurno, @CodigoCaja, @CodigoSalon, @FechaApertura, @CodigoUsuario, 0, @MontoInicial, @MontoInicialME)
            """;

        var affected = await connection.ExecuteAsync(sql, new
        {
            turno.CodigoTurno,
            turno.CodigoCaja,
            CodigoSalon = string.IsNullOrEmpty(turno.CodigoSalon) ? (object)DBNull.Value : turno.CodigoSalon,
            FechaApertura = turno.FechaApertura,
            turno.CodigoUsuario,
            turno.MontoInicial,
            turno.MontoInicialME
        });

        return affected > 0;
    }

    public async Task<TurnoExistente?> ObtenerUltimoTurnoAsync(
        string codigoCaja,
        string codigoUsuario,
        ModoConsultaTurno modo,
        CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);

        // Legacy: frmInicio.frm Form_Load — SELECT * FROM MTURNO WHERE ... ORDER BY tTurno; RsTurno.MoveLast
        var sql = modo switch
        {
            ModoConsultaTurno.PorUsuario =>
                "SELECT TOP 1 tTurno, tCaja, ISNULL(tUsuario,'') AS tUsuario, fInicial, CAST(ISNULL(lCierre,0) AS bit) AS lCierre, ISNULL(nMontoIN,0) AS nMontoIN, ISNULL(nMontoIE,0) AS nMontoIE FROM MTURNO WHERE tUsuario = @usuario ORDER BY tTurno DESC",
            ModoConsultaTurno.PorCajaYUsuario =>
                "SELECT TOP 1 tTurno, tCaja, ISNULL(tUsuario,'') AS tUsuario, fInicial, CAST(ISNULL(lCierre,0) AS bit) AS lCierre, ISNULL(nMontoIN,0) AS nMontoIN, ISNULL(nMontoIE,0) AS nMontoIE FROM MTURNO WHERE tCaja = @caja AND tUsuario = @usuario ORDER BY tTurno DESC",
            _ =>
                "SELECT TOP 1 tTurno, tCaja, ISNULL(tUsuario,'') AS tUsuario, fInicial, CAST(ISNULL(lCierre,0) AS bit) AS lCierre, ISNULL(nMontoIN,0) AS nMontoIN, ISNULL(nMontoIE,0) AS nMontoIE FROM MTURNO WHERE tCaja = @caja ORDER BY tTurno DESC"
        };

        var row = await connection.QueryFirstOrDefaultAsync<TurnoExistenteRow>(sql, new { caja = codigoCaja, usuario = codigoUsuario });
        if (row is null) return null;

        return new TurnoExistente(
            row.tTurno,
            row.tCaja,
            row.tUsuario,
            row.fInicial ?? DateTime.MinValue,
            row.lCierre,
            (decimal)row.nMontoIN,
            (decimal)row.nMontoIE);
    }

    public async Task<string> GenerarCorrelativoAsync(CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);

        // Legacy: frmInicio.frm — nCorrela = Calcular("SELECT MAX(tTurno) FROM MTURNO WHERE SUBSTRING(tTurno,1,2)=YY", Cn)
        // Si null o el año cambia → YY + "00000001"; sino YY + Correlativo(últimos 8 dígitos, 8)
        var yy = DateTime.Now.ToString("yy");
        var maxTurno = await connection.ExecuteScalarAsync<string?>(
            "SELECT MAX(tTurno) FROM MTURNO WHERE SUBSTRING(tTurno,1,2) = @yy",
            new { yy });

        if (string.IsNullOrEmpty(maxTurno) || maxTurno.Length < 10 || maxTurno[..2] != yy)
            return yy + "00000001";

        var secuencia = maxTurno[2..]; // últimos 8 dígitos
        if (!long.TryParse(secuencia, out var num))
            return yy + "00000001";

        return yy + (num + 1).ToString("D8");
    }

    public async Task<bool> ReAperturarAsync(
        string codigoTurno,
        string codigoUsuario,
        decimal montoInicial,
        decimal montoInicialME,
        CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        // Legacy: UPDATE MTURNO SET tUsuario, nMontoIN, nMontoIE WHERE tTurno (BR-TURNO-002)
        const string sql = """
            UPDATE MTURNO
            SET tUsuario = @usuario,
                nMontoIN = @montoN,
                nMontoIE = @montoE
            WHERE tTurno = @turno
            """;

        var affected = await connection.ExecuteAsync(sql, new
        {
            usuario = codigoUsuario,
            montoN = montoInicial,
            montoE = montoInicialME,
            turno = codigoTurno
        });
        return affected > 0;
    }

    public async Task<bool> CerrarAsync(string codigoTurno, decimal montoFinal, CierreTurnoBreakdown breakdown, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);

        // BR-CAJA-004: UPDATE MTURNO con desglose completo de montos.
        // Legacy: frmLiquidacionDetalle.frm — Isql = "Update MTURNO Set lCierre=1, fFinal=getdate(), ..."
        const string sql = """
            UPDATE MTURNO
            SET
                lCierre    = 1,
                fFinal     = GETDATE(),
                nMontoEN   = @EfectivoMN,
                nMontoEE   = @EfectivoME,
                nMontoEE2  = @EfectivoME2,
                nMontoEE3  = @EfectivoME3,
                nMontoCN   = @ChequesMN,
                nMontoCE   = @ChequesME,
                nMontoPN   = @PuntosMN,
                nMontoPE   = @PuntosME,
                nMontoFN   = @MontoFinalMN,
                nMontoFE   = @MontoFinalME,
                nMontoFE2  = @EfectivoME2,
                nMontoFE3  = @EfectivoME3,
                nTarjeta1  = @Tarjeta1,
                nTarjeta2  = @Tarjeta2,
                nTarjeta3  = @Tarjeta3,
                nTarjeta4  = @Tarjeta4,
                nTarjeta5  = @Tarjeta5,
                nTarjeta6  = @Tarjeta6,
                nTarjeta7  = @Tarjeta7,
                nTarjeta8  = @Tarjeta8,
                nIngresoN  = @IngresosMN,
                nIngresoE  = @IngresosME,
                nEgresoN   = @EgresosMN,
                nEgresoE   = @EgresosME
            WHERE tTurno = @codigo AND ISNULL(lCierre, 0) = 0
            """;

        var affected = await connection.ExecuteAsync(sql, new
        {
            breakdown.EfectivoMN,
            breakdown.EfectivoME,
            breakdown.EfectivoME2,
            breakdown.EfectivoME3,
            breakdown.ChequesMN,
            breakdown.ChequesME,
            breakdown.PuntosMN,
            breakdown.PuntosME,
            breakdown.MontoFinalMN,
            breakdown.MontoFinalME,
            Tarjeta1 = breakdown.ObtenerTarjeta(1),
            Tarjeta2 = breakdown.ObtenerTarjeta(2),
            Tarjeta3 = breakdown.ObtenerTarjeta(3),
            Tarjeta4 = breakdown.ObtenerTarjeta(4),
            Tarjeta5 = breakdown.ObtenerTarjeta(5),
            Tarjeta6 = breakdown.ObtenerTarjeta(6),
            Tarjeta7 = breakdown.ObtenerTarjeta(7),
            Tarjeta8 = breakdown.ObtenerTarjeta(8),
            breakdown.IngresosMN,
            breakdown.IngresosME,
            breakdown.EgresosMN,
            breakdown.EgresosME,
            codigo = codigoTurno
        });

        return affected > 0;
    }

    public async Task<IReadOnlyList<TurnoEntity>> ObtenerHistorialAsync(string codigoCaja, DateTime desde, CancellationToken ct = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", ct);
        const string sql = """
            SELECT
                tTurno AS CodigoTurno,
                tCaja AS CodigoCaja,
                ISNULL(tSalon, '') AS CodigoSalon,
                ISNULL(tUsuario, '') AS CodigoUsuario,
                fInicial AS FechaApertura,
                fFinal AS FechaCierre,
                CAST(ISNULL(fInicial, GETDATE()) AS date) AS FechaDiaContable,
                CAST(ISNULL(nMontoIN, 0) AS decimal(18,2)) AS MontoInicial,
                CAST(ISNULL(nMontoIE, 0) AS decimal(18,2)) AS MontoInicialME,
                CAST(ISNULL(nMontoFN, 0) AS decimal(18,2)) AS MontoFinal,
                CAST(ISNULL(lCierre, 0) AS bit) AS Cerrado
            FROM MTURNO
            WHERE tCaja = @caja AND fInicial >= @desde
            ORDER BY fInicial DESC
            """;

        var records = await connection.QueryAsync<TurnoRecord>(sql, new { caja = codigoCaja, desde });
        return records.Select(Map).ToList();
    }

    private static TurnoEntity Map(TurnoRecord record)
    {
        var turno = Domain.Entities.Configuracion.Turno.Abrir(
            record.CodigoTurno,
            record.CodigoCaja,
            record.CodigoUsuario,
            record.FechaDiaContable,
            record.MontoInicial,
            record.MontoInicialME,
            record.CodigoSalon);

        SetProperty(turno, nameof(Domain.Entities.Configuracion.Turno.FechaApertura), record.FechaApertura);

        if (record.Cerrado)
        {
            turno.Cerrar(record.MontoFinal);
            if (record.FechaCierre.HasValue)
                SetProperty(turno, nameof(Domain.Entities.Configuracion.Turno.FechaCierre), record.FechaCierre);
        }

        return turno;
    }

    private static void SetProperty<T>(object target, string propertyName, T value)
    {
        var property = target.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
            ?? throw new InvalidOperationException($"No se encontró la propiedad {propertyName}.");
        property.SetValue(target, value);
    }
}
