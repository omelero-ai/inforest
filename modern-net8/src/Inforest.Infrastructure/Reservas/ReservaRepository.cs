using Dapper;
using Inforest.Application.Interfaces;
using Inforest.Domain.Entities.Reservas;
using Inforest.Domain.Repositories;

namespace Inforest.Infrastructure.Reservas;

/// <summary>
/// Repositorio de reservas sobre tabla TRESERVA.
/// Legacy: frmReserva.frm, frmReservaDetalle.frm
/// BR-RESERVA-001/002/003/004
/// </summary>
internal sealed class ReservaRepository : IReservaRepository
{
    private readonly IDbConnectionFactory _connectionFactory;

    public ReservaRepository(IDbConnectionFactory connectionFactory)
        => _connectionFactory = connectionFactory;

    public async Task<Reserva?> ObtenerPorCodigoAsync(string codigoReserva, CancellationToken cancellationToken = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        const string sql = """
            SELECT
                tReserva          AS CodigoReserva,
                ISNULL(fRegistro, GETDATE()) AS FechaRegistro,
                fFecha            AS FechaReserva,
                ISNULL(fHora,'') AS HoraReserva,
                NULLIF(tCodigoDelivery,'') AS CodigoCliente,
                NULLIF(tApellido + ' ' + ISNULL(tNombre,''), ' ') AS NombreCliente,
                NULLIF(tTelefono,'') AS Telefono,
                NULLIF(tMotivo,'') AS CodigoMotivo,
                NULLIF(tMesero,'') AS CodigoMesero,
                ISNULL(nPax,0)    AS Pax,
                NULLIF(tMesa,'')  AS CodigoMesa,
                ISNULL(tEstadoReserva,'01') AS EstadoRaw,
                NULLIF(tObservacion,'') AS Observacion,
                fDiaContable      AS FechaDiaContable
            FROM TRESERVA
            WHERE tReserva = @codigo
            """;

        var row = await connection.QueryFirstOrDefaultAsync<ReservaRow>(sql, new { codigo = codigoReserva });
        return row is null ? null : Map(row);
    }

    public async Task<IEnumerable<Reserva>> ObtenerPorFechaAsync(DateTime fechaInicio, DateTime fechaFin, CancellationToken cancellationToken = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        const string sql = """
            SELECT
                tReserva          AS CodigoReserva,
                ISNULL(fRegistro, GETDATE()) AS FechaRegistro,
                fFecha            AS FechaReserva,
                ISNULL(fHora,'') AS HoraReserva,
                NULLIF(tCodigoDelivery,'') AS CodigoCliente,
                NULLIF(tApellido + ' ' + ISNULL(tNombre,''), ' ') AS NombreCliente,
                NULLIF(tTelefono,'') AS Telefono,
                NULLIF(tMotivo,'') AS CodigoMotivo,
                NULLIF(tMesero,'') AS CodigoMesero,
                ISNULL(nPax,0)    AS Pax,
                NULLIF(tMesa,'')  AS CodigoMesa,
                ISNULL(tEstadoReserva,'01') AS EstadoRaw,
                NULLIF(tObservacion,'') AS Observacion,
                fDiaContable      AS FechaDiaContable
            FROM TRESERVA
            WHERE tPrioridad = '1'
              AND fFecha >= @fechaInicio
              AND fFecha <  DATEADD(day,1,@fechaFin)
            ORDER BY tReserva
            """;

        var rows = await connection.QueryAsync<ReservaRow>(sql, new { fechaInicio, fechaFin });
        return rows.Select(Map);
    }

    /// <summary>
    /// Genera el próximo código de reserva siguiendo la lógica Legacy:
    /// YY + correlativo de 8 dígitos por año.
    /// Legacy: frmReservaDetalle.frm — max(tReserva) substring año
    /// </summary>
    public async Task<string> GenerarProximoCodigoAsync(CancellationToken cancellationToken = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        const string sql = """
            SELECT ISNULL(max(tReserva),'')
            FROM TRESERVA
            WHERE substring(tReserva,1,2) = substring(ltrim(str(year(getdate()))),3,2)
            """;

        var maxCodigo = await connection.ExecuteScalarAsync<string>(sql) ?? string.Empty;
        var yy = DateTime.Now.Year.ToString()[2..]; // últimos 2 dígitos

        if (string.IsNullOrEmpty(maxCodigo) || maxCodigo[..2] != yy)
            return yy + "00000001";

        var numero = int.Parse(maxCodigo[2..]) + 1;
        return yy + numero.ToString().PadLeft(8, '0');
    }

    public async Task InsertarAsync(Reserva reserva, CancellationToken cancellationToken = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        const string sql = """
            INSERT INTO TRESERVA
                (tReserva, fRegistro, tCodigoDelivery, tTelefono, tMotivo, tMesero,
                 fFecha, fHora, nPax, tEstadoReserva, tObservacion,
                 tPrioridad, tPrioridad2, tMesa, fDiaContable)
            VALUES
                (@tReserva, GETDATE(), @tCodigoDelivery, @tTelefono, @tMotivo, @tMesero,
                 @fFecha, @fHora, @nPax, @tEstadoReserva, @tObservacion,
                 '1', 0, @tMesa, @fDiaContable)
            """;

        await connection.ExecuteAsync(sql, ToParams(reserva));
    }

    public async Task ActualizarAsync(Reserva reserva, CancellationToken cancellationToken = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        const string sql = """
            UPDATE TRESERVA SET
                tCodigoDelivery = @tCodigoDelivery,
                tTelefono       = @tTelefono,
                tMotivo         = @tMotivo,
                tMesero         = @tMesero,
                fFecha          = @fFecha,
                fHora           = @fHora,
                nPax            = @nPax,
                tEstadoReserva  = @tEstadoReserva,
                tObservacion    = @tObservacion,
                tMesa           = @tMesa,
                fFechaModificacion = GETDATE()
            WHERE tReserva = @tReserva
            """;

        await connection.ExecuteAsync(sql, ToParams(reserva));
    }

    /// <summary>
    /// Ejecuta spIns_MPEDIDO_RESERVA y retorna el código de pedido generado (output @tPedido).
    /// Legacy: frmReservaDetalle.frm — cmdOpcion_Click Case 4 (oComando ExecSP + GetParameterValue)
    /// </summary>
    public async Task<string> ConvertirAPedidoAsync(
        string codigoReserva,
        string codigoCaja,
        string codigoUsuario,
        string codigoTurno,
        DateTime fechaDiaContable,
        CancellationToken cancellationToken = default)
    {
        using var connection = await _connectionFactory.CreateOpenConnectionAsync("Inforest", cancellationToken);
        var parameters = new DynamicParameters();
        parameters.Add("@tCaja",        codigoCaja);
        parameters.Add("@tUsuario",     codigoUsuario);
        parameters.Add("@tTurno",       codigoTurno);
        parameters.Add("@fDiaContable", fechaDiaContable);
        parameters.Add("@fProgramacion", fechaDiaContable);
        parameters.Add("@codigoReserva", codigoReserva);
        parameters.Add("@tPedido", dbType: System.Data.DbType.String, direction: System.Data.ParameterDirection.Output, size: 15);

        await connection.ExecuteAsync("spIns_MPEDIDO_RESERVA", parameters, commandType: System.Data.CommandType.StoredProcedure);
        return parameters.Get<string?>("@tPedido") ?? string.Empty;
    }

    // ──────────────────────────────────────────────────────────────────────────
    // Helpers
    // ──────────────────────────────────────────────────────────────────────────

    private static Reserva Map(ReservaRow row)
    {
        var estado = row.EstadoRaw switch
        {
            "02" => EstadoReserva.Atendida,
            "03" => EstadoReserva.Anulada,
            _    => EstadoReserva.Pendiente
        };

        return Reserva.Reconstruir(
            row.CodigoReserva,
            row.FechaRegistro,
            row.FechaReserva,
            row.HoraReserva,
            row.CodigoCliente,
            row.NombreCliente,
            row.Telefono,
            row.CodigoMotivo,
            row.CodigoMesero,
            row.Pax,
            row.CodigoMesa,
            estado,
            row.Observacion,
            row.FechaDiaContable);
    }

    private static object ToParams(Reserva r) => new
    {
        tReserva        = r.CodigoReserva,
        tCodigoDelivery = r.CodigoCliente ?? string.Empty,
        tTelefono       = r.Telefono ?? string.Empty,
        tMotivo         = r.CodigoMotivo ?? string.Empty,
        tMesero         = r.CodigoMesero ?? string.Empty,
        fFecha          = r.FechaReserva,
        fHora           = r.HoraReserva,
        nPax            = r.Pax,
        tEstadoReserva  = r.Estado switch
        {
            EstadoReserva.Atendida => "02",
            EstadoReserva.Anulada  => "03",
            _                      => "01"
        },
        tObservacion    = r.Observacion != null ? r.Observacion[..Math.Min(r.Observacion.Length, 200)] : string.Empty,
        tMesa           = r.CodigoMesa ?? string.Empty,
        fDiaContable    = (object?)(r.FechaDiaContable.HasValue ? r.FechaDiaContable.Value : DBNull.Value)
    };

    private sealed class ReservaRow
    {
        public string CodigoReserva { get; init; } = string.Empty;
        public DateTime FechaRegistro { get; init; }
        public DateTime FechaReserva { get; init; }
        public string HoraReserva { get; init; } = string.Empty;
        public string? CodigoCliente { get; init; }
        public string? NombreCliente { get; init; }
        public string? Telefono { get; init; }
        public string? CodigoMotivo { get; init; }
        public string? CodigoMesero { get; init; }
        public int Pax { get; init; }
        public string? CodigoMesa { get; init; }
        public string EstadoRaw { get; init; } = "01";
        public string? Observacion { get; init; }
        public DateTime? FechaDiaContable { get; init; }
    }
}
