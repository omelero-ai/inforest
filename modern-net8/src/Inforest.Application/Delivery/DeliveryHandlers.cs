using Inforest.Domain.Common;
using Inforest.Domain.Entities.Delivery;
using Inforest.Domain.Exceptions;
using Inforest.Domain.Repositories;

namespace Inforest.Application.Delivery;

// ──────────────────────────────────────────────────────────────────────────────
// COMMANDS
// ──────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Comando para crear un nuevo cliente de delivery.
/// <para>
/// Legacy: <c>frmNuevoDelivery.frm</c> — acción de guardar cliente frecuente.
/// Tabla: <c>TDELIVERY</c>.
/// </para>
/// </summary>
public sealed record CrearClienteDeliveryCommand(
    string CodigoDelivery,
    string? TipoCliente,
    string? Apellido,
    string? Nombre,
    string? Telefono,
    string? Direccion,
    string? Referencia,
    string? CodigoZona,
    string? CodigoDistrito,
    string? TipoIdentidad,
    string? NumeroIdentidad,
    string? CodigoTarjeta,
    string? NumeroTarjeta,
    string? Observacion,
    string? Email);

public sealed record ObtenerSiguienteCodigoClienteDeliveryQuery;

public sealed class ObtenerSiguienteCodigoClienteDeliveryHandler
{
    private readonly IClienteDeliveryRepository _repo;

    public ObtenerSiguienteCodigoClienteDeliveryHandler(IClienteDeliveryRepository repo)
        => _repo = repo;

    public async Task<Result<string>> HandleAsync(ObtenerSiguienteCodigoClienteDeliveryQuery _, CancellationToken ct = default)
    {
        var maximo = await _repo.ObtenerMaximoCodigoAsync(ct);
        return Result.Ok(CodigoClienteDeliveryGenerator.Generar(maximo));
    }
}

internal static class CodigoClienteDeliveryGenerator
{
    public static string Generar(string? maximoActual)
    {
        var codigo = (maximoActual ?? string.Empty).Trim().ToUpperInvariant();
        if (string.IsNullOrWhiteSpace(codigo))
            return "0000001";

        if (codigo.All(char.IsDigit))
            return (long.Parse(codigo) + 1).ToString(new string('0', Math.Max(7, codigo.Length)));

        var prefijo = codigo[0];
        var sufijo = codigo.Length > 1 ? codigo[1..] : "000001";
        if (!sufijo.All(char.IsDigit))
            return "0000001";

        var numero = int.Parse(sufijo);
        if (numero >= int.Parse(new string('9', sufijo.Length)))
            return $"{(char)(prefijo + 1)}{new string('0', sufijo.Length - 1)}1";

        return $"{prefijo}{(numero + 1).ToString(new string('0', sufijo.Length))}";
    }
}

/// <summary>
/// Handler de <see cref="CrearClienteDeliveryCommand"/>.
/// </summary>
public sealed class CrearClienteDeliveryHandler
{
    private readonly IClienteDeliveryRepository _repo;

    public CrearClienteDeliveryHandler(IClienteDeliveryRepository repo)
        => _repo = repo;

    public async Task<Result> HandleAsync(CrearClienteDeliveryCommand cmd, CancellationToken ct = default)
    {
        var codigoDelivery = string.IsNullOrWhiteSpace(cmd.CodigoDelivery)
            ? CodigoClienteDeliveryGenerator.Generar(await _repo.ObtenerMaximoCodigoAsync(ct))
            : cmd.CodigoDelivery.Trim().ToUpperInvariant();

        if (string.IsNullOrWhiteSpace(cmd.Telefono) && string.IsNullOrWhiteSpace(cmd.NumeroIdentidad))
            return Result.Fail("Ingrese el teléfono o número de identidad.", "DELIVERY_TELEFONO_O_IDENTIDAD_REQUERIDO");

        var existente = await _repo.ObtenerPorCodigoAsync(codigoDelivery, ct);
        if (existente is not null)
            return Result.Fail("Ya existe un cliente delivery con ese código.", "DELIVERY_CLIENTE_YA_EXISTE");

        if (!string.IsNullOrWhiteSpace(cmd.Telefono))
        {
            var existenteTelefono = await _repo.ObtenerPorTelefonoAsync(cmd.Telefono.Trim(), ct);
            if (existenteTelefono is not null)
                return Result.Fail("Teléfono o Id existente.", "DELIVERY_TELEFONO_EXISTENTE");
        }

        if (!string.IsNullOrWhiteSpace(cmd.TipoIdentidad) && !string.IsNullOrWhiteSpace(cmd.NumeroIdentidad))
        {
            var existenteIdentidad = await _repo.ObtenerPorIdentidadAsync(cmd.TipoIdentidad.Trim(), cmd.NumeroIdentidad.Trim(), ct);
            if (existenteIdentidad is not null)
                return Result.Fail("DNI existente.", "DELIVERY_IDENTIDAD_EXISTENTE");
        }

        var cliente = ClienteDelivery.Crear(
            codigoDelivery,
            cmd.TipoCliente?.Trim(),
            cmd.Apellido?.Trim().ToUpperInvariant(),
            cmd.Nombre?.Trim().ToUpperInvariant(),
            cmd.Telefono?.Trim(),
            cmd.Direccion?.Trim().ToUpperInvariant(),
            cmd.CodigoZona?.Trim(),
            cmd.CodigoDistrito?.Trim());

        if (!string.IsNullOrWhiteSpace(cmd.TipoIdentidad) && !string.IsNullOrWhiteSpace(cmd.NumeroIdentidad))
            cliente.AsignarIdentidad(cmd.TipoIdentidad.Trim(), cmd.NumeroIdentidad.Trim());

        if (!string.IsNullOrWhiteSpace(cmd.CodigoTarjeta) && !string.IsNullOrWhiteSpace(cmd.NumeroTarjeta))
            cliente.AsignarTarjeta(cmd.CodigoTarjeta.Trim(), cmd.NumeroTarjeta.Trim());

        await _repo.InsertarAsync(cliente, ct);
        return Result.Ok();
    }
}

/// <summary>
/// Comando para actualizar datos de un cliente delivery existente.
/// <para>
/// Legacy: <c>frmNuevoDelivery.frm</c> — edición de cliente frecuente.
/// </para>
/// </summary>
public sealed record ActualizarClienteDeliveryCommand(
    string CodigoDelivery,
    string? Apellido,
    string? Nombre,
    string? Telefono,
    string? Direccion,
    string? Referencia,
    string? CodigoZona,
    string? CodigoDistrito,
    string? TipoIdentidad,
    string? NumeroIdentidad,
    string? Observacion,
    string? Email);

/// <summary>Handler de <see cref="ActualizarClienteDeliveryCommand"/>.</summary>
public sealed class ActualizarClienteDeliveryHandler
{
    private readonly IClienteDeliveryRepository _repo;

    public ActualizarClienteDeliveryHandler(IClienteDeliveryRepository repo)
        => _repo = repo;

    public async Task<Result> HandleAsync(ActualizarClienteDeliveryCommand cmd, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(cmd.Telefono) && string.IsNullOrWhiteSpace(cmd.NumeroIdentidad))
            return Result.Fail("Ingrese el teléfono o número de identidad.", "DELIVERY_TELEFONO_O_IDENTIDAD_REQUERIDO");

        var cliente = await _repo.ObtenerPorCodigoAsync(cmd.CodigoDelivery, ct);
        if (cliente is null)
            return Result.Fail("No se encontró el cliente delivery.", "DELIVERY_CLIENTE_NO_ENCONTRADO");

        if (!string.IsNullOrWhiteSpace(cmd.Telefono))
        {
            var existenteTelefono = await _repo.ObtenerPorTelefonoAsync(cmd.Telefono.Trim(), ct);
            if (existenteTelefono is not null && !string.Equals(existenteTelefono.CodigoDelivery, cmd.CodigoDelivery, StringComparison.OrdinalIgnoreCase))
                return Result.Fail("Teléfono o Id existente.", "DELIVERY_TELEFONO_EXISTENTE");
        }

        if (!string.IsNullOrWhiteSpace(cmd.TipoIdentidad) && !string.IsNullOrWhiteSpace(cmd.NumeroIdentidad))
        {
            var existenteIdentidad = await _repo.ObtenerPorIdentidadAsync(cmd.TipoIdentidad.Trim(), cmd.NumeroIdentidad.Trim(), ct);
            if (existenteIdentidad is not null && !string.Equals(existenteIdentidad.CodigoDelivery, cmd.CodigoDelivery, StringComparison.OrdinalIgnoreCase))
                return Result.Fail("DNI existente.", "DELIVERY_IDENTIDAD_EXISTENTE");

            cliente.AsignarIdentidad(cmd.TipoIdentidad.Trim(), cmd.NumeroIdentidad.Trim());
        }

        cliente.Actualizar(
            cmd.Apellido?.Trim().ToUpperInvariant(),
            cmd.Nombre?.Trim().ToUpperInvariant(),
            cmd.Telefono?.Trim(),
            cmd.Direccion?.Trim().ToUpperInvariant(),
            cmd.Referencia?.Trim().ToUpperInvariant(),
            cmd.CodigoZona?.Trim(),
            cmd.CodigoDistrito?.Trim(),
            cmd.Observacion?.Trim().ToUpperInvariant(),
            cmd.Email?.Trim());

        await _repo.ActualizarAsync(cliente, ct);
        return Result.Ok();
    }
}

/// <summary>
/// Comando para actualizar la foto de un cliente delivery.
/// <para>
/// Legacy: <c>frmClienteDeliveryDetalle.frm → GuardarFoto()</c>.
/// SP: <c>sp_UpdFotoDelivery</c> sobre tabla <c>TDELIVERY.iFoto</c>.
/// </para>
/// </summary>
public sealed record ActualizarFotoClienteDeliveryCommand(
    string CodigoDelivery,
    byte[] Foto);

/// <summary>Handler de <see cref="ActualizarFotoClienteDeliveryCommand"/>.</summary>
public sealed class ActualizarFotoClienteDeliveryHandler
{
    private readonly IClienteDeliveryRepository _repo;

    public ActualizarFotoClienteDeliveryHandler(IClienteDeliveryRepository repo)
        => _repo = repo;

    public async Task<Result> HandleAsync(ActualizarFotoClienteDeliveryCommand cmd, CancellationToken ct = default)
    {
        var cliente = await _repo.ObtenerPorCodigoAsync(cmd.CodigoDelivery, ct);
        if (cliente is null)
            return Result.Fail("No se encontró el cliente delivery.", "DELIVERY_CLIENTE_NO_ENCONTRADO");

        if (cmd.Foto.Length == 0)
            return Result.Fail("La foto no puede estar vacía.", "DELIVERY_FOTO_VACIA");

        await _repo.ActualizarFotoAsync(cmd.CodigoDelivery, cmd.Foto, ct);
        return Result.Ok();
    }
}

/// <summary>
/// Comando para crear un pedido delivery.
/// <para>
/// Legacy: <c>frmNuevoDelivery.frm</c> + <c>frmPedido.frm</c> en módulo Despachador.
/// Tabla: <c>MPEDIDO</c> con campos de delivery.
/// </para>
/// Reglas: BR-DEL-001, BR-DEL-010.
/// </summary>
public sealed record CrearPedidoDeliveryCommand(
    string CodigoPedido,
    string CodigoTurno,
    string CodigoUsuario,
    string CodigoCaja,
    string CodigoClienteDelivery,
    string? CodigoZona,
    string? CodigoDistrito,
    string? TipoPedido,
    DateTime? HoraEstimadaEntrega);

/// <summary>Handler de <see cref="CrearPedidoDeliveryCommand"/>.</summary>
public sealed class CrearPedidoDeliveryHandler
{
    private readonly IPedidoDeliveryRepository _pedidoRepo;

    public CrearPedidoDeliveryHandler(IPedidoDeliveryRepository pedidoRepo)
        => _pedidoRepo = pedidoRepo;

    public async Task<Result> HandleAsync(CrearPedidoDeliveryCommand cmd, CancellationToken ct = default)
    {
        var pedido = PedidoDelivery.Crear(
            cmd.CodigoPedido,
            cmd.CodigoTurno,
            cmd.CodigoUsuario,
            cmd.CodigoCaja,
            cmd.CodigoClienteDelivery,
            cmd.CodigoZona,
            cmd.CodigoDistrito,
            cmd.TipoPedido);

        if (cmd.HoraEstimadaEntrega.HasValue)
            pedido.EstablecerHoraEstimada(cmd.HoraEstimadaEntrega.Value);

        await _pedidoRepo.InsertarAsync(pedido, ct);
        return Result.Ok();
    }
}

/// <summary>
/// Comando para actualizar el estado de un pedido delivery.
/// <para>
/// Legacy: encapsula <c>sp_CD_Modificar_EstadoDelivery_Cabecera</c> en CENTRALDELIVERY.
/// </para>
/// Regla BR-DEL-009.
/// </summary>
public sealed record ActualizarEstadoDeliveryCommand(
    string CodigoPedido,
    EstadoDelivery NuevoEstado,
    string Usuario);

/// <summary>Handler de <see cref="ActualizarEstadoDeliveryCommand"/>.</summary>
public sealed class ActualizarEstadoDeliveryHandler
{
    private readonly IPedidoDeliveryRepository _pedidoRepo;
    private readonly ICentralPedidosRepository _centralRepo;

    public ActualizarEstadoDeliveryHandler(
        IPedidoDeliveryRepository pedidoRepo,
        ICentralPedidosRepository centralRepo)
    {
        _pedidoRepo = pedidoRepo;
        _centralRepo = centralRepo;
    }

    public async Task<Result> HandleAsync(ActualizarEstadoDeliveryCommand cmd, CancellationToken ct = default)
    {
        var pedido = await _pedidoRepo.ObtenerPorCodigoAsync(cmd.CodigoPedido, ct);
        if (pedido is null)
            return Result.Fail("No se encontró el pedido delivery.", "DELIVERY_PEDIDO_NO_ENCONTRADO");

        pedido.ActualizarEstado(cmd.NuevoEstado);
        await _pedidoRepo.ActualizarEstadoAsync(cmd.CodigoPedido, cmd.NuevoEstado, ct);

        // Propagar al Central de Pedidos si la conexión está disponible (BR-DEL-003, BR-DEL-009)
        if (await _centralRepo.VerificarConexionAsync(ct))
        {
            await _centralRepo.ModificarEstadoDeliveryAsync(
                cmd.CodigoPedido,
                cmd.NuevoEstado.ToString(),
                cmd.Usuario,
                ct);
        }

        return Result.Ok();
    }
}

// ──────────────────────────────────────────────────────────────────────────────
// QUERIES
// ──────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Query para buscar clientes de delivery.
/// <para>
/// Legacy: vista <c>vDelivery</c> con filtros en <c>frmClienteDelivery.frm</c>.
/// </para>
/// </summary>
public sealed record BuscarClienteDeliveryQuery(
    string? Apellido = null,
    string? Nombre = null,
    string? Telefono = null,
    string? CodigoDistrito = null);

/// <summary>Handler de <see cref="BuscarClienteDeliveryQuery"/>.</summary>
public sealed class BuscarClienteDeliveryHandler
{
    private readonly IClienteDeliveryRepository _repo;

    public BuscarClienteDeliveryHandler(IClienteDeliveryRepository repo)
        => _repo = repo;

    public async Task<Result<IEnumerable<ClienteDelivery>>> HandleAsync(
        BuscarClienteDeliveryQuery query, CancellationToken ct = default)
    {
        var clientes = await _repo.BuscarAsync(query.Apellido, query.Nombre, query.Telefono, query.CodigoDistrito, ct);
        return Result.Ok(clientes);
    }
}

public sealed record ObtenerClienteDeliveryPorCodigoQuery(string CodigoDelivery);

public sealed class ObtenerClienteDeliveryPorCodigoHandler
{
    private readonly IClienteDeliveryRepository _repo;

    public ObtenerClienteDeliveryPorCodigoHandler(IClienteDeliveryRepository repo)
        => _repo = repo;

    public async Task<Result<ClienteDelivery?>> HandleAsync(ObtenerClienteDeliveryPorCodigoQuery query, CancellationToken ct = default)
    {
        var cliente = await _repo.ObtenerPorCodigoAsync(query.CodigoDelivery, ct);
        return Result.Ok(cliente);
    }
}

/// <summary>
/// Query para obtener clientes frecuentes.
/// <para>
/// Legacy: <c>frmClienteDeliveryFrecuente.frm</c>.
/// </para>
/// </summary>
public sealed record ObtenerClientesFrecuentesQuery(int Top = 20);

/// <summary>Handler de <see cref="ObtenerClientesFrecuentesQuery"/>.</summary>
public sealed class ObtenerClientesFrecuentesHandler
{
    private readonly IClienteDeliveryRepository _repo;

    public ObtenerClientesFrecuentesHandler(IClienteDeliveryRepository repo)
        => _repo = repo;

    public async Task<Result<IEnumerable<ClienteDelivery>>> HandleAsync(
        ObtenerClientesFrecuentesQuery query, CancellationToken ct = default)
    {
        var clientes = await _repo.ObtenerFrecuentesAsync(query.Top, ct);
        return Result.Ok(clientes);
    }
}

/// <summary>
/// Query para obtener pedidos delivery pendientes.
/// <para>
/// Legacy: vista <c>vDelivery</c> con filtros de fecha/estado/zona en <c>frmNuevoDelivery.frm</c>.
/// </para>
/// </summary>
public sealed record ObtenerDeliveryPendientesQuery(
    string CodigoCaja,
    DateTime FechaInicio,
    DateTime FechaFin);

/// <summary>Handler de <see cref="ObtenerDeliveryPendientesQuery"/>.</summary>
public sealed class ObtenerDeliveryPendientesHandler
{
    private readonly IPedidoDeliveryRepository _repo;

    public ObtenerDeliveryPendientesHandler(IPedidoDeliveryRepository repo)
        => _repo = repo;

    public async Task<Result<IEnumerable<PedidoDelivery>>> HandleAsync(
        ObtenerDeliveryPendientesQuery query, CancellationToken ct = default)
    {
        var pedidos = await _repo.ObtenerPendientesAsync(query.CodigoCaja, query.FechaInicio, query.FechaFin, ct);
        return Result.Ok(pedidos);
    }
}

// ──────────────────────────────────────────────────────────────────────────────
// CENTRAL PEDIDOS — frmCentralPedidos.frm
// ──────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Comando para confirmar la entrega de un pedido desde el Central de Pedidos.
/// <para>
/// Legacy: <c>frmCentralPedidos.frm Case 3</c>.
/// SQL: <c>Update MPEDIDO Set lEntregado='1', tusuarioentregado=@usuario, fregentregado=getdate()</c>.
/// </para>
/// Reglas: BR-DEL-012, BR-DEL-013.
/// </summary>
public sealed record ConfirmarEntregaCentralCommand(
    string CodigoPedido,
    string Usuario,
    bool SupervisorAutorizado = false);

/// <summary>Handler de <see cref="ConfirmarEntregaCentralCommand"/>.</summary>
public sealed class ConfirmarEntregaCentralHandler
{
    private readonly IPedidoDeliveryRepository _repo;

    public ConfirmarEntregaCentralHandler(IPedidoDeliveryRepository repo)
        => _repo = repo;

    /// <summary>
    /// Confirma la entrega del pedido.
    /// <list type="bullet">
    ///   <item>BR-DEL-012: no se puede confirmar entrega si ya fue entregado.</item>
    ///   <item>BR-DEL-013: si el pedido no ha sido pagado (POR COBRAR/ANTICIPO) se requiere
    ///     autorización de supervisor (acción "22") para proceder.</item>
    /// </list>
    /// </summary>
    public async Task<Result> HandleAsync(ConfirmarEntregaCentralCommand cmd, CancellationToken ct = default)
    {
        // BR-DEL-012: verificar si ya fue entregado
        var yaEntregado = await _repo.EstaEntregadoAsync(cmd.CodigoPedido, ct);
        if (yaEntregado)
            return Result.Fail("El pedido ya se encuentra Entregado.", "DELIVERY_YA_ENTREGADO");

        // BR-DEL-013: verificar estado de pago
        var estadoPago = await _repo.ObtenerEstadoPagoAsync(cmd.CodigoPedido, ct);
        if (estadoPago is "NO PAGADO" or "POR COBRAR" or "ANTICIPO")
        {
            if (!cmd.SupervisorAutorizado)
                return Result.Fail(
                    "El pedido no ha sido Cancelado. Se requiere autorización de supervisor (acción 22).",
                    "REQUIERE_SUPERVISOR_22");
        }

        await _repo.ConfirmarEntregaAsync(cmd.CodigoPedido, cmd.Usuario, ct);
        return Result.Ok();
    }
}

/// <summary>
/// Comando para revertir la confirmación de entrega de un pedido.
/// <para>
/// Legacy: <c>frmCentralPedidos.frm Case 5</c> — requiere supervisor acción "22".
/// SQL: <c>Update MPEDIDO Set lEntregado='0', tusuarioentregado=@usuario, fregentregado=getdate()</c>.
/// </para>
/// Regla: BR-DEL-012.
/// </summary>
public sealed record RevertirEntregaCentralCommand(
    string CodigoPedido,
    string Usuario,
    bool SupervisorAutorizado = false);

/// <summary>Handler de <see cref="RevertirEntregaCentralCommand"/>.</summary>
public sealed class RevertirEntregaCentralHandler
{
    private readonly IPedidoDeliveryRepository _repo;

    public RevertirEntregaCentralHandler(IPedidoDeliveryRepository repo)
        => _repo = repo;

    /// <summary>
    /// Revierte la entrega del pedido. Requiere autorización de supervisor (acción "22").
    /// BR-DEL-012: el pedido debe estar previamente entregado para poder revertir.
    /// </summary>
    public async Task<Result> HandleAsync(RevertirEntregaCentralCommand cmd, CancellationToken ct = default)
    {
        // Legacy: Supervisor("22") — siempre requerido para revertir (Case 5)
        if (!cmd.SupervisorAutorizado)
            return Result.Fail(
                "Se requiere autorización de supervisor (acción 22) para revertir la entrega.",
                "REQUIERE_SUPERVISOR_22");

        var yaEntregado = await _repo.EstaEntregadoAsync(cmd.CodigoPedido, ct);
        if (!yaEntregado)
            return Result.Fail("El pedido no ha sido Entregado.", "DELIVERY_NO_ENTREGADO");

        await _repo.RevertirEntregaAsync(cmd.CodigoPedido, cmd.Usuario, ct);
        return Result.Ok();
    }
}

/// <summary>
/// Comando para modificar la fecha programada de entrega de un pedido.
/// <para>
/// Legacy: <c>frmCentralPedidos.frm Case 2</c>.
/// SQL: <c>Update MPEDIDO set fregistro=@fecha, fProgramacion=@fecha where tCodigoPedido=@sPedido</c>.
/// </para>
/// Regla: BR-DEL-014.
/// </summary>
public sealed record ModificarFechaProgramadaDeliveryCommand(
    string CodigoPedido,
    DateTime NuevaFecha);

/// <summary>Handler de <see cref="ModificarFechaProgramadaDeliveryCommand"/>.</summary>
public sealed class ModificarFechaProgramadaDeliveryHandler
{
    private readonly IPedidoDeliveryRepository _repo;

    public ModificarFechaProgramadaDeliveryHandler(IPedidoDeliveryRepository repo)
        => _repo = repo;

    /// <summary>
    /// Modifica la fecha programada de entrega.
    /// BR-DEL-014: no se puede modificar fecha de un pedido ya entregado.
    /// </summary>
    public async Task<Result> HandleAsync(ModificarFechaProgramadaDeliveryCommand cmd, CancellationToken ct = default)
    {
        // BR-DEL-014: verificar que no fue entregado
        var yaEntregado = await _repo.EstaEntregadoAsync(cmd.CodigoPedido, ct);
        if (yaEntregado)
            return Result.Fail("No se puede modificar la fecha de un pedido ya Entregado.", "DELIVERY_YA_ENTREGADO");

        if (cmd.NuevaFecha == default)
            return Result.Fail("La nueva fecha programada es obligatoria.", "DELIVERY_FECHA_REQUERIDA");

        await _repo.ModificarFechaProgramadaAsync(cmd.CodigoPedido, cmd.NuevaFecha, ct);
        return Result.Ok();
    }
}

/// <summary>
/// Query para obtener pedidos pendientes de delivery desde la vista <c>vDespachador</c>.
/// <para>
/// Legacy: <c>frmPedidoDelivery.frm Form_Load</c>.
/// SQL: <c>select * from vDespachador where tTipoPedido='02' and tEstadoPedido='02' and isnull(fLlegada,0)=0</c>.
/// </para>
/// Regla: BR-DEL-009.
/// </summary>
public sealed record ObtenerPedidosSeguimientoDeliveryQuery;

/// <summary>Handler de <see cref="ObtenerPedidosSeguimientoDeliveryQuery"/>.</summary>
public sealed class ObtenerPedidosSeguimientoDeliveryHandler
{
    private readonly IPedidoDeliveryRepository _repo;

    public ObtenerPedidosSeguimientoDeliveryHandler(IPedidoDeliveryRepository repo)
        => _repo = repo;

    public async Task<Result<IEnumerable<PedidoDelivery>>> HandleAsync(
        ObtenerPedidosSeguimientoDeliveryQuery _, CancellationToken ct = default)
    {
        var hoy = DateTime.Today;
        var pedidos = await _repo.ObtenerParaDespachadorAsync(hoy, hoy.AddDays(1).AddSeconds(-1), ct);
        return Result.Ok(pedidos);
    }

    /// <summary>
    /// Query para obtener pedidos delivery entregados desde <c>vDespachador</c>.
    /// <para>
    /// Legacy: <c>frmPedidoDeliveryNo.frm</c> (<c>Form_Load</c>, <c>cmdProcesa_Click</c>).
    /// SQL: <c>tTipoPedido='02' and tEstadoPedido='02' and isnull(fLlegada,0)<>0</c>.
    /// </para>
    /// Regla: BR-DEL-038.
    /// </summary>
    public sealed record ObtenerPedidosSeguimientoDeliveryEntregadosQuery(DateTime FechaInicio, DateTime FechaFin);

    /// <summary>Handler de <see cref="ObtenerPedidosSeguimientoDeliveryEntregadosQuery"/>.</summary>
    public sealed class ObtenerPedidosSeguimientoDeliveryEntregadosHandler
    {
        private readonly IPedidoDeliveryRepository _repo;

        public ObtenerPedidosSeguimientoDeliveryEntregadosHandler(IPedidoDeliveryRepository repo)
            => _repo = repo;

        public async Task<Result<IReadOnlyList<PedidoDeliverySeguimiento>>> HandleAsync(
            ObtenerPedidosSeguimientoDeliveryEntregadosQuery query,
            CancellationToken ct = default)
        {
            var inicio = query.FechaInicio.Date;
            var fin = query.FechaFin.Date.AddDays(1).AddTicks(-1);
            if (fin < inicio)
                return Result.Fail<IReadOnlyList<PedidoDeliverySeguimiento>>("Rango de fechas inválido.", "DELIVERY_RANGO_INVALIDO");

            var rows = await _repo.ObtenerSeguimientoEntregadosAsync(inicio, fin, ct);
            return Result.Ok(rows);
        }
    }
}

// ──────────────────────────────────────────────────────────────────────────────
// POS-FUNC-037 — Mantenimiento de Clientes Delivery (frmClienteDelivery.frm)
// ──────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Obtiene el listado principal de clientes delivery para mantenimiento.
/// Legacy: <c>frmClienteDelivery.frm</c> Form_Load — <c>SELECT … FROM vDelivery</c>.
/// Reglas: BR-DEL-037.
/// </summary>
public sealed record ObtenerClientesDeliveryListadoQuery;

/// <summary>Handler de <see cref="ObtenerClientesDeliveryListadoQuery"/>.</summary>
public sealed class ObtenerClientesDeliveryListadoHandler
{
    private readonly IClienteDeliveryReadRepository _repo;

    public ObtenerClientesDeliveryListadoHandler(IClienteDeliveryReadRepository repo)
        => _repo = repo;

    public async Task<Result<IReadOnlyList<ClienteDeliveryListadoItem>>> HandleAsync(
        ObtenerClientesDeliveryListadoQuery _, CancellationToken ct = default)
    {
        var items = await _repo.ListarMantenimientoAsync(ct);
        return Result.Ok(items);
    }
}

// ──────────────────────────────────────────────────────────────────────────────
// POS-FUNC-036 — Búsqueda de Clientes Delivery (frmBusquedaDelivery.frm)
// ──────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Obtiene todos los clientes delivery activos para la grilla de búsqueda.
/// Legacy: <c>frmBusquedaDelivery.frm</c> Form_Load — TDELIVERY LEFT JOIN vZona WHERE lActivo=1.
/// Regla BR-DEL-036.
/// </summary>
public sealed record ObtenerClientesDeliveryBusquedaQuery;

/// <summary>Handler de <see cref="ObtenerClientesDeliveryBusquedaQuery"/>.</summary>
public sealed class ObtenerClientesDeliveryBusquedaHandler
{
    private readonly IClienteDeliveryReadRepository _repo;

    public ObtenerClientesDeliveryBusquedaHandler(IClienteDeliveryReadRepository repo)
        => _repo = repo;

    public async Task<Result<IReadOnlyList<ClienteDeliveryBusquedaItem>>> HandleAsync(
        ObtenerClientesDeliveryBusquedaQuery _, CancellationToken ct = default)
    {
        var items = await _repo.ListarActivosConZonaAsync(ct);
        return Result.Ok(items);
    }
}

/// <summary>
/// Obtiene el detalle de un cliente delivery para el panel lateral.
/// Legacy: Sub Asigna() de <c>frmBusquedaDelivery.frm</c>.
/// Regla BR-DEL-036.
/// </summary>
public sealed record ObtenerDetalleClienteDeliveryQuery(string CodigoDelivery);

/// <summary>Handler de <see cref="ObtenerDetalleClienteDeliveryQuery"/>.</summary>
public sealed class ObtenerDetalleClienteDeliveryHandler
{
    private readonly IClienteDeliveryReadRepository _repo;

    public ObtenerDetalleClienteDeliveryHandler(IClienteDeliveryReadRepository repo)
        => _repo = repo;

    public async Task<Result<ClienteDeliveryDetalleBusqueda?>> HandleAsync(
        ObtenerDetalleClienteDeliveryQuery query, CancellationToken ct = default)
    {
        var detalle = await _repo.ObtenerDetalleAsync(query.CodigoDelivery, ct);
        return Result.Ok(detalle);
    }
}

/// <summary>
/// Obtiene estadísticas históricas de un cliente delivery (panel "Otros Datos").
/// Legacy: cmdOpcion(3) de <c>frmBusquedaDelivery.frm</c>.
/// Regla BR-DEL-036.
/// </summary>
public sealed record ObtenerEstadisticasClienteDeliveryQuery(
    string CodigoDelivery,
    int DiasHistorico = 30);

/// <summary>Handler de <see cref="ObtenerEstadisticasClienteDeliveryQuery"/>.</summary>
public sealed class ObtenerEstadisticasClienteDeliveryHandler
{
    private readonly IClienteDeliveryReadRepository _repo;

    public ObtenerEstadisticasClienteDeliveryHandler(IClienteDeliveryReadRepository repo)
        => _repo = repo;

    public async Task<Result<EstadisticasClienteDelivery>> HandleAsync(
        ObtenerEstadisticasClienteDeliveryQuery query, CancellationToken ct = default)
    {
        var stats = await _repo.ObtenerEstadisticasAsync(query.CodigoDelivery, query.DiasHistorico, ct);
        return Result.Ok(stats);
    }
}

/// <summary>
/// Obtiene las tiendas/sucursales activas de un cliente delivery.
/// Legacy: Tienda_Click de <c>frmBusquedaDelivery.frm</c> — vTienda WHERE lActivo=1 AND tCodigoDelivery=…
/// </summary>
public sealed record ObtenerTiendasClienteDeliveryQuery(string CodigoDelivery);

/// <summary>Handler de <see cref="ObtenerTiendasClienteDeliveryQuery"/>.</summary>
public sealed class ObtenerTiendasClienteDeliveryHandler
{
    private readonly IClienteDeliveryReadRepository _repo;

    public ObtenerTiendasClienteDeliveryHandler(IClienteDeliveryReadRepository repo)
        => _repo = repo;

    public async Task<Result<IReadOnlyList<TiendaDeliveryItem>>> HandleAsync(
        ObtenerTiendasClienteDeliveryQuery query, CancellationToken ct = default)
    {
        var tiendas = await _repo.ObtenerTiendasAsync(query.CodigoDelivery, ct);
        return Result.Ok(tiendas);
    }
}
