using Inforest.Application.Interfaces;
using Inforest.Domain.Common;
using Inforest.Domain.Entities.Ventas;
using Inforest.Domain.Exceptions;
using Inforest.Domain.Repositories;

namespace Inforest.Application.Ventas;

public sealed record EmitirDocumentoCommand(
    string CodigoPedido,
    string TipoDocumento,
    string? CodigoCliente,
    decimal Propina,
    decimal Descuento);

public sealed class EmitirDocumentoHandler
{
    private readonly IPedidoRepository _pedidoRepository;
    private readonly IDocumentoRepository _documentoRepository;
    private readonly ISessionService _sessionService;

    public EmitirDocumentoHandler(
        IPedidoRepository pedidoRepository,
        IDocumentoRepository documentoRepository,
        ISessionService sessionService)
    {
        _pedidoRepository = pedidoRepository;
        _documentoRepository = documentoRepository;
        _sessionService = sessionService;
    }

    public async Task<Result<Documento>> HandleAsync(EmitirDocumentoCommand command, CancellationToken ct = default)
    {
        try
        {
            var pedido = await _pedidoRepository.ObtenerPorCodigoAsync(command.CodigoPedido, ct);
            if (pedido is null)
                return Result.Fail<Documento>("No se encontró el pedido origen.", "VENTA_PEDIDO_NO_ENCONTRADO");

            if (pedido.Detalles.Count == 0)
                return Result.Fail<Documento>("El pedido no tiene detalles para documentar.", "VENTA_PEDIDO_SIN_DETALLES");

            var codigoCaja = _sessionService.SesionActual?.CodigoCaja;
            if (string.IsNullOrWhiteSpace(codigoCaja))
                return Result.Fail<Documento>("No existe una sesión activa con caja asociada.", "VENTA_CAJA_NO_DISPONIBLE");

            var codigoDocumento = $"{command.TipoDocumento.Trim()}-{command.CodigoPedido.Trim()}";
            if (await _documentoRepository.ObtenerPorCodigoAsync(codigoDocumento, ct) is not null)
                return Result.Fail<Documento>("Ya existe un documento emitido para el pedido indicado.", "VENTA_DOCUMENTO_DUPLICADO");

            var neto = pedido.Detalles.Sum(d => d.PrecioNeto * d.Cantidad);
            var impuesto1 = pedido.Detalles.Sum(d => d.Impuesto1 * d.Cantidad);
            var impuesto2 = pedido.Detalles.Sum(d => d.Impuesto2 * d.Cantidad);
            var impuesto3 = pedido.Detalles.Sum(d => d.Impuesto3 * d.Cantidad);

            var documento = Documento.Emitir(
                codigoDocumento,
                command.TipoDocumento,
                command.CodigoCliente,
                neto,
                impuesto1,
                impuesto2,
                impuesto3,
                command.Propina,
                command.Descuento,
                pedido.CodigoTurno ?? string.Empty,
                codigoCaja);

            var detalles = pedido.Detalles
                .Select(d => DetalleDocumento.Crear(
                    codigoDocumento,
                    d.Item,
                    pedido.CodigoPedido,
                    d.CodigoProducto,
                    d.Cantidad,
                    d.PrecioNeto,
                    d.PrecioVenta,
                    d.Impuesto1,
                    d.Impuesto2,
                    d.Impuesto3,
                    d.Descuento,
                    d.Area))
                .ToList();

            var insertado = await _documentoRepository.InsertarAsync(documento, detalles, ct);
            return insertado
                ? Result.Ok(documento)
                : Result.Fail<Documento>("No se pudo registrar el documento emitido.", "VENTA_DOCUMENTO_NO_INSERTADO");
        }
        catch (DomainException ex)
        {
            return Result.Fail<Documento>(ex.Message, ex.Code);
        }
    }
}

public sealed record ObtenerDocumentoQuery(string CodigoDocumento);

public sealed class ObtenerDocumentoHandler
{
    private readonly IDocumentoRepository _documentoRepository;

    public ObtenerDocumentoHandler(IDocumentoRepository documentoRepository)
        => _documentoRepository = documentoRepository;

    public async Task<Result<Documento>> HandleAsync(ObtenerDocumentoQuery query, CancellationToken ct = default)
    {
        var documento = await _documentoRepository.ObtenerPorCodigoAsync(query.CodigoDocumento, ct);
        return documento is null
            ? Result.Fail<Documento>("No se encontró el documento solicitado.", "VENTA_DOCUMENTO_NO_ENCONTRADO")
            : Result.Ok(documento);
    }
}

public sealed record AnularDocumentoCommand(string CodigoDocumento, string Usuario, string Motivo);

public sealed class AnularDocumentoHandler
{
    private readonly IDocumentoRepository _documentoRepository;

    public AnularDocumentoHandler(IDocumentoRepository documentoRepository)
        => _documentoRepository = documentoRepository;

    public async Task<Result> HandleAsync(AnularDocumentoCommand command, CancellationToken ct = default)
    {
        try
        {
            var documento = await _documentoRepository.ObtenerPorCodigoAsync(command.CodigoDocumento, ct);
            if (documento is null)
                return Result.Fail("No se encontró el documento a anular.", "VENTA_DOCUMENTO_NO_ENCONTRADO");

            documento.Anular(command.Usuario, command.Motivo);
            var anulado = await _documentoRepository.AnularAsync(command.CodigoDocumento, command.Usuario, command.Motivo, ct);

            return anulado
                ? Result.Ok()
                : Result.Fail("No se pudo anular el documento.", "VENTA_DOCUMENTO_NO_ANULADO");
        }
        catch (DomainException ex)
        {
            return Result.Fail(ex.Message, ex.Code);
        }
    }
}
