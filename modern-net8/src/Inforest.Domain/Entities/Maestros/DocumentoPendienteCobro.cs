namespace Inforest.Domain.Entities.Maestros;

/// <summary>
/// Documento pendiente de cobro por cuenta corriente.
/// Legacy: vDocumentoGrilla WHERE tEstadoDocumento='03' / frmCuentaCobrar.frm.
/// BR-CTACTE-002: Los documentos en estado '03' son deuda ctacte pendiente de cobro.
/// </summary>
public sealed record DocumentoPendienteCobro(
    string Documento,
    string Caja,
    string? CodigoCliente,
    string? NombreCliente,
    string? CodigoClientePago,
    string? NombreClientePago,
    decimal Venta,
    DateTime Fecha,
    string Turno,
    string? Usuario,
    string? Vencimiento);
