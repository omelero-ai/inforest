namespace Inforest.Domain.Entities.Hardware;

/// <summary>
/// Terminal de cobro PinPad asignado a una caja.
/// Legacy: frmPagoPinPad.frm, TCAJATERMINAL, TTERMINAL.
/// Regla BR-PERIPH-001.
/// </summary>
public sealed record TerminalPinPad(
    string CodigoTerminal,
    string Descripcion,
    string CodigoIntegracion);
