using Inforest.Application.Caja;
using Inforest.Application.Turno;
using Inforest.Domain.Entities.Caja;

namespace Inforest.Desktop.Caja;

/// <summary>
/// Pantalla de Cierre de Turno con desglose multi-moneda.
/// Legacy: frmLiquidacionDetalle.frm — "Cierre de Turno".
/// Reglas: BR-CAJA-001 (lObligaCierre → supervisor), BR-CAJA-002 (lActivaConsultaDescargo),
///         BR-CAJA-003 (desglose de pagos), BR-CAJA-004 (UPDATE MTURNO).
/// </summary>
public sealed class FrmLiquidacionCierre : Form
{
    private readonly CerrarTurnoHandler? _cerrarHandler;
    private readonly ObtenerResumenTurnoHandler? _resumenHandler;
    private readonly string _codigoTurno;
    private readonly string _codigoCaja;

    // ── campos de desglose ──
    private readonly TextBox _txtEfectivoMN = new() { Text = "0.00", Width = 120 };
    private readonly TextBox _txtEfectivoME = new() { Text = "0.00", Width = 120 };
    private readonly TextBox _txtChequesMN = new() { Text = "0.00", Width = 120 };
    private readonly TextBox _txtChequesME = new() { Text = "0.00", Width = 120 };
    private readonly TextBox _txtPuntosMN = new() { Text = "0.00", Width = 120 };
    private readonly TextBox _txtTarjetas = new() { Text = "0.00", Width = 120 };
    private readonly TextBox _txtOtros = new() { Text = "0.00", Width = 120 };
    private readonly TextBox _txtRetiroMN = new() { Text = "0.00", Width = 120 };
    private readonly TextBox _txtRetiroME = new() { Text = "0.00", Width = 120 };

    private readonly Label _lblResumen = new() { AutoSize = true };
    private readonly Label _lblTotal = new() { AutoSize = true, Font = new Font("Arial", 10, FontStyle.Bold) };
    private readonly Label _lblDiferencia = new() { AutoSize = true };

    private decimal _totalVentasTurno;

    public FrmLiquidacionCierre(
        string codigoTurno,
        string codigoCaja,
        CerrarTurnoHandler? cerrarHandler = null,
        ObtenerResumenTurnoHandler? resumenHandler = null)
    {
        _codigoTurno = codigoTurno;
        _codigoCaja = codigoCaja;
        _cerrarHandler = cerrarHandler;
        _resumenHandler = resumenHandler;

        Text = "Cierre de Turno";
        Width = 520;
        Height = 560;
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        BuildLayout();

        Load += async (_, _) => await CargarResumenAsync();

        // Recalcular diferencia en cada cambio
        foreach (var txt in new[] { _txtEfectivoMN, _txtEfectivoME, _txtChequesMN, _txtChequesME,
                                     _txtPuntosMN, _txtTarjetas, _txtOtros, _txtRetiroMN, _txtRetiroME })
            txt.TextChanged += (_, _) => ActualizarDiferencia();
    }

    private void BuildLayout()
    {
        var table = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(12),
            RowCount = 14,
            ColumnCount = 2,
            AutoSize = true
        };
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));

        int row = 0;
        AddRow(table, row++, "Turno:", _lblResumen);
        AddRow(table, row++, "Efectivo MN:", _txtEfectivoMN);
        AddRow(table, row++, "Efectivo ME:", _txtEfectivoME);
        AddRow(table, row++, "Cheques MN:", _txtChequesMN);
        AddRow(table, row++, "Cheques ME:", _txtChequesME);
        AddRow(table, row++, "Puntos:", _txtPuntosMN);
        AddRow(table, row++, "Tarjetas (total):", _txtTarjetas);
        AddRow(table, row++, "Otros pagos:", _txtOtros);
        AddRow(table, row++, "Retiro MN:", _txtRetiroMN);
        AddRow(table, row++, "Retiro ME:", _txtRetiroME);
        AddRow(table, row++, "Total ingresado:", _lblTotal);
        AddRow(table, row++, "Diferencia vs ventas:", _lblDiferencia);

        var btnCerrar = new Button { Text = "Cerrar turno", Width = 120, Height = 34 };
        var btnCancelar = new Button { Text = "Cancelar", Width = 90, Height = 34 };

        btnCerrar.Click += async (_, _) => await CerrarTurnoAsync();
        btnCancelar.Click += (_, _) => Close();

        var pnlBotones = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.LeftToRight,
            AutoSize = true,
            Margin = new Padding(0, 8, 0, 0)
        };
        pnlBotones.Controls.Add(btnCerrar);
        pnlBotones.Controls.Add(btnCancelar);

        table.Controls.Add(pnlBotones, 0, row);
        table.SetColumnSpan(pnlBotones, 2);

        Controls.Add(table);
    }

    private static void AddRow(TableLayoutPanel table, int row, string label, Control control)
    {
        table.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Top, Margin = new Padding(0, 6, 6, 0) }, 0, row);
        table.Controls.Add(control, 1, row);
    }

    private async Task CargarResumenAsync()
    {
        _lblResumen.Text = _codigoTurno;
        if (_resumenHandler is null) return;

        try
        {
            var result = await _resumenHandler.HandleAsync(new ObtenerResumenTurnoQuery(_codigoTurno));
            if (result.EsExitoso && result.Valor is not null)
            {
                _totalVentasTurno = result.Valor.TotalVentas;
                ActualizarDiferencia();
            }
        }
        catch { /* ignorar errores de carga inicial */ }
    }

    private void ActualizarDiferencia()
    {
        var total = ParseDecimal(_txtEfectivoMN) + ParseDecimal(_txtEfectivoME)
                  + ParseDecimal(_txtChequesMN) + ParseDecimal(_txtChequesME)
                  + ParseDecimal(_txtPuntosMN) + ParseDecimal(_txtTarjetas)
                  + ParseDecimal(_txtOtros)
                  - ParseDecimal(_txtRetiroMN) - ParseDecimal(_txtRetiroME);

        _lblTotal.Text = total.ToString("0.00");
        _lblDiferencia.Text = (total - _totalVentasTurno).ToString("0.00");
    }

    private static decimal ParseDecimal(TextBox txt)
        => decimal.TryParse(txt.Text, out var v) ? v : 0m;

    private async Task CerrarTurnoAsync()
    {
        if (_cerrarHandler is null)
        {
            MessageBox.Show("Handler de cierre no configurado.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // Confirmación final antes de cerrar — Legacy: MsgBox("Seguro de Cerrar el Turno " & sTurno & "?")
        var confirmar = MessageBox.Show(
            $"¿Seguro que desea cerrar el turno {_codigoTurno}?",
            Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
        if (confirmar == DialogResult.Cancel) return;

        var efecMN = ParseDecimal(_txtEfectivoMN);
        var retiroMN = ParseDecimal(_txtRetiroMN);
        if (retiroMN > efecMN)
        {
            // BR-CAJA-003: no se puede retirar más del efectivo disponible
            MessageBox.Show("No se puede retirar más del Efectivo MN disponible.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var efecME = ParseDecimal(_txtEfectivoME);
        var retiroME = ParseDecimal(_txtRetiroME);
        if (retiroME > efecME)
        {
            MessageBox.Show("No se puede retirar más del Efectivo ME disponible.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var tarjetas = ParseDecimal(_txtTarjetas);
        var breakdown = new CierreTurnoBreakdown
        {
            EfectivoMN = efecMN,
            EfectivoME = efecME,
            ChequesMN = ParseDecimal(_txtChequesMN),
            ChequesME = ParseDecimal(_txtChequesME),
            PuntosMN = ParseDecimal(_txtPuntosMN),
            Tarjetas = [tarjetas, 0, 0, 0, 0, 0, 0, 0],
            OtrosMN = ParseDecimal(_txtOtros),
            RetiroMN = retiroMN,
            RetiroME = retiroME
        };

        var montoFinal = breakdown.MontoFinalMN;

        var command = new CerrarTurnoCommand(
            CodigoTurno: _codigoTurno,
            CodigoCaja: _codigoCaja,
            MontoFinal: montoFinal,
            Breakdown: breakdown,
            CodigoUsuario: Environment.UserName);

        var result = await _cerrarHandler.HandleAsync(command);

        // BR-CAJA-001: lObligaCierre → pedir contraseña de supervisor y reintentar
        if (!result.EsExitoso && result.CodigoError == "REQUIERE_SUPERVISOR")
        {
            var clave = PromptSupervisor();
            if (clave is null) return;

            result = await _cerrarHandler.HandleAsync(command with { SupervisorAutorizado = true });
        }

        // BR-CAJA-002: lActivaConsultaDescargo → pedir confirmación.
        // Legacy: si el usuario dice "No", ofrece abrir frmDescargo (GAP-CAJA-002: frmDescargo aún no migrado).
        if (!result.EsExitoso && result.CodigoError == "REQUIERE_CONFIRMACION_DESCARGO")
        {
            var resp = MessageBox.Show(
                "Debe descargar ventas antes de cerrar el turno. ¿Realizó este proceso?",
                Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (resp == DialogResult.No)
            {
                // Legacy ofrece abrir frmDescargo aquí; pendiente de migración (GAP-CAJA-002).
                MessageBox.Show("Operación cancelada. Realice el descargo antes de cerrar.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            result = await _cerrarHandler.HandleAsync(command with { DescargoPendienteConfirmado = true });
        }

        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        MessageBox.Show("Turno cerrado correctamente.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        DialogResult = DialogResult.OK;
        Close();
    }

    /// <summary>
    /// Solicita contraseña de supervisor.
    /// Legacy: Supervisor("11") — verifica acceso código 11 en TACCESO.
    /// </summary>
    private static string? PromptSupervisor()
    {
        using var frm = new Form
        {
            Text = "Autorización de Supervisor",
            Width = 300,
            Height = 160,
            StartPosition = FormStartPosition.CenterParent,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false,
            MinimizeBox = false
        };
        var lbl = new Label { Left = 20, Top = 20, Text = "Ingrese contraseña de supervisor:", AutoSize = true };
        var txt = new TextBox { Left = 20, Top = 50, Width = 240, PasswordChar = '*' };
        var btnOk = new Button { Left = 60, Top = 85, Text = "Aceptar", Width = 80 };
        var btnCancel = new Button { Left = 160, Top = 85, Text = "Cancelar", Width = 80 };
        btnOk.DialogResult = DialogResult.OK;
        btnCancel.DialogResult = DialogResult.Cancel;
        frm.Controls.AddRange([lbl, txt, btnOk, btnCancel]);
        frm.AcceptButton = btnOk;
        frm.CancelButton = btnCancel;

        return frm.ShowDialog() == DialogResult.OK ? txt.Text : null;
    }
}
