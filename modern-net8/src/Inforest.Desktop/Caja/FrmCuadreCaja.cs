using System.ComponentModel;
using Inforest.Application.Caja;
using Inforest.Application.Interfaces;
using Inforest.Application.Turno;

namespace Inforest.Desktop.Caja;

/// <summary>
/// Pantalla de cuadre y cierre de caja.
/// Legacy: frmCaja.frm. BR-007.
/// </summary>
public class FrmCuadreCaja : Form
{
    private readonly ISessionService? _sessionService;
    private readonly ObtenerTurnoActualHandler? _turnoHandler;
    private readonly ObtenerResumenTurnoHandler? _resumenHandler;
    private readonly CerrarTurnoHandler? _cerrarHandler;

    private readonly BindingList<CuadreRow> _rows = [];
    private readonly TextBox _txtMontoCierre;
    private readonly Label _lblDiferencia;
    private readonly Label _lblTurno;
    private string? _codigoTurnoActual;

    public FrmCuadreCaja(
        ISessionService? sessionService = null,
        ObtenerTurnoActualHandler? turnoHandler = null,
        ObtenerResumenTurnoHandler? resumenHandler = null,
        CerrarTurnoHandler? cerrarHandler = null)
    {
        _sessionService = sessionService;
        _turnoHandler = turnoHandler;
        _resumenHandler = resumenHandler;
        _cerrarHandler = cerrarHandler;

        _txtMontoCierre = new TextBox { Width = 140, Text = "0.00" };
        _lblDiferencia = new Label { AutoSize = true };
        _lblTurno = new Label { AutoSize = true, Text = "Turno: cargando..." };

        Text = "Cuadre de Caja";
        Width = 700;
        Height = 520;
        StartPosition = FormStartPosition.CenterParent;

        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = true,
            ReadOnly = true,
            DataSource = _rows
        };

        var panel = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 90,
            Padding = new Padding(12),
            FlowDirection = FlowDirection.LeftToRight
        };

        var btnCerrarTurno = new Button { Text = "Cerrar turno", Width = 110 };
        var btnCancelar = new Button { Text = "Cancelar", Width = 90 };
        btnCancelar.Click += (_, _) => Close();
        btnCerrarTurno.Click += async (_, _) => await CerrarTurnoAsync();
        _txtMontoCierre.TextChanged += (_, _) => ActualizarDiferencia();

        panel.Controls.Add(_lblTurno);
        panel.Controls.Add(new Label { Text = "Monto cierre:", AutoSize = true, Margin = new Padding(0, 8, 6, 0) });
        panel.Controls.Add(_txtMontoCierre);
        panel.Controls.Add(new Label { Text = "Diferencia:", AutoSize = true, Margin = new Padding(18, 8, 6, 0) });
        panel.Controls.Add(_lblDiferencia);
        panel.Controls.Add(btnCerrarTurno);
        panel.Controls.Add(btnCancelar);

        Controls.Add(grid);
        Controls.Add(panel);

        Load += async (_, _) => await CargarResumenAsync();
        ActualizarDiferencia();
    }

    private async Task CargarResumenAsync()
    {
        if (_turnoHandler is null || _sessionService is null) return;

        var codigoCaja = _sessionService.SesionActual?.CodigoCaja;
        if (string.IsNullOrWhiteSpace(codigoCaja)) return;

        try
        {
            var turnoResult = await _turnoHandler.HandleAsync(new ObtenerTurnoActualQuery(codigoCaja));
            if (!turnoResult.EsExitoso || turnoResult.Valor is null)
            {
                _lblTurno.Text = "Sin turno abierto";
                return;
            }

            _codigoTurnoActual = turnoResult.Valor.CodigoTurno;
            _lblTurno.Text = $"Turno: {_codigoTurnoActual}";

            if (_resumenHandler is not null)
            {
                var resumen = await _resumenHandler.HandleAsync(new ObtenerResumenTurnoQuery(_codigoTurnoActual));
                if (resumen.EsExitoso && resumen.Valor is not null)
                {
                    _rows.Clear();
                    _rows.Add(new CuadreRow("Total Ventas", resumen.Valor.TotalVentas));
                    _rows.Add(new CuadreRow("Total Pagos", resumen.Valor.TotalPagos));
                    _rows.Add(new CuadreRow("Total Propinas", resumen.Valor.TotalPropinas));
                    _rows.Add(new CuadreRow($"Documentos Cobrados ({resumen.Valor.DocumentosCobrados})", resumen.Valor.TotalPagos));
                }
            }
        }
        catch
        {
            _lblTurno.Text = "Turno: error al cargar";
        }

        ActualizarDiferencia();
    }

    private async Task CerrarTurnoAsync()
    {
        if (_cerrarHandler is null || string.IsNullOrWhiteSpace(_codigoTurnoActual))
        {
            MessageBox.Show("No hay turno activo para cerrar.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!decimal.TryParse(_txtMontoCierre.Text, out var montoFinal))
        {
            MessageBox.Show("Ingrese un monto de cierre válido.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var result = await _cerrarHandler.HandleAsync(
            new CerrarTurnoCommand(
                _codigoTurnoActual,
                _sessionService?.SesionActual?.CodigoCaja ?? string.Empty,
                montoFinal,
                CodigoUsuario: _sessionService?.SesionActual?.CodigoUsuario ?? Environment.UserName));
        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        MessageBox.Show("Turno cerrado correctamente.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        DialogResult = DialogResult.OK;
        Close();
    }

    private void ActualizarDiferencia()
    {
        _ = decimal.TryParse(_txtMontoCierre.Text, out var cierre);
        var esperado = _rows.Sum(x => x.MontoEsperado);
        _lblDiferencia.Text = $"{cierre - esperado:0.00}";
    }

    private sealed record CuadreRow(string Concepto, decimal MontoEsperado);
}
