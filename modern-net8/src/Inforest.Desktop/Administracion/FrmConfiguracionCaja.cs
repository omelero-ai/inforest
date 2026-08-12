using System.Reflection;
using Inforest.Application.Configuracion;
using Inforest.Domain.Entities.Configuracion;

namespace Inforest.Desktop.Administracion;

/// <summary>
/// Legacy: <c>frmConfiguracionParametro.frm</c> + <c>TCAJA</c>.
/// </summary>
public class FrmConfiguracionCaja : Form
{
    private readonly ObtenerTodasCajasHandler _todasCajasHandler;
    private readonly ObtenerConfiguracionCajaHandler _configuracionCajaHandler;
    private readonly ComboBox _cmbCajas;
    private readonly FlowLayoutPanel _flagsPanel;
    private readonly DataGridView _gridResumen;
    private readonly Dictionary<string, CheckBox> _checks = new(StringComparer.OrdinalIgnoreCase);

    public FrmConfiguracionCaja(
        ObtenerTodasCajasHandler todasCajasHandler,
        ObtenerConfiguracionCajaHandler configuracionCajaHandler)
    {
        _todasCajasHandler = todasCajasHandler;
        _configuracionCajaHandler = configuracionCajaHandler;

        Text = "Configuración de Caja";
        Width = 1100;
        Height = 700;

        _cmbCajas = new ComboBox { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList, Height = 28 };
        _cmbCajas.SelectedIndexChanged += async (_, _) => await CargarCajaSeleccionadaAsync();

        _flagsPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Left,
            Width = 420,
            AutoScroll = true,
            Padding = new Padding(8),
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false
        };

        _gridResumen = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = true,
            AllowUserToAddRows = false,
            ReadOnly = true
        };

        var botones = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 44, Padding = new Padding(8) };
        botones.Controls.AddRange([
            new Button { Text = "Guardar", Width = 100 },
            new Button { Text = "Cancelar", Width = 100 }
        ]);
        botones.Controls[0].Click += (_, _) => MessageBox.Show("Cambios preparados para wiring de actualización en TCAJA.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        botones.Controls[1].Click += async (_, _) => await CargarCajaSeleccionadaAsync();

        Controls.Add(_gridResumen);
        Controls.Add(_flagsPanel);
        Controls.Add(botones);
        Controls.Add(_cmbCajas);

        ConstruirChecks();
        Load += async (_, _) => await CargarCajasAsync();
    }

    private void ConstruirChecks()
    {
        foreach (var propiedad in typeof(ConfiguracionCaja).GetProperties().Where(static p => p.PropertyType == typeof(bool)))
        {
            var check = new CheckBox { Text = propiedad.Name, AutoSize = true };
            _checks[propiedad.Name] = check;
            _flagsPanel.Controls.Add(check);
        }
    }

    private async Task CargarCajasAsync()
    {
        var result = await _todasCajasHandler.HandleAsync(new ObtenerTodasCajasQuery());
        if (!result.EsExitoso)
            return;

        _cmbCajas.DataSource = result.Valor?.ToList();
        _cmbCajas.DisplayMember = "tDescripcion";
        _cmbCajas.ValueMember = "tCaja";
        if (_cmbCajas.Items.Count > 0)
            _cmbCajas.SelectedIndex = 0;
    }

    private async Task CargarCajaSeleccionadaAsync()
    {
        if (_cmbCajas.SelectedValue is not string codigoCaja)
            return;

        var result = await _configuracionCajaHandler.HandleAsync(new ObtenerConfiguracionCajaQuery(codigoCaja));
        if (!result.EsExitoso || result.Valor is null)
            return;

        var caja = result.Valor;
        foreach (var propiedad in typeof(ConfiguracionCaja).GetProperties())
        {
            if (propiedad.PropertyType == typeof(bool) && _checks.TryGetValue(propiedad.Name, out var check))
                check.Checked = (bool)(propiedad.GetValue(caja) ?? false);
        }

        _gridResumen.DataSource = typeof(ConfiguracionCaja)
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(static p => p.PropertyType != typeof(bool))
            .Select(propiedad => new CajaDetalleRow(propiedad.Name, propiedad.GetValue(caja)?.ToString() ?? string.Empty))
            .ToList();
    }

    private sealed record CajaDetalleRow(string Propiedad, string Valor);
}
