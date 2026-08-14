using System.Reflection;
using Inforest.Application.Configuracion;
using Inforest.Domain.Entities.Configuracion;

namespace Inforest.Desktop.Administracion;

/// <summary>
/// Legacy: <c>frmParametro.frm</c>.
/// </summary>
public class FrmParametro : Form
{
    private readonly ObtenerConfiguracionSistemaHandler _handler;
    private readonly BindingSource _bindingSource = new();
    private readonly DataGridView _grid;

    public FrmParametro(ObtenerConfiguracionSistemaHandler handler)
    {
        _handler = handler;
        Text = "Parámetros del Sistema";

        _grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = true,
            AllowUserToAddRows = false,
            DataSource = _bindingSource
        };

        var botones = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 44, Padding = new Padding(8) };
        var btnGuardar = new Button { Text = "Guardar", Width = 100 };
        var btnCancelar = new Button { Text = "Cancelar", Width = 100 };
        btnGuardar.Click += (_, _) => MessageBox.Show("Edición local lista para wiring con TPARAMETRO.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        btnCancelar.Click += async (_, _) => await CargarAsync();
        botones.Controls.AddRange([btnGuardar, btnCancelar]);

        Controls.Add(_grid);
        Controls.Add(botones);
        Load += async (_, _) => await CargarAsync();
    }

    private async Task CargarAsync()
    {
        var result = await _handler.HandleAsync(new ObtenerConfiguracionSistemaQuery());
        var filas = new List<ParametroRow>();
        if (result.EsExitoso && result.Valor is ConfiguracionSistema configuracion)
        {
            foreach (var property in typeof(ConfiguracionSistema).GetProperties(BindingFlags.Public | BindingFlags.Instance))
                filas.Add(new ParametroRow(property.Name, property.GetValue(configuracion)?.ToString() ?? string.Empty));
        }

        _bindingSource.DataSource = filas;
    }

    private sealed record ParametroRow(string Clave, string Valor);
}
