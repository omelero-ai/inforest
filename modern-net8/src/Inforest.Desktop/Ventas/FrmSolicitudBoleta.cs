using Inforest.Application.Maestros;
using Inforest.Application.Ventas;
using Inforest.Domain.Entities.Maestros;

namespace Inforest.Desktop.Ventas;

/// <summary>
/// Diálogo para decidir si una boleta se emite con datos o sin datos.
/// Legacy: frmSolicitudBoleta.frm.
/// Reglas BR-BOLETA-001/002/003.
/// </summary>
public sealed class FrmSolicitudBoleta : Form
{
    private readonly ObtenerClienteGeneralBoletaHandler _clienteGeneralHandler;
    private readonly ObtenerClientesActivosHandler _clientesHandler;

    public FrmSolicitudBoleta(
        ObtenerClienteGeneralBoletaHandler clienteGeneralHandler,
        ObtenerClientesActivosHandler clientesHandler)
    {
        _clienteGeneralHandler = clienteGeneralHandler;
        _clientesHandler = clientesHandler;

        Text = "Emitir Boleta";
        Width = 500;
        Height = 260;
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        BackColor = Color.FromArgb(245, 246, 247);

        var titulo = new Label
        {
            Dock = DockStyle.Top,
            Height = 72,
            Text = "¿Emitir boleta con datos del cliente?",
            TextAlign = ContentAlignment.MiddleCenter,
            Font = new Font("Segoe UI", 13, FontStyle.Bold),
            ForeColor = Color.FromArgb(35, 35, 35)
        };

        var btnConDatos = CrearBoton("Con datos", Color.FromArgb(0, 122, 204));
        var btnSinDatos = CrearBoton("Sin datos", Color.FromArgb(108, 117, 125));
        var btnCancelar = CrearBoton("Cancelar", Color.FromArgb(220, 53, 69));

        btnConDatos.Click += async (_, _) => await SeleccionarClienteAsync();
        btnSinDatos.Click += async (_, _) => await UsarClienteGeneralAsync();
        btnCancelar.Click += (_, _) => { DialogResult = DialogResult.Cancel; Close(); };

        var botones = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 3,
            RowCount = 1,
            Padding = new Padding(24, 12, 24, 24)
        };
        botones.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
        botones.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
        botones.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33f));
        botones.Controls.Add(btnConDatos, 0, 0);
        botones.Controls.Add(btnSinDatos, 1, 0);
        botones.Controls.Add(btnCancelar, 2, 0);

        Controls.Add(botones);
        Controls.Add(titulo);
    }

    public Cliente? ClienteSeleccionado { get; private set; }

    private async Task SeleccionarClienteAsync()
    {
        using var frm = new FrmSeleccionClienteBoleta(_clientesHandler);
        if (frm.ShowDialog(this) != DialogResult.OK || frm.ClienteSeleccionado is null)
            return;

        ClienteSeleccionado = frm.ClienteSeleccionado;
        DialogResult = DialogResult.OK;
        Close();
    }

    private async Task UsarClienteGeneralAsync()
    {
        var result = await _clienteGeneralHandler.HandleAsync(new ObtenerClienteGeneralBoletaQuery());
        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        ClienteSeleccionado = result.Valor;
        DialogResult = DialogResult.OK;
        Close();
    }

    private static Button CrearBoton(string texto, Color color)
    {
        var boton = new Button
        {
            Dock = DockStyle.Fill,
            Height = 80,
            Margin = new Padding(8),
            Text = texto,
            BackColor = color,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        boton.FlatAppearance.BorderSize = 0;
        return boton;
    }
}
