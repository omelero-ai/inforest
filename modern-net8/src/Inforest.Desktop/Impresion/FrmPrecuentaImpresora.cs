using Inforest.Application.Impresion;
using Inforest.Application.Interfaces;

namespace Inforest.Desktop.Impresion;

/// <summary>
/// Formulario para seleccionar impresora e imprimir la pre-cuenta de un pedido.
/// Legacy: frmPrecuentaImpresora.frm. BR-008.
/// </summary>
public class FrmPrecuentaImpresora : Form
{
    private readonly ObtenerImpresorasPorCajaHandler? _impresorasHandler;
    private readonly ImprimirPrecuentaHandler? _imprimirHandler;
    private readonly string _codigoPedido;
    private readonly string _codigoCaja;
    private readonly ListBox _lstImpresoras;
    private readonly Label _lblEstado;

    public FrmPrecuentaImpresora(
        string codigoPedido,
        string codigoCaja,
        ObtenerImpresorasPorCajaHandler? impresorasHandler = null,
        ImprimirPrecuentaHandler? imprimirHandler = null)
    {
        _codigoPedido = codigoPedido;
        _codigoCaja = codigoCaja;
        _impresorasHandler = impresorasHandler;
        _imprimirHandler = imprimirHandler;

        Text = "Seleccionar Impresora — Pre-Cuenta";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterParent;
        ClientSize = new Size(320, 280);

        var lblTitle = new Label
        {
            Left = 10, Top = 10, Width = 300,
            Text = $"Pedido: {codigoPedido}",
            Font = new Font(Font, FontStyle.Bold)
        };

        var lblImpresoras = new Label { Left = 10, Top = 36, Width = 200, Text = "Seleccione la impresora:" };

        _lstImpresoras = new ListBox
        {
            Left = 10, Top = 56, Width = 290, Height = 140,
            SelectionMode = SelectionMode.One
        };
        _lstImpresoras.DoubleClick += async (_, _) => await ImprimirAsync();

        _lblEstado = new Label
        {
            Left = 10, Top = 206, Width = 290, Height = 28,
            ForeColor = Color.DarkRed
        };

        var btnImprimir = new Button { Left = 100, Top = 240, Width = 100, Text = "Imprimir" };
        btnImprimir.Click += async (_, _) => await ImprimirAsync();

        var btnCerrar = new Button
        {
            Left = 210, Top = 240, Width = 100, Text = "Cerrar",
            DialogResult = DialogResult.Cancel
        };
        btnCerrar.Click += (_, _) => Close();

        Controls.AddRange([lblTitle, lblImpresoras, _lstImpresoras, _lblEstado, btnImprimir, btnCerrar]);

        Load += async (_, _) => await CargarImpresorasAsync();
    }

    private async Task CargarImpresorasAsync()
    {
        _lstImpresoras.Items.Clear();

        if (_impresorasHandler is null)
        {
            _lblEstado.Text = "Sin impresoras (handler no configurado).";
            return;
        }

        var result = await _impresorasHandler.HandleAsync(
            new ObtenerImpresorasPorCajaQuery(_codigoCaja));

        if (!result.EsExitoso)
        {
            _lblEstado.Text = result.MensajeError ?? "Error al cargar impresoras.";
            return;
        }

        if (result.Valor!.Count == 0)
        {
            // Legacy: "No existe impresoras configuradas para esta caja"
            MessageBox.Show(
                "No existe impresoras configuradas para esta caja.",
                "Impresión", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            Close();
            return;
        }

        foreach (var imp in result.Valor)
            _lstImpresoras.Items.Add(imp);

        _lstImpresoras.DisplayMember = nameof(ImpresoraDisponible.Descripcion);
        _lstImpresoras.ValueMember = nameof(ImpresoraDisponible.CodigoImpresora);
        _lstImpresoras.SelectedIndex = 0;
    }

    private async Task ImprimirAsync()
    {
        if (_lstImpresoras.SelectedItem is not ImpresoraDisponible seleccionada)
        {
            _lblEstado.Text = "Seleccione una impresora.";
            return;
        }

        if (_imprimirHandler is null)
        {
            _lblEstado.Text = "Handler de impresión no configurado.";
            return;
        }

        _lblEstado.Text = "Imprimiendo...";
        var result = await _imprimirHandler.HandleAsync(
            new ImprimirPrecuentaCommand(_codigoPedido, seleccionada.CodigoImpresora));

        if (result.EsExitoso)
        {
            Close();
        }
        else
        {
            _lblEstado.Text = result.MensajeError ?? "Error al imprimir.";
        }
    }
}
