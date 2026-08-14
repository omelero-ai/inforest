using Inforest.Application.Ventas;
using Inforest.Domain.Entities.Ventas;

namespace Inforest.Desktop.Ventas;

/// <summary>
/// Vista de documento emitido.
/// Legacy: frmDocumento.frm.
/// </summary>
public class FrmDocumento : Form
{
    private readonly Documento _documento;
    private readonly AnularDocumentoHandler? _anularHandler;
    private readonly Label _lblEstado;

    public FrmDocumento(Documento documento, AnularDocumentoHandler? anularHandler = null)
    {
        _documento = documento;
        _anularHandler = anularHandler;

        Text = "Documento Emitido";
        Width = 520;
        Height = 320;
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(16),
            ColumnCount = 2,
            RowCount = 7
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 65));

        _lblEstado = new Label { Text = MapEstado(_documento.Estado), AutoSize = true };

        layout.Controls.Add(new Label { Text = "Número:", AutoSize = true }, 0, 0);
        layout.Controls.Add(new Label { Text = _documento.CodigoDocumento, AutoSize = true }, 1, 0);
        layout.Controls.Add(new Label { Text = "Tipo:", AutoSize = true }, 0, 1);
        layout.Controls.Add(new Label { Text = _documento.TipoDocumento, AutoSize = true }, 1, 1);
        layout.Controls.Add(new Label { Text = "Cliente:", AutoSize = true }, 0, 2);
        layout.Controls.Add(new Label { Text = _documento.CodigoCliente ?? "Mostrador", AutoSize = true }, 1, 2);
        layout.Controls.Add(new Label { Text = "Neto:", AutoSize = true }, 0, 3);
        layout.Controls.Add(new Label { Text = $"{_documento.Neto:0.00}", AutoSize = true }, 1, 3);
        layout.Controls.Add(new Label { Text = "Total:", AutoSize = true }, 0, 4);
        layout.Controls.Add(new Label { Text = $"{_documento.Total:0.00}", Font = new Font("Segoe UI", 10, FontStyle.Bold), AutoSize = true }, 1, 4);
        layout.Controls.Add(new Label { Text = "Estado:", AutoSize = true }, 0, 5);
        layout.Controls.Add(_lblEstado, 1, 5);

        var btnAnular = new Button { Text = "Anular", Width = 90 };
        var btnCerrar = new Button { Text = "Cerrar", Width = 90 };
        btnCerrar.Click += (_, _) => Close();
        btnAnular.Click += async (_, _) => await AnularAsync();

        var acciones = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
        acciones.Controls.Add(btnAnular);
        acciones.Controls.Add(btnCerrar);
        layout.Controls.Add(acciones, 1, 6);

        Controls.Add(layout);
    }

    private async Task AnularAsync()
    {
        if (_anularHandler is null)
        {
            MessageBox.Show("Función de anulación no disponible.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        using var motivo = new Form { Text = "Motivo de Anulación", Width = 360, Height = 160, StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false };
        var txtMotivo = new TextBox { Left = 20, Top = 20, Width = 300, Name = "txtMotivo" };
        var btnOk = new Button { Left = 100, Top = 60, Width = 90, Text = "Anular", DialogResult = DialogResult.OK };
        var btnCancel = new Button { Left = 200, Top = 60, Width = 90, Text = "Cancelar", DialogResult = DialogResult.Cancel };
        motivo.Controls.AddRange([txtMotivo, btnOk, btnCancel]);
        motivo.AcceptButton = btnOk;
        motivo.CancelButton = btnCancel;

        if (motivo.ShowDialog(this) != DialogResult.OK || string.IsNullOrWhiteSpace(txtMotivo.Text))
            return;

        var result = await _anularHandler.HandleAsync(new AnularDocumentoCommand(_documento.CodigoDocumento, Environment.UserName, txtMotivo.Text.Trim()));
        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        _lblEstado.Text = "Anulado";
        MessageBox.Show("Documento anulado correctamente.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private static string MapEstado(string estado) => estado switch
    {
        "PE" => "Pendiente",
        "CO" => "Cobrado",
        "AN" => "Anulado",
        _ => estado
    };
}
