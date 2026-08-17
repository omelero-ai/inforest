using Inforest.Application.Ventas;
using Inforest.Domain.Entities.Configuracion;

namespace Inforest.Desktop.Ventas;

/// <summary>
/// Diálogo modal de selección de motivo de anulación.
/// Legacy: frmListaMotivos.frm.
/// BR-MOTIVO-001: lista activos. BR-MOTIVO-002: falla si no hay motivos. BR-MOTIVO-003: devuelve selección.
/// </summary>
public sealed class FrmListaMotivos : Form
{
    private readonly ObtenerMotivosAnulacionHandler _handler;

    /// <summary>
    /// Motivo seleccionado por el usuario.
    /// Null si el usuario canceló (equivalente a wEnter = False en VB6).
    /// BR-MOTIVO-003.
    /// </summary>
    public MotivoAnulacion? MotivoSeleccionado { get; private set; }

    private readonly FlowLayoutPanel _panelMotivos;
    private readonly Button _btnCancelar;

    public FrmListaMotivos(ObtenerMotivosAnulacionHandler handler)
    {
        _handler = handler;

        Text = "Motivo de Anulación";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        MaximizeBox = false;
        MinimizeBox = false;
        Width = 480;
        Height = 400;
        BackColor = Color.FromArgb(245, 246, 247);

        var titulo = new Label
        {
            Text = "Seleccione el motivo de anulación",
            Dock = DockStyle.Top,
            Height = 40,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(12, 0, 0, 0),
            BackColor = Color.FromArgb(235, 236, 240)
        };

        _panelMotivos = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            Padding = new Padding(16),
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false
        };

        var footer = new Panel { Dock = DockStyle.Bottom, Height = 48 };
        _btnCancelar = new Button
        {
            Text = "Cancelar",
            Width = 110,
            Height = 34,
            Left = footer.Width - 130,
            Top = 7,
            Anchor = AnchorStyles.Right | AnchorStyles.Top,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(220, 53, 69),
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        _btnCancelar.FlatAppearance.BorderSize = 0;
        _btnCancelar.Click += (_, _) =>
        {
            MotivoSeleccionado = null;
            DialogResult = DialogResult.Cancel;
            Close();
        };
        footer.Controls.Add(_btnCancelar);

        Controls.Add(_panelMotivos);
        Controls.Add(footer);
        Controls.Add(titulo);

        Load += async (_, _) => await CargarMotivosAsync();
    }

    private async Task CargarMotivosAsync()
    {
        var resultado = await _handler.HandleAsync(new ObtenerMotivosAnulacionQuery());

        if (!resultado.EsExitoso)
        {
            MessageBox.Show(resultado.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            DialogResult = DialogResult.Cancel;
            Close();
            return;
        }

        foreach (var motivo in resultado.Valor!.Motivos)
        {
            var btn = CrearBotonMotivo(motivo);
            _panelMotivos.Controls.Add(btn);
        }

        // Ajustar altura según cantidad de motivos
        var totalHeight = resultado.Valor.Motivos.Count * 50 + 120;
        Height = Math.Min(Math.Max(totalHeight, 200), 600);
    }

    private Button CrearBotonMotivo(MotivoAnulacion motivo)
    {
        var btn = new Button
        {
            Text = motivo.Descripcion,
            Width = _panelMotivos.Width - 36,
            Height = 42,
            Margin = new Padding(0, 4, 0, 4),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.White,
            ForeColor = Color.FromArgb(35, 35, 35),
            Font = new Font("Segoe UI", 9.5f),
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(8, 0, 0, 0),
            Cursor = Cursors.Hand,
            Tag = motivo
        };
        btn.FlatAppearance.BorderColor = Color.FromArgb(207, 211, 216);
        btn.FlatAppearance.BorderSize = 1;
        btn.MouseEnter += (_, _) =>
        {
            btn.BackColor = Color.FromArgb(234, 245, 255);
            btn.FlatAppearance.BorderColor = Color.FromArgb(64, 22, 13);
        };
        btn.MouseLeave += (_, _) =>
        {
            btn.BackColor = Color.White;
            btn.FlatAppearance.BorderColor = Color.FromArgb(207, 211, 216);
        };
        btn.Click += (_, _) =>
        {
            MotivoSeleccionado = motivo;
            DialogResult = DialogResult.OK;
            Close();
        };
        return btn;
    }
}
