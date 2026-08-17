using Inforest.Application.Ventas;

namespace Inforest.Desktop.Ventas;

/// <summary>
/// Formulario de cambio de tipo de documento.
/// Permite cambiar el tipo de comprobante de un documento ya emitido (ej. boleta → factura).
/// Legacy: frmCambio.frm. BR-CAMBIO-001..005.
/// SQL: MDOCUMENTO, DDOCUMENTO, DPEDIDO, DPAGODOCUMENTO, TTIPODOCUMENTOIMPRESORA, vTipoDocumento.
/// </summary>
public sealed class FrmCambioDocumento : Form
{
    private readonly string _codigoDocumentoOrigen;
    private readonly string _codigoClienteOriginal;
    private readonly string _descripcionClienteOriginal;
    private readonly string _codigoCaja;
    private readonly string _codigoTurno;
    private readonly string _codigoSalon;
    private readonly string _pais;
    private readonly ObtenerTiposDocumentoCambioHandler? _obtenerTiposHandler;
    private readonly CambiarDocumentoHandler? _cambiarHandler;

    private IReadOnlyList<TipoDocumentoDisponible> _tipos = [];
    private TipoDocumentoDisponible? _tipoSeleccionado;
    private string _codigoClienteNuevo = string.Empty;
    private string _descripcionClienteNuevo = string.Empty;

    private readonly Label _lblDocumentoActual;
    private readonly Label _lblDocumentoNuevo;
    private readonly Label _lblCliente;
    private readonly TextBox _txtCliente;
    private readonly Panel _panelTipos;
    private readonly Button _btnAceptar;
    private readonly Button _btnCancelar;
    private readonly Button _btnCliente;

    /// <summary>Número del nuevo documento generado tras el cambio.</summary>
    public string? CodigoDocumentoNuevo { get; private set; }

    public FrmCambioDocumento(
        string codigoDocumentoOrigen,
        string codigoClienteOriginal,
        string descripcionClienteOriginal,
        string codigoCaja,
        string pais,
        string codigoTurno = "",
        string codigoSalon = "",
        ObtenerTiposDocumentoCambioHandler? obtenerTiposHandler = null,
        CambiarDocumentoHandler? cambiarHandler = null)
    {
        _codigoDocumentoOrigen = codigoDocumentoOrigen;
        _codigoClienteOriginal = codigoClienteOriginal;
        _descripcionClienteOriginal = descripcionClienteOriginal;
        _codigoCaja = codigoCaja;
        _pais = pais;
        _codigoTurno = codigoTurno;
        _codigoSalon = codigoSalon;
        _obtenerTiposHandler = obtenerTiposHandler;
        _cambiarHandler = cambiarHandler;

        _codigoClienteNuevo = codigoClienteOriginal;
        _descripcionClienteNuevo = descripcionClienteOriginal;

        Text = "Cambio de Documento";
        Width = 480;
        Height = 380;
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        _lblDocumentoActual = new Label { AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
        _lblDocumentoNuevo = new Label { AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold), ForeColor = Color.DarkGreen };
        _lblCliente = new Label { Text = "Cliente:", AutoSize = true };
        _txtCliente = new TextBox { Width = 240, ReadOnly = true };
        _panelTipos = new Panel { Height = 60, Dock = DockStyle.None };
        _btnCliente = new Button { Text = "Buscar Cliente", Width = 120, Height = 28 };
        _btnAceptar = new Button { Text = "✔ Aceptar", Width = 110, Height = 32, BackColor = Color.FromArgb(0, 122, 204), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
        _btnCancelar = new Button { Text = "✘ Cancelar", Width = 110, Height = 32, FlatStyle = FlatStyle.Flat };

        _btnAceptar.FlatAppearance.BorderSize = 0;
        _btnCancelar.Click += (_, _) => { DialogResult = DialogResult.Cancel; Close(); };
        _btnAceptar.Click += async (_, _) => await AceptarAsync();
        _btnCliente.Click += (_, _) => BuscarCliente();

        var main = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(16),
            ColumnCount = 2,
            RowCount = 7,
            AutoSize = true
        };
        main.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
        main.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        main.Controls.Add(new Label { Text = "Documento actual:", AutoSize = true }, 0, 0);
        main.Controls.Add(_lblDocumentoActual, 1, 0);
        main.Controls.Add(new Label { Text = "Documento nuevo:", AutoSize = true }, 0, 1);
        main.Controls.Add(_lblDocumentoNuevo, 1, 1);

        main.Controls.Add(new Label { Text = "Tipo de documento:", AutoSize = true, Font = new Font("Segoe UI", 9, FontStyle.Bold) }, 0, 2);
        main.SetColumnSpan(_panelTipos, 2);
        main.Controls.Add(_panelTipos, 0, 3);

        main.Controls.Add(_lblCliente, 0, 4);
        var clienteRow = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight };
        clienteRow.Controls.Add(_txtCliente);
        clienteRow.Controls.Add(_btnCliente);
        main.Controls.Add(clienteRow, 1, 4);

        var btnRow = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, AutoSize = true };
        btnRow.Controls.Add(_btnCancelar);
        btnRow.Controls.Add(_btnAceptar);
        main.SetColumnSpan(btnRow, 2);
        main.Controls.Add(btnRow, 0, 6);

        Controls.Add(main);

        Load += async (_, _) => await CargarAsync();
    }

    private async Task CargarAsync()
    {
        _lblDocumentoActual.Text = FormatearDocumento(_codigoDocumentoOrigen);
        _txtCliente.Text = _descripcionClienteNuevo;

        if (_obtenerTiposHandler is null)
        {
            MessageBox.Show("Servicio de tipos de documento no disponible.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var result = await _obtenerTiposHandler.HandleAsync(
            new ObtenerTiposDocumentoCambioQuery(_codigoCaja, _pais));

        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        _tipos = result.Valor!;
        _panelTipos.Controls.Clear();

        int x = 0;
        foreach (var tipo in _tipos)
        {
            var btn = new Button
            {
                Text = tipo.Descripcion,
                Width = 120,
                Height = 36,
                Left = x,
                Top = 8,
                FlatStyle = FlatStyle.Flat,
                Tag = tipo
            };
            btn.Click += OnTipoDocumentoClick;
            _panelTipos.Controls.Add(btn);
            x += 128;
        }
    }

    private void OnTipoDocumentoClick(object? sender, EventArgs e)
    {
        if (sender is not Button btn || btn.Tag is not TipoDocumentoDisponible tipo)
            return;

        // Resaltar seleccionado
        foreach (Control c in _panelTipos.Controls)
            if (c is Button b)
                b.BackColor = SystemColors.Control;
        btn.BackColor = Color.FromArgb(0, 122, 204);
        btn.ForeColor = Color.White;

        _tipoSeleccionado = tipo;

        // Generar nuevo número
        var correlativo = Lib.Correlativo(tipo.UltimoNumero, 9);
        var nuevoDoc = tipo.Prefijo + tipo.Serie + correlativo;
        _lblDocumentoNuevo.Text = FormatearDocumento(nuevoDoc);

        // Mostrar/ocultar cliente
        bool mostrarCliente = tipo.RequiereCliente;
        _lblCliente.Visible = mostrarCliente;
        _txtCliente.Visible = mostrarCliente;
        _btnCliente.Visible = mostrarCliente;

        if (mostrarCliente && string.IsNullOrWhiteSpace(_codigoClienteNuevo))
            BuscarCliente();
    }

    private void BuscarCliente()
    {
        // En la implementación real se abre frmBusquedaRapida o FrmBusquedaCliente.
        // Por ahora, un input simple de código de cliente.
        using var dlg = new Form
        {
            Text = "Buscar Cliente",
            Width = 340,
            Height = 130,
            StartPosition = FormStartPosition.CenterParent,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false
        };
        var lbl = new Label { Text = "Código / RUC / DNI:", Left = 12, Top = 16, AutoSize = true };
        var txt = new TextBox { Left = 12, Top = 38, Width = 280 };
        var btnOk = new Button { Text = "Buscar", Left = 120, Top = 70, Width = 80, DialogResult = DialogResult.OK };
        var btnCan = new Button { Text = "Cancelar", Left = 210, Top = 70, Width = 80, DialogResult = DialogResult.Cancel };
        dlg.Controls.AddRange([lbl, txt, btnOk, btnCan]);
        dlg.AcceptButton = btnOk;
        dlg.CancelButton = btnCan;

        if (dlg.ShowDialog(this) == DialogResult.OK && !string.IsNullOrWhiteSpace(txt.Text))
        {
            _codigoClienteNuevo = txt.Text.Trim();
            _descripcionClienteNuevo = txt.Text.Trim();
            _txtCliente.Text = _descripcionClienteNuevo;
        }
    }

    private async Task AceptarAsync()
    {
        if (_tipoSeleccionado is null)
        {
            MessageBox.Show("Seleccione el tipo de documento destino.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (_tipoSeleccionado.RequiereCliente && string.IsNullOrWhiteSpace(_codigoClienteNuevo))
        {
            MessageBox.Show("El tipo de documento requiere un cliente.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // Pedir motivo
        string motivo = PedirTexto("Motivo del Cambio", "Ingrese el motivo del cambio de documento:");
        if (string.IsNullOrWhiteSpace(motivo)) return;

        // Generar nuevo número
        var correlativo = Lib.Correlativo(_tipoSeleccionado.UltimoNumero, 9);
        var nuevoDoc = _tipoSeleccionado.Prefijo + _tipoSeleccionado.Serie + correlativo;

        // Preguntar por consumo
        bool esConsumo = false;
        string descConsumo = string.Empty;
        if (!nuevoDoc.StartsWith("0", StringComparison.Ordinal))
        {
            var resp = MessageBox.Show("¿Por Consumo?", Text, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (resp == DialogResult.Cancel) return;
            if (resp == DialogResult.Yes)
            {
                descConsumo = PedirTexto("Consumo", "Descripción del consumo:", "POR CONSUMO");
                if (string.IsNullOrWhiteSpace(descConsumo))
                {
                    MessageBox.Show("La descripción del consumo no puede estar vacía.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                esConsumo = true;
            }
        }

        if (_cambiarHandler is null)
        {
            MessageBox.Show("Servicio de cambio de documento no disponible.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var command = new CambiarDocumentoCommand(
            CodigoDocumentoOrigen: _codigoDocumentoOrigen,
            TipoEmisionDestino: _tipoSeleccionado.TipoEmision,
            CodigoDocumentoNuevo: nuevoDoc,
            CodigoClienteNuevo: _codigoClienteNuevo,
            Motivo: motivo,
            EsConsumo: esConsumo,
            DescripcionConsumo: descConsumo,
            CodigoTurno: _codigoTurno,
            CodigoSalon: _codigoSalon);

        _btnAceptar.Enabled = false;
        Cursor = Cursors.WaitCursor;

        try
        {
            var result = await _cambiarHandler.HandleAsync(command);

            if (!result.EsExitoso)
            {
                MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            CodigoDocumentoNuevo = result.Valor!.CodigoDocumentoNuevo;
            MessageBox.Show("Documento cambiado satisfactoriamente.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            DialogResult = DialogResult.OK;
            Close();
        }
        finally
        {
            _btnAceptar.Enabled = true;
            Cursor = Cursors.Default;
        }
    }

    private string PedirTexto(string titulo, string mensaje, string valorInicial = "")
    {
        using var dlg = new Form
        {
            Text = titulo,
            Width = 400,
            Height = 140,
            StartPosition = FormStartPosition.CenterParent,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false
        };
        var lbl = new Label { Text = mensaje, Left = 12, Top = 14, AutoSize = true };
        var txt = new TextBox { Left = 12, Top = 36, Width = 360, Text = valorInicial };
        var btnOk = new Button { Text = "Aceptar", Left = 160, Top = 74, Width = 90, DialogResult = DialogResult.OK };
        var btnCan = new Button { Text = "Cancelar", Left = 260, Top = 74, Width = 90, DialogResult = DialogResult.Cancel };
        dlg.Controls.AddRange([lbl, txt, btnOk, btnCan]);
        dlg.AcceptButton = btnOk;
        dlg.CancelButton = btnCan;

        return dlg.ShowDialog(this) == DialogResult.OK ? txt.Text.Trim() : string.Empty;
    }

    private static string FormatearDocumento(string doc)
    {
        // Formato: P-SSSSS-NNNNNNNNN  (similar al Format VB6 "@-@@@@@-@@@@@@@@@")
        if (doc.Length == 15)
            return $"{doc[0]}-{doc[1..6]}-{doc[6..15]}";
        return doc;
    }
}

/// <summary>Utilidades compatibles con el Legacy para el manejo de correlativos en el Desktop.</summary>
internal static class Lib
{
    /// <summary>
    /// Incrementa un correlativo numérico al siguiente número con el ancho indicado.
    /// Legacy: Lib.Correlativo(tUltimoNumero, 9).
    /// </summary>
    public static string Correlativo(string ultimoNumero, int ancho)
    {
        if (!long.TryParse(ultimoNumero, out var n))
            n = 0;
        return (n + 1).ToString().PadLeft(ancho, '0');
    }
}
