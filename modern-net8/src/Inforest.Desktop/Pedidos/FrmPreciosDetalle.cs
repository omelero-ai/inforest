using Inforest.Application.Configuracion;
using Inforest.Application.Interfaces;
using Inforest.Application.Interfaces.Country;
using Inforest.Application.Precios;
using Inforest.Infrastructure.Country;
using Microsoft.Extensions.Configuration;

namespace Inforest.Desktop.Pedidos;

/// <summary>
/// Editor de precio / descuento / recargo para un ítem de pedido.
/// Legacy: frmPreciosDetalle.frm
/// Permite modificar precio, descuento, recargo e impuestos de un ítem activo de DPEDIDO.
/// BR-PRECIO-001: Solo lModificable=true puede editar precio
/// BR-PRECIO-002: Auditoría en TLOG_MODPRECIO
/// BR-PRECIO-003: Validación descuento/recargo
/// BR-PRECIO-005: Bolivia vs Perú/Ecuador cálculo impuestos
/// BR-PRECIO-006: Exoneración/Inafectación oculta impuestos
/// </summary>
public sealed class FrmPreciosDetalle : Form
{
    // ── Dependencias ──────────────────────────────────────────────────────────
    private readonly ObtenerItemParaPrecioHandler  _obtenerHandler;
    private readonly ModificarPrecioItemHandler    _modificarHandler;
    private readonly IParametroRepository          _parametroRepo;
    private readonly IPaisPolicy                   _paisPolicy;
    private readonly ISessionService               _session;
    private readonly string                        _codigoPedido;
    private readonly string                        _item;

    // ── Datos del ítem ────────────────────────────────────────────────────────
    private ItemPedidoParaPrecio? _itemData;
    private decimal _precioVenta;
    private decimal _precioOficial;
    private decimal _precioNeto;
    private decimal _descuento;
    private decimal _recargo;
    private decimal _impuesto1;
    private decimal _impuesto2;
    private decimal _impuesto3;
    private decimal _cantidad;
    private decimal _porcentaje1;
    private decimal _porcentaje2;
    private decimal _porcentaje3;
    private string  _codigoPais = string.Empty;  // ISO: "BO" = Bolivia, others = Perú/Ecuador

    // ── Resultado ─────────────────────────────────────────────────────────────
    public bool GuardadoOk { get; private set; }

    // ── Controls ──────────────────────────────────────────────────────────────
    private readonly Label  lblProducto    = new() { Dock = DockStyle.Top, Height = 30, Font = new System.Drawing.Font("Segoe UI", 11f, System.Drawing.FontStyle.Bold), TextAlign = ContentAlignment.MiddleCenter };
    private readonly Label  lblOficial     = new() { Text = "0.00", TextAlign = ContentAlignment.MiddleRight, Width = 140, Height = 24 };
    private readonly Label  lblNeto        = new() { Text = "0.00", TextAlign = ContentAlignment.MiddleRight, Width = 140, Height = 24 };
    private readonly Label  lblImpuesto1   = new() { Text = "0.00", TextAlign = ContentAlignment.MiddleRight, Width = 140, Height = 24 };
    private readonly Label  lblImpuesto2   = new() { Text = "0.00", TextAlign = ContentAlignment.MiddleRight, Width = 140, Height = 24 };
    private readonly Label  lblImpuesto3   = new() { Text = "0.00", TextAlign = ContentAlignment.MiddleRight, Width = 140, Height = 24 };
    private readonly Label  lblPrecioVenta = new() { Text = "0.00", TextAlign = ContentAlignment.MiddleRight, Width = 140, Height = 24, Font = new System.Drawing.Font("Segoe UI", 10f, System.Drawing.FontStyle.Bold) };
    private readonly Label  lblDescuento   = new() { Text = "0.00", TextAlign = ContentAlignment.MiddleRight, Width = 140, Height = 24 };
    private readonly Label  lblRecargo     = new() { Text = "0.00", TextAlign = ContentAlignment.MiddleRight, Width = 140, Height = 24 };
    private readonly Label  lblCantidad    = new() { Text = "0.00", TextAlign = ContentAlignment.MiddleRight, Width = 140, Height = 24 };
    private readonly Label  lblVenta       = new() { Text = "0.00", TextAlign = ContentAlignment.MiddleRight, Width = 140, Height = 24, Font = new System.Drawing.Font("Segoe UI", 10f, System.Drawing.FontStyle.Bold) };
    private readonly Label  lblImpNombre1  = new() { AutoSize = true };
    private readonly Label  lblImpNombre2  = new() { AutoSize = true };
    private readonly Label  lblImpNombre3  = new() { AutoSize = true };
    private readonly Label  lblDescProd    = new() { AutoSize = true, ForeColor = System.Drawing.Color.DarkBlue };

    private readonly Button btnDsctoMonto   = new() { Text = "Monto Descuento",    Height = 36 };
    private readonly Button btnDsctoPct     = new() { Text = "% Descuento",        Height = 36 };
    private readonly Button btnRecargoMonto = new() { Text = "Monto Recargo",      Height = 36 };
    private readonly Button btnRecargoPct   = new() { Text = "% Recargo",          Height = 36 };
    private readonly Button btnPrecio       = new() { Text = "Modificar Precio",   Height = 36 };
    private readonly Button btnGrabar       = new() { Text = "Grabar",  Width = 100, Height = 36 };
    private readonly Button btnSalir        = new() { Text = "Salir",   Width = 100, Height = 36 };
    private readonly Label  lblStatus       = new() { Dock = DockStyle.Bottom, Height = 22, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(4, 0, 0, 0) };

    public FrmPreciosDetalle(
        ObtenerItemParaPrecioHandler obtenerHandler,
        ModificarPrecioItemHandler modificarHandler,
        IParametroRepository parametroRepo,
        PaisPolicyFactory paisFactory,
        IConfiguration configuration,
        ISessionService session,
        string codigoPedido,
        string item)
    {
        _obtenerHandler   = obtenerHandler;
        _modificarHandler = modificarHandler;
        _parametroRepo    = parametroRepo;
        _session          = session;
        _codigoPedido     = codigoPedido;
        _item             = item;
        var codigoPais    = configuration["Inforest:Pais"] ?? "PE";
        _paisPolicy       = paisFactory.ObtenerPolitica(codigoPais);

        Text            = $"Precio / Descuento — Pedido {codigoPedido} · Ítem {item}";
        Width           = 560;
        Height          = 560;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox     = false;
        MinimizeBox     = false;
        StartPosition   = FormStartPosition.CenterParent;

        BuildLayout();
        WireEvents();

        Load += async (_, _) => await InicializarAsync();
    }

    // ── Layout ────────────────────────────────────────────────────────────────

    private void BuildLayout()
    {
        Controls.Add(lblProducto);
        Controls.Add(lblStatus);

        var main = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 15,
            Padding = new Padding(8)
        };
        main.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        main.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

        void AddRow(string caption, Control valueCtrl, int rowIndex)
        {
            var lbl = new Label { Text = caption, AutoSize = false, Width = 200, Height = 24, TextAlign = ContentAlignment.MiddleRight };
            main.Controls.Add(lbl,       0, rowIndex);
            main.Controls.Add(valueCtrl, 1, rowIndex);
        }

        AddRow("Precio Oficial:",  lblOficial,     0);
        AddRow("Descuento:",       lblDescuento,   1);
        AddRow("Recargo:",         lblRecargo,     2);
        AddRow("Precio Venta:",    lblPrecioVenta, 3);
        AddRow("Precio Neto:",     lblNeto,        4);
        main.Controls.Add(lblImpNombre1, 0, 5); main.Controls.Add(lblImpuesto1, 1, 5);
        main.Controls.Add(lblImpNombre2, 0, 6); main.Controls.Add(lblImpuesto2, 1, 6);
        main.Controls.Add(lblImpNombre3, 0, 7); main.Controls.Add(lblImpuesto3, 1, 7);
        AddRow("Cantidad:",        lblCantidad,    8);
        AddRow("Total Venta:",     lblVenta,       9);
        main.Controls.Add(lblDescProd, 0, 10); main.SetColumnSpan(lblDescProd, 2);

        // Buttons grid (2x3)
        var btnPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, AutoSize = true };
        btnPanel.Controls.AddRange([btnDsctoMonto, btnDsctoPct, btnRecargoMonto, btnRecargoPct, btnPrecio]);
        main.Controls.Add(btnPanel, 0, 11); main.SetColumnSpan(btnPanel, 2);

        var footer = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, AutoSize = true };
        footer.Controls.AddRange([btnSalir, btnGrabar]);
        main.Controls.Add(footer, 0, 12); main.SetColumnSpan(footer, 2);

        Controls.Add(main);

        foreach (var btn in new[] { btnDsctoMonto, btnDsctoPct, btnRecargoMonto, btnRecargoPct, btnPrecio })
        {
            btn.AutoSize = true;
            btn.Margin   = new Padding(2);
        }
    }

    private void WireEvents()
    {
        btnPrecio.Click       += async (_, _) => await CambiarPrecioAsync();
        btnDsctoMonto.Click   += async (_, _) => await AplicarDescuentoMontoAsync();
        btnDsctoPct.Click     += async (_, _) => await AplicarDescuentoPorcentajeAsync();
        btnRecargoMonto.Click += async (_, _) => await AplicarRecargoMontoAsync();
        btnRecargoPct.Click   += async (_, _) => await AplicarRecargoPorcentajeAsync();
        btnGrabar.Click       += async (_, _) => await GrabarAsync();
        btnSalir.Click        += (_, _) => Close();
    }

    // ── Inicialización ────────────────────────────────────────────────────────

    private async Task InicializarAsync()
    {
        lblStatus.Text = "Cargando…";

        // Leer configuración (país, tasas de impuesto)
        var config = await _parametroRepo.ObtenerConfiguracionAsync();
        _codigoPais  = _paisPolicy.CodigoPais;  // "BO" = Bolivia (impuestos incluidos)
        _porcentaje1 = (decimal)(config?.Impuesto1 ?? 0);
        _porcentaje2 = (decimal)(config?.Impuesto2 ?? 0);
        _porcentaje3 = (decimal)(config?.Impuesto3 ?? 0);

        var nomImp1 = config?.tImpuesto1 ?? "IMP1";
        var nomImp2 = config?.tImpuesto2 ?? "IMP2";
        var nomImp3 = config?.tImpuesto3;
        lblImpNombre1.Text = !string.IsNullOrEmpty(nomImp1) ? $"{nomImp1}:" : string.Empty;
        lblImpNombre2.Text = !string.IsNullOrEmpty(nomImp2) ? $"{nomImp2}:" : string.Empty;
        lblImpNombre3.Text = !string.IsNullOrEmpty(nomImp3) ? $"{nomImp3}:" : string.Empty;

        // Cargar ítem del pedido
        var result = await _obtenerHandler.HandleAsync(new ObtenerItemParaPrecioQuery(_codigoPedido, _item));
        if (!result.EsExitoso)
        {
            lblStatus.Text = result.MensajeError;
            MessageBox.Show(result.MensajeError, "INFOREST", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            Close();
            return;
        }

        _itemData = result.Valor!;
        CargarDatos(_itemData);
        lblStatus.Text = "Listo.";
    }

    private void CargarDatos(ItemPedidoParaPrecio item)
    {
        lblProducto.Text = item.Producto;

        // BR-PRECIO-001: solo modificable si lModificable
        bool editable = item.Modificable;
        btnPrecio.Enabled       = editable;
        btnDsctoMonto.Enabled   = editable;
        btnDsctoPct.Enabled     = editable;
        btnRecargoMonto.Enabled = editable;
        btnRecargoPct.Enabled   = editable;
        btnGrabar.Enabled       = editable;

        _precioOficial = item.PrecioOficial;
        _precioNeto    = item.PrecioNeto;
        _precioVenta   = item.PrecioVenta;
        _descuento     = item.Descuento;
        _recargo       = item.Recargo;
        _impuesto1     = item.Impuesto1;
        _impuesto2     = item.Impuesto2;
        _impuesto3     = item.Impuesto3;
        _cantidad      = item.Cantidad;

        // BR-PRECIO-006: Exoneración / Inafectación oculta impuesto1
        if (item.AplicaExoneracion)
            lblDescProd.Text = "Producto Exonerado";
        else if (item.AplicaInafectacion)
            lblDescProd.Text = "Producto Inafecto";
        else if (item.Impuesto1 == 0)
            lblDescProd.Text = $"Producto no afecto a {lblImpNombre1.Text.TrimEnd(':')}";
        else
            lblDescProd.Text = string.Empty;

        RefrescarVista();
    }

    private void RefrescarVista()
    {
        string F(decimal v) => v.ToString("N2");
        lblOficial.Text     = F(_precioOficial);
        lblDescuento.Text   = F(_descuento);
        lblRecargo.Text     = F(_recargo);
        lblPrecioVenta.Text = F(_precioVenta);
        lblNeto.Text        = F(_precioNeto);
        lblImpuesto1.Text   = F(_impuesto1);
        lblImpuesto2.Text   = F(_impuesto2);
        lblImpuesto3.Text   = F(_impuesto3);
        lblCantidad.Text    = F(_cantidad);
        lblVenta.Text       = F(_precioVenta * _cantidad);
    }

    // ── Cálculo de precios ────────────────────────────────────────────────────

    /// <summary>
    /// BR-PRECIO-005: Bolivia nPVenta incluye impuesto. Perú/Ecuador nPVenta es antes de impuesto.
    /// </summary>
    private void RecalcularImpuestos()
    {
        bool imp1Activo = _impuesto1 != 0 || (_itemData?.AplicaExoneracion == false && _itemData?.AplicaInafectacion == false);

        if (_codigoPais == "BO") // Bolivia — impuesto incluido en precio venta
        {
            var imp1 = imp1Activo ? _precioVenta * _porcentaje1 / 100 : 0;
            var imp2 = _precioVenta * _porcentaje2 / 100;
            var imp3 = _precioVenta * _porcentaje3 / 100;
            _impuesto1 = imp1;
            _impuesto2 = imp2;
            _impuesto3 = imp3;
            _precioNeto = _precioVenta - imp1 - imp2 - imp3;
        }
        else // Perú, Ecuador
        {
            decimal acc = 0;
            if (imp1Activo) acc += _porcentaje1;
            acc += _porcentaje2;
            acc += _porcentaje3;
            decimal div = 1 + (acc / 100);
            _impuesto1 = div > 0 && imp1Activo ? _precioVenta / div * _porcentaje1 / 100 : 0;
            _impuesto2 = div > 0 ? _precioVenta / div * _porcentaje2 / 100 : 0;
            _impuesto3 = div > 0 ? _precioVenta / div * _porcentaje3 / 100 : 0;
            _precioNeto = _precioVenta - _impuesto1 - _impuesto2 - _impuesto3;
        }
    }

    // ── Eventos de botones ────────────────────────────────────────────────────

    private async Task CambiarPrecioAsync()
    {
        var texto = await PedirNumeroAsync("Ingrese nuevo precio de venta:");
        if (texto is null) return;
        if (!decimal.TryParse(texto, out decimal valor) || valor <= 0) { MostrarError("Precio no válido."); return; }

        _precioVenta   = valor;
        _precioOficial = valor;
        _descuento     = 0;
        _recargo       = 0;
        RecalcularImpuestos();
        RefrescarVista();
    }

    private async Task AplicarDescuentoMontoAsync()
    {
        if (_precioNeto == 0) return;
        var texto = await PedirNumeroAsync("Ingrese monto de descuento:");
        if (texto is null) return;
        if (!decimal.TryParse(texto, out decimal valor) || valor <= 0) { MostrarError("Monto no válido, debe ser mayor a 0."); return; }
        if (valor > _precioOficial) { MostrarError("Monto no válido."); return; }

        _descuento   = valor;
        _recargo     = 0;
        _precioVenta = _precioOficial - _descuento + _recargo;
        RecalcularImpuestos();
        RefrescarVista();
    }

    private async Task AplicarDescuentoPorcentajeAsync()
    {
        if (_precioNeto == 0) return;
        var texto = await PedirNumeroAsync("Ingrese porcentaje de descuento (0-99):");
        if (texto is null) return;
        if (!decimal.TryParse(texto, out decimal pct) || pct <= 0 || pct >= 100) { MostrarError("Porcentaje no válido, debe ser mayor a 0 y menor a 100."); return; }

        _descuento   = _precioOficial * pct / 100;
        _recargo     = 0;
        _precioVenta = _precioOficial - _descuento + _recargo;
        RecalcularImpuestos();
        RefrescarVista();
    }

    private async Task AplicarRecargoMontoAsync()
    {
        if (_precioNeto == 0) return;
        var texto = await PedirNumeroAsync("Ingrese monto de recargo:");
        if (texto is null) return;
        if (!decimal.TryParse(texto, out decimal valor) || valor <= 0) { MostrarError("Monto no válido, debe ser mayor a 0."); return; }

        _recargo     = valor;
        _precioVenta = _precioOficial - _descuento + _recargo;
        RecalcularImpuestos();
        RefrescarVista();
    }

    private async Task AplicarRecargoPorcentajeAsync()
    {
        if (_precioNeto == 0) return;
        var texto = await PedirNumeroAsync("Ingrese porcentaje de recargo (0-99):");
        if (texto is null) return;
        if (!decimal.TryParse(texto, out decimal pct) || pct <= 0 || pct >= 100) { MostrarError("Porcentaje no válido, debe ser mayor a 0 y menor a 100."); return; }

        _recargo     = _precioOficial * pct / 100;
        _precioVenta = _precioOficial - _descuento + _recargo;
        RecalcularImpuestos();
        RefrescarVista();
    }

    // ── Grabar ────────────────────────────────────────────────────────────────

    private async Task GrabarAsync()
    {
        // BR-PRECIO-003: validación
        if (_descuento == 0 && _recargo == 0 && _precioVenta == (_itemData?.PrecioVenta ?? _precioVenta))
        {
            // Nada cambió, confirmar de todas formas
        }

        var usuarioActual = _session.SesionActual?.CodigoUsuario ?? string.Empty;

        var command = new ModificarPrecioItemCommand(
            CodigoPedido    : _codigoPedido,
            Item            : _item,
            PrecioNeto      : _precioNeto,
            PrecioOficial   : _precioOficial,
            PrecioVenta     : _precioVenta,
            Descuento       : _descuento,
            Recargo         : _recargo,
            Impuesto1       : _impuesto1,
            Impuesto2       : _impuesto2,
            Impuesto3       : _impuesto3,
            Cantidad        : _cantidad,
            UsuarioAutoriza : usuarioActual,
            Usuario         : usuarioActual);

        var result = await _modificarHandler.HandleAsync(command);
        if (!result.EsExitoso)
        {
            MostrarError(result.MensajeError!);
            return;
        }

        GuardadoOk = true;
        MessageBox.Show("Precio modificado correctamente.", "INFOREST", MessageBoxButtons.OK, MessageBoxIcon.Information);
        Close();
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    /// <summary>Muestra un diálogo de entrada numérica simple.</summary>
    private static Task<string?> PedirNumeroAsync(string mensaje)
    {
        // Simple input dialog en lugar de frmNumPad (no disponible aún)
        using var frm = new Form
        {
            Text = "INFOREST", Width = 320, Height = 130,
            FormBorderStyle = FormBorderStyle.FixedDialog, StartPosition = FormStartPosition.CenterParent,
            MaximizeBox = false, MinimizeBox = false
        };
        var lbl = new Label { Text = mensaje, Dock = DockStyle.Top, Height = 30, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(8, 0, 0, 0) };
        var txt = new TextBox { Dock = DockStyle.Top, Font = new System.Drawing.Font("Segoe UI", 14f), TextAlign = HorizontalAlignment.Right };
        var btnOk  = new Button { Text = "Aceptar", DialogResult = DialogResult.OK,  Width = 100 };
        var btnCnl = new Button { Text = "Cancelar",DialogResult = DialogResult.Cancel, Width = 100 };
        var pnl = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 40, FlowDirection = FlowDirection.RightToLeft };
        pnl.Controls.AddRange([btnCnl, btnOk]);
        frm.Controls.AddRange([lbl, txt, pnl]);
        frm.AcceptButton = btnOk; frm.CancelButton = btnCnl;

        var res = frm.ShowDialog();
        return Task.FromResult(res == DialogResult.OK && !string.IsNullOrWhiteSpace(txt.Text) ? txt.Text.Trim() : (string?)null);
    }

    private static void MostrarError(string msg) =>
        MessageBox.Show(msg, "INFOREST", MessageBoxButtons.OK, MessageBoxIcon.Warning);
}
