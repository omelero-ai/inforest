using Inforest.Application.Caja;
using Inforest.Application.Interfaces;
using Inforest.Application.Maestros;
using Inforest.Application.Pedidos;
using Inforest.Application.Turno;
using Inforest.Application.Ventas;
using Inforest.Desktop.Ventas;
using Inforest.Domain.Entities.Maestros;
using Inforest.Domain.Entities.Ventas;

namespace Inforest.Desktop.Pedidos;

/// <summary>
/// Formulario de gestión de pedido de mesa.
/// Legacy: frmPedido.frm. BR-SQL-001, BR-SQL-002, BR-003.
/// </summary>
public class FrmPedido : Form
{
    private readonly Mesa _mesa;
    private readonly IProductoMaestroRepository _productoRepository;
    private readonly ISessionService _sessionService;
    private readonly ObtenerTurnoActualHandler _turnoHandler;
    private readonly CreatePedidoHandler _createPedidoHandler;
    private readonly UpdatePedidoHandler _updatePedidoHandler;
    private readonly EmitirDocumentoHandler? _emitirHandler;
    private readonly ObtenerMediosPagoHandler? _mediosPagoHandler;
    private readonly PagarDocumentoHandler? _pagarHandler;
    private readonly RegistrarPagosMultiplesHandler? _registrarPagosMultiplesHandler;
    private readonly ProcesarPagoPinPadHandler? _procesarPagoPinPadHandler;
    private readonly ObtenerTerminalesPinPadHandler? _obtenerTerminalesPinPadHandler;
    private readonly AnularDocumentoHandler? _anularHandler;

    private DataGridView _dgvProductos = null!;
    private DataGridView _dgvDetalle = null!;
    private TextBox _txtMesa = null!;
    private TextBox _txtMozo = null!;
    private TextBox _txtObservacion = null!;
    private Label _lblTurno = null!;
    private Label _lblCaja = null!;
    private string? _codigoPedidoActual;
    private Pedido? _pedidoCreado;

    public FrmPedido(
        Mesa mesa,
        IProductoMaestroRepository productoRepository,
        ISessionService sessionService,
        ObtenerTurnoActualHandler turnoHandler,
        CreatePedidoHandler createPedidoHandler,
        UpdatePedidoHandler updatePedidoHandler,
        EmitirDocumentoHandler? emitirHandler = null,
        ObtenerMediosPagoHandler? mediosPagoHandler = null,
        PagarDocumentoHandler? pagarHandler = null,
        RegistrarPagosMultiplesHandler? registrarPagosMultiplesHandler = null,
        ProcesarPagoPinPadHandler? procesarPagoPinPadHandler = null,
        ObtenerTerminalesPinPadHandler? obtenerTerminalesPinPadHandler = null,
        AnularDocumentoHandler? anularHandler = null)
    {
        _mesa = mesa;
        _productoRepository = productoRepository;
        _sessionService = sessionService;
        _turnoHandler = turnoHandler;
        _createPedidoHandler = createPedidoHandler;
        _updatePedidoHandler = updatePedidoHandler;
        _emitirHandler = emitirHandler;
        _mediosPagoHandler = mediosPagoHandler;
        _pagarHandler = pagarHandler;
        _registrarPagosMultiplesHandler = registrarPagosMultiplesHandler;
        _procesarPagoPinPadHandler = procesarPagoPinPadHandler;
        _obtenerTerminalesPinPadHandler = obtenerTerminalesPinPadHandler;
        _anularHandler = anularHandler;
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        Text = $"Pedido — {_mesa.Detallado}";
        WindowState = FormWindowState.Maximized;

        var topPanel = new Panel { Dock = DockStyle.Top, Height = 70 };
        _lblCaja = new Label { Left = 20, Top = 15, Width = 220, Text = $"Caja: {_sessionService.SesionActual?.CodigoCaja ?? "--"}" };
        _lblTurno = new Label { Left = 260, Top = 15, Width = 260, Text = "Turno: cargando..." };
        var lblMesa = new Label { Left = 20, Top = 42, Width = 40, Text = "Mesa:" };
        _txtMesa = new TextBox { Left = 65, Top = 39, Width = 80, Name = "txtMesa", Text = _mesa.CodigoMesa };
        var lblMozo = new Label { Left = 160, Top = 42, Width = 45, Text = "Mozo:" };
        _txtMozo = new TextBox { Left = 210, Top = 39, Width = 100, Name = "txtMozo" };
        var lblObs = new Label { Left = 325, Top = 42, Width = 85, Text = "Observación:" };
        _txtObservacion = new TextBox { Left = 415, Top = 39, Width = 260, Name = "txtObservacion" };
        topPanel.Controls.AddRange([_lblCaja, _lblTurno, lblMesa, _txtMesa, lblMozo, _txtMozo, lblObs, _txtObservacion]);

        var split = new SplitContainer { Dock = DockStyle.Fill, SplitterDistance = 450 };

        _dgvProductos = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoGenerateColumns = true
        };

        _dgvDetalle = new DataGridView
        {
            Dock = DockStyle.Fill,
            AllowUserToAddRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoGenerateColumns = true
        };
        _dgvDetalle.DataSource = new BindingSource { DataSource = new List<DetalleGridRow>() };

        split.Panel1.Controls.Add(_dgvProductos);
        split.Panel2.Controls.Add(_dgvDetalle);

        var bottomPanel = new Panel { Dock = DockStyle.Bottom, Height = 50 };
        var btnAgregar = new Button { Left = 20, Top = 10, Width = 110, Text = "Agregar", Name = "btnAgregar" };
        var btnEliminar = new Button { Left = 140, Top = 10, Width = 110, Text = "Eliminar", Name = "btnEliminar" };
        var btnEnviarCocina = new Button { Left = 260, Top = 10, Width = 130, Text = "Enviar Cocina", Name = "btnEnviarCocina" };
        var btnCobrar = new Button { Left = 400, Top = 10, Width = 110, Text = "Cobrar", Name = "btnCobrar" };
        var btnCerrar = new Button { Left = 520, Top = 10, Width = 110, Text = "Cerrar", Name = "btnCerrar" };
        bottomPanel.Controls.AddRange([btnAgregar, btnEliminar, btnEnviarCocina, btnCobrar, btnCerrar]);

        btnAgregar.Click += (_, _) => AgregarProducto();
        btnEliminar.Click += (_, _) => EliminarProducto();
        btnEnviarCocina.Click += async (_, _) => await EnviarPedidoAsync();
        btnCobrar.Click += async (_, _) => await CobrarAsync();
        btnCerrar.Click += (_, _) => Close();

        Controls.Add(split);
        Controls.Add(bottomPanel);
        Controls.Add(topPanel);

        Load += async (_, _) => await CargarDatosAsync();
    }

    private async Task CargarDatosAsync()
    {
        await CargarTurnoAsync();
        await CargarProductosAsync();
    }

    private async Task CargarTurnoAsync()
    {
        var codigoCaja = _sessionService.SesionActual?.CodigoCaja;
        if (string.IsNullOrWhiteSpace(codigoCaja)) return;

        try
        {
            var result = await _turnoHandler.HandleAsync(new ObtenerTurnoActualQuery(codigoCaja));
            if (result.EsExitoso && result.Valor is not null)
                _lblTurno.Text = $"Turno: {result.Valor.CodigoTurno} ({result.Valor.FechaDiaContable:dd/MM/yyyy})";
            else
                _lblTurno.Text = "Turno: sin turno abierto";
        }
        catch
        {
            _lblTurno.Text = "Turno: no disponible";
        }
    }

    private async Task CargarProductosAsync()
    {
        IReadOnlyList<ProductoMaestro> productos;
        try
        {
            productos = await _productoRepository.ObtenerTodosAsync();
        }
        catch
        {
            productos = [];
        }

        if (productos.Count == 0)
            productos =
            [
                ProductoMaestro.Crear("P001", "COM", "Producto 1", "SISTEMA", 10m),
                ProductoMaestro.Crear("P002", "BEB", "Producto 2", "SISTEMA", 15m)
            ];

        _dgvProductos.DataSource = productos.Select(p => new
        {
            p.CodigoProducto,
            p.Detallado,
            Precio = p.PrecioVenta,
            Area = p.Grupo
        }).ToList();
    }

    private void AgregarProducto()
    {
        if (_dgvProductos.CurrentRow?.DataBoundItem is null) return;

        var codigo = _dgvProductos.CurrentRow.Cells["CodigoProducto"].Value?.ToString() ?? string.Empty;
        var descripcion = _dgvProductos.CurrentRow.Cells["Detallado"].Value?.ToString() ?? string.Empty;
        var precio = Convert.ToDecimal(_dgvProductos.CurrentRow.Cells["Precio"].Value ?? 0m);
        var area = _dgvProductos.CurrentRow.Cells["Area"].Value?.ToString();

        var source = (BindingSource)_dgvDetalle.DataSource!;
        var rows = (List<DetalleGridRow>)source.DataSource!;
        var existente = rows.FirstOrDefault(r => r.CodigoProducto == codigo);
        if (existente is null)
            rows.Add(new DetalleGridRow(codigo, descripcion, 1m, precio, area));
        else
        {
            var idx = rows.IndexOf(existente);
            rows[idx] = existente with { Cantidad = existente.Cantidad + 1 };
        }
        source.ResetBindings(false);
    }

    private void EliminarProducto()
    {
        if (_dgvDetalle.CurrentRow?.DataBoundItem is not DetalleGridRow row) return;

        var source = (BindingSource)_dgvDetalle.DataSource!;
        var rows = (List<DetalleGridRow>)source.DataSource!;
        rows.Remove(row);
        source.ResetBindings(false);
    }

    private async Task EnviarPedidoAsync()
    {
        var sesion = _sessionService.SesionActual;
        if (sesion is null)
        {
            MessageBox.Show("No existe sesión activa.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var turnoResult = await _turnoHandler.HandleAsync(new ObtenerTurnoActualQuery(sesion.CodigoCaja));
        if (!turnoResult.EsExitoso || turnoResult.Valor is null)
        {
            MessageBox.Show("No hay turno abierto para esta caja. Abra un turno primero.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var source = (BindingSource)_dgvDetalle.DataSource!;
        var rows = ((List<DetalleGridRow>)source.DataSource!).ToList();
        if (rows.Count == 0)
        {
            MessageBox.Show("Agregue al menos un producto al pedido.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var command = new CreatePedidoCommand(
            CodigoCaja: sesion.CodigoCaja,
            CodigoTurno: turnoResult.Valor.CodigoTurno,
            CodigoMesa: _txtMesa.Text.Trim(),
            CodigoMozo: _txtMozo.Text.Trim(),
            CodigoUsuario: sesion.CodigoUsuario,
            Canal: CanalVenta.Local,
            NumeroAdultos: 1,
            NumeroNinos: 0,
            Observacion: _txtObservacion.Text.Trim(),
            Items: rows.Select(r => new CreateDetalleItem(r.CodigoProducto, string.Empty, string.Empty, r.Cantidad, r.PrecioUnitario, r.Area, null)).ToList());

        var result = await _createPedidoHandler.HandleAsync(command);
        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        _codigoPedidoActual = result.Valor!.CodigoPedido;
        _pedidoCreado = result.Valor!;
        MessageBox.Show($"Pedido {_codigoPedidoActual} enviado a cocina.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private async Task CobrarAsync()
    {
        if (_emitirHandler is null)
        {
            MessageBox.Show("Función de cobro no configurada.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (_pedidoCreado is null)
        {
            MessageBox.Show("Envíe el pedido a cocina antes de cobrar.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        using var frmVenta = new FrmVenta(
            _pedidoCreado,
            _emitirHandler,
            _mediosPagoHandler,
            _pagarHandler,
            _registrarPagosMultiplesHandler,
            _procesarPagoPinPadHandler,
            _obtenerTerminalesPinPadHandler,
            _anularHandler);
        if (frmVenta.ShowDialog(this) == DialogResult.OK)
            Close();
    }

    private sealed record DetalleGridRow(string CodigoProducto, string Descripcion, decimal Cantidad, decimal PrecioUnitario, string? Area);
}
