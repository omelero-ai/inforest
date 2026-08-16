using System.ComponentModel;
using Inforest.Application.Caja;
using Inforest.Application.Interfaces;
using Inforest.Application.Maestros;
using Inforest.Application.Pedidos;
using Inforest.Application.Turno;
using Inforest.Application.Ventas;
using Inforest.Desktop.Caja;
using Inforest.Domain.Entities.Maestros;
using Inforest.Domain.Entities.Ventas;

namespace Inforest.Desktop.CajaRapida;

/// <summary>
/// Legacy: <c>frmCajaRapida.frm</c>, <c>CajaRapida.exe</c>, <c>modCajaRapida.bas</c>.
/// Venta rápida sin mesa, con cobro inmediato.
/// </summary>
public class FrmCajaRapida : Form
{
    private readonly IProductoMaestroRepository _productoRepository;
    private readonly ISessionService _sessionService;
    private readonly ObtenerTurnoActualHandler _turnoHandler;
    private readonly CreatePedidoHandler _createPedidoHandler;
    private readonly EmitirDocumentoHandler _emitirHandler;
    private readonly ObtenerMediosPagoHandler _mediosPagoHandler;
    private readonly PagarDocumentoHandler _pagarHandler;
    private readonly RegistrarPagosMultiplesHandler _registrarPagosMultiplesHandler;
    private readonly ProcesarPagoPinPadHandler _procesarPagoPinPadHandler;
    private readonly ObtenerTerminalesPinPadHandler _obtenerTerminalesPinPadHandler;

    private readonly FlowLayoutPanel _catalogoPanel;
    private readonly BindingList<ItemRapidoRow> _items = [];
    private readonly DataGridView _grid;
    private readonly Label _lblTotal;

    public FrmCajaRapida(
        IProductoMaestroRepository productoRepository,
        ISessionService sessionService,
        ObtenerTurnoActualHandler turnoHandler,
        CreatePedidoHandler createPedidoHandler,
        EmitirDocumentoHandler emitirHandler,
        ObtenerMediosPagoHandler mediosPagoHandler,
        PagarDocumentoHandler pagarHandler,
        RegistrarPagosMultiplesHandler registrarPagosMultiplesHandler,
        ProcesarPagoPinPadHandler procesarPagoPinPadHandler,
        ObtenerTerminalesPinPadHandler obtenerTerminalesPinPadHandler)
    {
        _productoRepository = productoRepository;
        _sessionService = sessionService;
        _turnoHandler = turnoHandler;
        _createPedidoHandler = createPedidoHandler;
        _emitirHandler = emitirHandler;
        _mediosPagoHandler = mediosPagoHandler;
        _pagarHandler = pagarHandler;
        _registrarPagosMultiplesHandler = registrarPagosMultiplesHandler;
        _procesarPagoPinPadHandler = procesarPagoPinPadHandler;
        _obtenerTerminalesPinPadHandler = obtenerTerminalesPinPadHandler;

        Text = "Caja Rápida";
        WindowState = FormWindowState.Maximized;

        _catalogoPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Left,
            Width = 420,
            AutoScroll = true,
            Padding = new Padding(8)
        };

        _grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AutoGenerateColumns = true,
            AllowUserToAddRows = false,
            DataSource = _items
        };

        _lblTotal = new Label
        {
            Dock = DockStyle.Bottom,
            Height = 42,
            Padding = new Padding(8),
            Font = new Font("Segoe UI", 12, FontStyle.Bold),
            Text = "Total: 0.00"
        };

        var botones = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 44, Padding = new Padding(8), FlowDirection = FlowDirection.RightToLeft };
        var btnSalir = new Button { Text = "Salir", Width = 100 };
        var btnLimpiar = new Button { Text = "Limpiar", Width = 100 };
        var btnCobrar = new Button { Text = "Cobrar", Width = 100 };
        btnSalir.Click += (_, _) => Close();
        btnLimpiar.Click += (_, _) => { _items.Clear(); ActualizarTotal(); };
        btnCobrar.Click += async (_, _) => await CobrarAsync();
        botones.Controls.AddRange([btnSalir, btnLimpiar, btnCobrar]);

        Controls.Add(_grid);
        Controls.Add(_catalogoPanel);
        Controls.Add(_lblTotal);
        Controls.Add(botones);
        Load += async (_, _) => await CargarProductosAsync();
    }

    private async Task CargarProductosAsync()
    {
        IReadOnlyList<ProductoMaestro> productos;
        try { productos = await _productoRepository.ObtenerTodosAsync(); }
        catch { productos = []; }

        if (productos.Count == 0)
            productos =
            [
                ProductoMaestro.Crear("P001", "BEB", "Gaseosa", "SISTEMA", 8m),
                ProductoMaestro.Crear("P002", "COM", "Hamburguesa", "SISTEMA", 18m),
                ProductoMaestro.Crear("P003", "POS", "Papas fritas", "SISTEMA", 10m)
            ];

        _catalogoPanel.Controls.Clear();
        foreach (var producto in productos.Take(24))
        {
            var boton = new Button
            {
                Width = 120,
                Height = 80,
                Margin = new Padding(6),
                Text = $"{producto.Detallado}\n{producto.PrecioVenta:C}"
            };
            boton.Click += (_, _) => AgregarItem(producto);
            _catalogoPanel.Controls.Add(boton);
        }
    }

    private void AgregarItem(ProductoMaestro producto)
    {
        var existente = _items.FirstOrDefault(i => i.CodigoProducto == producto.CodigoProducto);
        if (existente is null)
            _items.Add(new ItemRapidoRow(producto.CodigoProducto, producto.Detallado, 1, producto.PrecioVenta));
        else
        {
            var index = _items.IndexOf(existente);
            _items[index] = existente with { Cantidad = existente.Cantidad + 1 };
        }
        _grid.Refresh();
        ActualizarTotal();
    }

    private void ActualizarTotal() => _lblTotal.Text = $"Total: {_items.Sum(i => i.Total):C}";

    private async Task CobrarAsync()
    {
        if (_items.Count == 0)
        {
            MessageBox.Show("Agregue productos antes de cobrar.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var sesion = _sessionService.SesionActual;
        if (sesion is null)
        {
            MessageBox.Show("No existe sesión activa.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var turnoResult = await _turnoHandler.HandleAsync(new ObtenerTurnoActualQuery(sesion.CodigoCaja));
        if (!turnoResult.EsExitoso || turnoResult.Valor is null)
        {
            MessageBox.Show("No hay turno abierto. Abra un turno primero.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // Crear pedido rápido
        var command = new CreatePedidoCommand(
            CodigoCaja: sesion.CodigoCaja,
            CodigoTurno: turnoResult.Valor.CodigoTurno,
            CodigoMesa: "RAPIDA",
            CodigoMozo: string.Empty,
            CodigoUsuario: sesion.CodigoUsuario,
            Canal: CanalVenta.Local,
            NumeroAdultos: 1,
            NumeroNinos: 0,
            Observacion: "Caja Rápida",
            Items: _items.Select(i => new CreateDetalleItem(i.CodigoProducto, string.Empty, string.Empty, i.Cantidad, i.PrecioUnitario, null, null)).ToList());

        var pedidoResult = await _createPedidoHandler.HandleAsync(command);
        if (!pedidoResult.EsExitoso)
        {
            MessageBox.Show(pedidoResult.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // Emitir documento
        var emitirCommand = new EmitirDocumentoCommand(pedidoResult.Valor!.CodigoPedido, "01", null, 0m, 0m);
        var docResult = await _emitirHandler.HandleAsync(emitirCommand);
        if (!docResult.EsExitoso)
        {
            MessageBox.Show(docResult.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // Cobrar
        using var frmPago = new FrmPago(
            docResult.Valor!.CodigoDocumento,
            docResult.Valor!.Total,
            docResult.Valor!.CodigoCaja,
            _mediosPagoHandler,
            _pagarHandler,
            _registrarPagosMultiplesHandler,
            _procesarPagoPinPadHandler,
            _obtenerTerminalesPinPadHandler);
        if (frmPago.ShowDialog(this) == DialogResult.OK)
        {
            _items.Clear();
            ActualizarTotal();
            MessageBox.Show("Venta registrada correctamente.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private sealed record ItemRapidoRow(string CodigoProducto, string Descripcion, int Cantidad, decimal PrecioUnitario)
    {
        public decimal Total => Cantidad * PrecioUnitario;
    }
}
