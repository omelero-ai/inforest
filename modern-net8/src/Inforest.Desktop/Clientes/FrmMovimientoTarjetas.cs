using Inforest.Application.Delivery;
using Inforest.Domain.Common;
using Inforest.Domain.Entities.Delivery;

namespace Inforest.Desktop.Clientes;

/// <summary>
/// Vista de los últimos movimientos de una tarjeta RFID/proximidad.
/// Legacy: <c>FrmMovimientoTarjetas.frm</c> (shell vacío en VB6) +
///         funcionalidad de impresión de <c>FrmTarjetaAproximidad.cmdOpcion(5)</c>.
/// Consulta: <c>SELECT top 10 ... FROM TMOVIMIENTOTARJETASRFID WHERE CodidoRFID = @rfid ORDER BY fregistro DESC</c>.
/// Reglas: BR-RFID-001 (contexto de visualización de movimientos).
/// </summary>
public sealed class FrmMovimientoTarjetas : Form
{
    private static readonly Color SurfaceColor = Color.FromArgb(245, 246, 247);
    private static readonly Color HeaderColor = Color.FromArgb(235, 236, 240);
    private static readonly Color AccentColor = Color.FromArgb(0, 122, 204);

    private readonly ObtenerMovimientosTarjetaProximidadHandler _handler;
    private readonly string _codigoRfid;
    private readonly string _nombreCliente;

    private readonly DataGridView _dgv;
    private readonly Label _lblInfo;

    /// <summary>
    /// Muestra los últimos movimientos de una tarjeta RFID.
    /// </summary>
    /// <param name="handler">Handler de consulta de movimientos.</param>
    /// <param name="codigoRfid">Código de la tarjeta RFID (CodidoRFID).</param>
    /// <param name="nombreCliente">Nombre del cliente asociado para mostrar en cabecera.</param>
    public FrmMovimientoTarjetas(
        ObtenerMovimientosTarjetaProximidadHandler handler,
        string codigoRfid,
        string nombreCliente = "")
    {
        _handler = handler;
        _codigoRfid = codigoRfid;
        _nombreCliente = nombreCliente;

        Text = "Movimientos de Tarjeta de Proximidad";
        Size = new Size(720, 480);
        MinimumSize = new Size(600, 380);
        StartPosition = FormStartPosition.CenterParent;
        BackColor = SurfaceColor;
        FormBorderStyle = FormBorderStyle.Sizable;

        // Cabecera
        var header = new Panel
        {
            Dock = DockStyle.Top,
            Height = 56,
            BackColor = HeaderColor,
            Padding = new Padding(12, 0, 12, 0)
        };

        var lblTarjeta = new Label
        {
            Text = $"Tarjeta: {codigoRfid}",
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            ForeColor = Color.FromArgb(35, 35, 35),
            AutoSize = true,
            Left = 12,
            Top = 8
        };

        _lblInfo = new Label
        {
            Text = $"Cliente: {nombreCliente}",
            Font = new Font("Segoe UI", 9),
            ForeColor = Color.FromArgb(80, 80, 80),
            AutoSize = true,
            Left = 12,
            Top = 30
        };

        header.Controls.Add(lblTarjeta);
        header.Controls.Add(_lblInfo);

        // Grid de movimientos
        _dgv = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            BackgroundColor = SurfaceColor,
            BorderStyle = BorderStyle.None,
            GridColor = Color.FromArgb(220, 222, 225),
            RowHeadersVisible = false,
            Font = new Font("Segoe UI", 9),
            ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = HeaderColor,
                ForeColor = Color.FromArgb(35, 35, 35),
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                Alignment = DataGridViewContentAlignment.MiddleLeft
            },
            DefaultCellStyle = new DataGridViewCellStyle
            {
                SelectionBackColor = Color.FromArgb(210, 230, 255),
                SelectionForeColor = Color.FromArgb(35, 35, 35)
            }
        };

        // Pie
        var footer = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 48,
            BackColor = SurfaceColor,
            Padding = new Padding(12, 6, 12, 6)
        };

        var btnCerrar = new Button
        {
            Text = "Cerrar",
            Width = 110,
            Height = 34,
            Left = footer.Width - 130,
            Top = 7,
            Anchor = AnchorStyles.Right | AnchorStyles.Top,
            FlatStyle = FlatStyle.Flat,
            BackColor = AccentColor,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnCerrar.FlatAppearance.BorderSize = 0;
        btnCerrar.Click += (_, _) => Close();

        footer.Controls.Add(btnCerrar);

        Controls.Add(_dgv);
        Controls.Add(header);
        Controls.Add(footer);

        Load += async (_, _) => await CargarMovimientosAsync();
    }

    private async Task CargarMovimientosAsync()
    {
        var result = await _handler.HandleAsync(
            new ObtenerMovimientosTarjetaProximidadQuery(_codigoRfid, 10));

        if (!result.EsExitoso)
        {
            MessageBox.Show(result.MensajeError, Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            return;
        }

        var movimientos = result.Valor!;

        _dgv.DataSource = null;
        _dgv.DataSource = movimientos
            .Select(m => new
            {
                Fecha = m.FechaRegistro.ToString("dd/MM/yyyy"),
                Ingreso = m.MontoIngreso,
                Salida = m.MontoSalida,
                Anterior = m.MontoAnterior,
                Final = m.MontoFinal
            })
            .ToList();

        if (_dgv.Columns.Count == 0)
            return;

        _dgv.Columns[0].HeaderText = "Fecha";
        _dgv.Columns[0].FillWeight = 20;
        _dgv.Columns[1].HeaderText = "M. Ingreso";
        _dgv.Columns[1].DefaultCellStyle.Format = "N2";
        _dgv.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        _dgv.Columns[1].FillWeight = 20;
        _dgv.Columns[2].HeaderText = "M. Salida";
        _dgv.Columns[2].DefaultCellStyle.Format = "N2";
        _dgv.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        _dgv.Columns[2].FillWeight = 20;
        _dgv.Columns[3].HeaderText = "M. Anterior";
        _dgv.Columns[3].DefaultCellStyle.Format = "N2";
        _dgv.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        _dgv.Columns[3].FillWeight = 20;
        _dgv.Columns[4].HeaderText = "M. Final";
        _dgv.Columns[4].DefaultCellStyle.Format = "N2";
        _dgv.Columns[4].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        _dgv.Columns[4].FillWeight = 20;

        if (movimientos.Count == 0)
            _lblInfo.Text += "  |  Sin movimientos registrados";
    }
}
