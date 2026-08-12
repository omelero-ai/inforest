namespace Inforest.Desktop.Ventas;

/// <summary>
/// Vista de documento emitido.
/// Legacy: frmDocumento.frm.
/// </summary>
public class FrmDocumento : Form
{
    public FrmDocumento()
    {
        Text = "Documento Emitido";
        Width = 520;
        Height = 300;
        StartPosition = FormStartPosition.CenterParent;

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(16),
            ColumnCount = 2,
            RowCount = 6
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 65));

        var lblNumero = new Label { Text = "01-PED001", AutoSize = true };
        var lblTipo = new Label { Text = "Boleta", AutoSize = true };
        var lblCliente = new Label { Text = "Cliente mostrador", AutoSize = true };
        var lblMonto = new Label { Text = "23.60", AutoSize = true };
        var lblEstado = new Label { Text = "Pendiente", AutoSize = true };

        layout.Controls.Add(new Label { Text = "Número:", AutoSize = true }, 0, 0);
        layout.Controls.Add(lblNumero, 1, 0);
        layout.Controls.Add(new Label { Text = "Tipo:", AutoSize = true }, 0, 1);
        layout.Controls.Add(lblTipo, 1, 1);
        layout.Controls.Add(new Label { Text = "Cliente:", AutoSize = true }, 0, 2);
        layout.Controls.Add(lblCliente, 1, 2);
        layout.Controls.Add(new Label { Text = "Monto:", AutoSize = true }, 0, 3);
        layout.Controls.Add(lblMonto, 1, 3);
        layout.Controls.Add(new Label { Text = "Estado:", AutoSize = true }, 0, 4);
        layout.Controls.Add(lblEstado, 1, 4);

        var btnAnular = new Button { Text = "Anular", Width = 90 };
        var btnReimprimir = new Button { Text = "Reimprimir", Width = 90 };
        var btnCerrar = new Button { Text = "Cerrar", Width = 90 };
        btnCerrar.Click += (_, _) => Close();

        var acciones = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
        acciones.Controls.Add(btnAnular);
        acciones.Controls.Add(btnReimprimir);
        acciones.Controls.Add(btnCerrar);
        layout.Controls.Add(acciones, 1, 5);

        Controls.Add(layout);
    }
}
