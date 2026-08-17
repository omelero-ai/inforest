namespace Inforest.Desktop;

partial class Form1
{
    /// <summary>
    ///  Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    ///  Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    ///  Required method for Designer support - do not modify
    ///  the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
        panelHeader = new Panel();
        lblTitulo = new Label();
        lblSubtitulo = new Label();
        panelBody = new TableLayoutPanel();
        txtUsuario = new TextBox();
        txtPassword = new TextBox();
        lblUsuario = new Label();
        lblPassword = new Label();
        btnIngresar = new Button();
        btnCancelar = new Button();
        lblEstado = new Label();
        lblSesion = new Label();
        panelButtons = new FlowLayoutPanel();
        panelHeader.SuspendLayout();
        panelBody.SuspendLayout();
        panelButtons.SuspendLayout();
        SuspendLayout();
        AutoScaleMode = AutoScaleMode.Font;
        BackColor = Color.FromArgb(245, 246, 247);
        ClientSize = new Size(520, 320);
        Controls.Add(panelBody);
        Controls.Add(panelHeader);
        AcceptButton = btnIngresar;
        CancelButton = btnCancelar;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        Name = "Form1";
        StartPosition = FormStartPosition.CenterScreen;
        Text = "INFOREST — Acceso";
        Load += Form1_Load;
        panelHeader.BackColor = Color.FromArgb(235, 236, 240);
        panelHeader.Controls.Add(lblSubtitulo);
        panelHeader.Controls.Add(lblTitulo);
        panelHeader.Dock = DockStyle.Top;
        panelHeader.Location = new Point(0, 0);
        panelHeader.Name = "panelHeader";
        panelHeader.Padding = new Padding(18, 16, 18, 12);
        panelHeader.Size = new Size(520, 88);
        lblTitulo.AutoSize = true;
        lblTitulo.Font = new Font("Segoe UI", 13F, FontStyle.Bold);
        lblTitulo.ForeColor = Color.FromArgb(35, 35, 35);
        lblTitulo.Location = new Point(18, 16);
        lblTitulo.Name = "lblTitulo";
        lblTitulo.Size = new Size(153, 25);
        lblTitulo.Text = "Inicio de sesión";
        lblSubtitulo.AutoSize = true;
        lblSubtitulo.ForeColor = Color.DimGray;
        lblSubtitulo.Location = new Point(20, 50);
        lblSubtitulo.Name = "lblSubtitulo";
        lblSubtitulo.Size = new Size(313, 15);
        lblSubtitulo.Text = "Acceso operativo a INFOREST por usuario, caja y terminal.";
        panelBody.ColumnCount = 2;
        panelBody.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110F));
        panelBody.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        panelBody.Controls.Add(lblUsuario, 0, 0);
        panelBody.Controls.Add(txtUsuario, 1, 0);
        panelBody.Controls.Add(lblPassword, 0, 1);
        panelBody.Controls.Add(txtPassword, 1, 1);
        panelBody.Controls.Add(lblEstado, 0, 2);
        panelBody.Controls.Add(lblSesion, 0, 3);
        panelBody.Controls.Add(panelButtons, 0, 4);
        panelBody.Dock = DockStyle.Fill;
        panelBody.Location = new Point(0, 88);
        panelBody.Name = "panelBody";
        panelBody.Padding = new Padding(24, 22, 24, 18);
        panelBody.RowCount = 5;
        panelBody.RowStyles.Add(new RowStyle(SizeType.Absolute, 42F));
        panelBody.RowStyles.Add(new RowStyle(SizeType.Absolute, 42F));
        panelBody.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));
        panelBody.RowStyles.Add(new RowStyle(SizeType.Absolute, 34F));
        panelBody.RowStyles.Add(new RowStyle(SizeType.Absolute, 46F));
        panelBody.Size = new Size(520, 232);
        panelBody.SetColumnSpan(lblEstado, 2);
        panelBody.SetColumnSpan(lblSesion, 2);
        panelBody.SetColumnSpan(panelButtons, 2);
        txtUsuario.Anchor = AnchorStyles.Left | AnchorStyles.Right;
        txtUsuario.CharacterCasing = CharacterCasing.Upper;
        txtUsuario.Name = "txtUsuario";
        txtUsuario.Size = new Size(362, 23);
        txtUsuario.Leave += txtUsuario_Leave;
        txtPassword.Anchor = AnchorStyles.Left | AnchorStyles.Right;
        txtPassword.Name = "txtPassword";
        txtPassword.PasswordChar = '*';
        txtPassword.Size = new Size(362, 23);
        lblUsuario.AutoSize = true;
        lblUsuario.Anchor = AnchorStyles.Left;
        lblUsuario.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
        lblUsuario.Name = "lblUsuario";
        lblUsuario.Size = new Size(52, 15);
        lblUsuario.Text = "Usuario";
        lblPassword.AutoSize = true;
        lblPassword.Anchor = AnchorStyles.Left;
        lblPassword.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
        lblPassword.Name = "lblPassword";
        lblPassword.Size = new Size(67, 15);
        lblPassword.Text = "Contraseña";
        btnIngresar.BackColor = Color.FromArgb(0, 122, 204);
        btnIngresar.FlatAppearance.BorderSize = 0;
        btnIngresar.FlatStyle = FlatStyle.Flat;
        btnIngresar.ForeColor = Color.White;
        btnIngresar.Margin = new Padding(0, 0, 12, 0);
        btnIngresar.Name = "btnIngresar";
        btnIngresar.Size = new Size(120, 30);
        btnIngresar.Text = "Ingresar";
        btnIngresar.UseVisualStyleBackColor = false;
        btnIngresar.Click += btnIngresar_Click;
        btnCancelar.BackColor = Color.White;
        btnCancelar.DialogResult = DialogResult.Cancel;
        btnCancelar.FlatAppearance.BorderColor = Color.FromArgb(207, 211, 216);
        btnCancelar.FlatStyle = FlatStyle.Flat;
        btnCancelar.Name = "btnCancelar";
        btnCancelar.Size = new Size(120, 30);
        btnCancelar.Text = "Cancelar";
        btnCancelar.UseVisualStyleBackColor = false;
        btnCancelar.Click += btnCancelar_Click;
        lblEstado.AutoSize = true;
        lblEstado.ForeColor = Color.DimGray;
        lblEstado.Name = "lblEstado";
        lblEstado.Size = new Size(125, 15);
        lblEstado.Text = "Esperando validación.";
        lblSesion.AutoSize = true;
        lblSesion.ForeColor = Color.DimGray;
        lblSesion.Name = "lblSesion";
        lblSesion.Size = new Size(0, 15);
        panelButtons.Anchor = AnchorStyles.Right;
        panelButtons.Controls.Add(btnCancelar);
        panelButtons.Controls.Add(btnIngresar);
        panelButtons.FlowDirection = FlowDirection.RightToLeft;
        panelButtons.Location = new Point(241, 142);
        panelButtons.Margin = new Padding(0, 10, 0, 0);
        panelButtons.Name = "panelButtons";
        panelButtons.Size = new Size(255, 34);
        panelHeader.ResumeLayout(false);
        panelHeader.PerformLayout();
        panelBody.ResumeLayout(false);
        panelBody.PerformLayout();
        panelButtons.ResumeLayout(false);
        ResumeLayout(false);
    }

    #endregion

    private Panel panelHeader;
    private Label lblTitulo;
    private Label lblSubtitulo;
    private TableLayoutPanel panelBody;
    private TextBox txtUsuario;
    private TextBox txtPassword;
    private Label lblUsuario;
    private Label lblPassword;
    private Button btnIngresar;
    private Button btnCancelar;
    private Label lblEstado;
    private Label lblSesion;
    private FlowLayoutPanel panelButtons;
}
