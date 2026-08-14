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
        txtUsuario = new TextBox();
        txtPassword = new TextBox();
        lblUsuario = new Label();
        lblPassword = new Label();
        btnIngresar = new Button();
        btnCancelar = new Button();
        lblEstado = new Label();
        lblSesion = new Label();
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(480, 240);
        Controls.Add(lblSesion);
        Controls.Add(lblEstado);
        Controls.Add(btnCancelar);
        Controls.Add(btnIngresar);
        Controls.Add(lblPassword);
        Controls.Add(lblUsuario);
        Controls.Add(txtPassword);
        Controls.Add(txtUsuario);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        Name = "Form1";
        StartPosition = FormStartPosition.CenterScreen;
        Text = "INFOREST — Acceso";
        Load += Form1_Load;
        txtUsuario.Location = new Point(136, 38);
        txtUsuario.Name = "txtUsuario";
        txtUsuario.Size = new Size(276, 23);
        txtPassword.Location = new Point(136, 78);
        txtPassword.Name = "txtPassword";
        txtPassword.PasswordChar = '*';
        txtPassword.Size = new Size(276, 23);
        lblUsuario.AutoSize = true;
        lblUsuario.Location = new Point(48, 41);
        lblUsuario.Name = "lblUsuario";
        lblUsuario.Size = new Size(50, 15);
        lblUsuario.Text = "Usuario";
        lblPassword.AutoSize = true;
        lblPassword.Location = new Point(48, 81);
        lblPassword.Name = "lblPassword";
        lblPassword.Size = new Size(67, 15);
        lblPassword.Text = "Contraseña";
        btnIngresar.Location = new Point(136, 123);
        btnIngresar.Name = "btnIngresar";
        btnIngresar.Size = new Size(120, 30);
        btnIngresar.Text = "Ingresar";
        btnIngresar.UseVisualStyleBackColor = true;
        btnIngresar.Click += btnIngresar_Click;
        btnCancelar.Location = new Point(292, 123);
        btnCancelar.Name = "btnCancelar";
        btnCancelar.Size = new Size(120, 30);
        btnCancelar.Text = "Cancelar";
        btnCancelar.UseVisualStyleBackColor = true;
        btnCancelar.Click += btnCancelar_Click;
        lblEstado.AutoSize = true;
        lblEstado.Location = new Point(48, 177);
        lblEstado.Name = "lblEstado";
        lblEstado.Size = new Size(125, 15);
        lblEstado.Text = "Esperando validación.";
        lblSesion.AutoSize = true;
        lblSesion.Location = new Point(48, 201);
        lblSesion.Name = "lblSesion";
        lblSesion.Size = new Size(0, 15);
    }

    #endregion

    private TextBox txtUsuario;
    private TextBox txtPassword;
    private Label lblUsuario;
    private Label lblPassword;
    private Button btnIngresar;
    private Button btnCancelar;
    private Label lblEstado;
    private Label lblSesion;
}
