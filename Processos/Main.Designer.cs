using System.Windows.Forms;

namespace ValidarCSV
{
    partial class Main
    {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código gerado pelo Windows Form Designer

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Main));
            this.excel = new System.Windows.Forms.Button();
            this.labellog = new System.Windows.Forms.Label();
            this.validar = new System.Windows.Forms.Button();
            this.layouts = new System.Windows.Forms.ComboBox();
            this.btnSelectFile = new System.Windows.Forms.Button();
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.LayoutLabel = new System.Windows.Forms.Label();
            this.ArquivoLabel = new System.Windows.Forms.Label();
            this.panelmid = new System.Windows.Forms.Panel();
            this.CabecalhoLabel = new System.Windows.Forms.Label();
            this.Cabecalho = new System.Windows.Forms.ComboBox();
            this.Mateador = new System.Windows.Forms.PictureBox();
            this.MensagemErro = new System.Windows.Forms.TextBox();
            this.NiveisCombo = new System.Windows.Forms.ComboBox();
            this.Niveis = new System.Windows.Forms.Label();
            this.NivelCombo = new System.Windows.Forms.ComboBox();
            this.Nivel = new System.Windows.Forms.Label();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.versao = new System.Windows.Forms.Label();
            this.zoom = new System.Windows.Forms.Label();
            this.btnZoomOut = new System.Windows.Forms.Label();
            this.btnZoomIn = new System.Windows.Forms.Label();
            this.LC = new System.Windows.Forms.Label();
            this.grid = new System.Windows.Forms.DataGridView();
            this.depuracao = new System.Windows.Forms.Label();
            this.possuiCabecalho = new System.Windows.Forms.CheckBox();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.erroTela = new System.Windows.Forms.ErrorProvider(this.components);
            this.Vermelho = new System.Windows.Forms.PictureBox();
            this.Verde = new System.Windows.Forms.PictureBox();
            this.Prata = new System.Windows.Forms.PictureBox();
            this.Ouro = new System.Windows.Forms.PictureBox();
            this.Preto = new System.Windows.Forms.PictureBox();
            this.Normal = new System.Windows.Forms.PictureBox();
            this.panelmid.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Mateador)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.grid)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.erroTela)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Vermelho)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Verde)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Prata)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Ouro)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Preto)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.Normal)).BeginInit();
            this.SuspendLayout();
            // 
            // excel
            // 
            this.excel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.excel.Location = new System.Drawing.Point(79, 274);
            this.excel.Margin = new System.Windows.Forms.Padding(2);
            this.excel.Name = "excel";
            this.excel.Size = new System.Drawing.Size(64, 24);
            this.excel.TabIndex = 1;
            this.excel.Text = "Exportar";
            this.excel.UseVisualStyleBackColor = true;
            this.excel.Visible = false;
            this.excel.Click += new System.EventHandler(this.Exportar_click);
            // 
            // labellog
            // 
            this.labellog.AutoSize = true;
            this.labellog.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labellog.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.labellog.Location = new System.Drawing.Point(263, 18);
            this.labellog.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.labellog.Name = "labellog";
            this.labellog.Size = new System.Drawing.Size(74, 17);
            this.labellog.TabIndex = 5;
            this.labellog.Text = "Registro:";
            // 
            // validar
            // 
            this.validar.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.validar.Location = new System.Drawing.Point(9, 274);
            this.validar.Margin = new System.Windows.Forms.Padding(2);
            this.validar.Name = "validar";
            this.validar.Size = new System.Drawing.Size(65, 24);
            this.validar.TabIndex = 6;
            this.validar.Text = "Validar";
            this.validar.UseVisualStyleBackColor = true;
            this.validar.Click += new System.EventHandler(this.Validar_click);
            // 
            // layouts
            // 
            this.layouts.BackColor = System.Drawing.SystemColors.Control;
            this.layouts.FormattingEnabled = true;
            this.layouts.Location = new System.Drawing.Point(9, 39);
            this.layouts.Margin = new System.Windows.Forms.Padding(2);
            this.layouts.Name = "layouts";
            this.layouts.Size = new System.Drawing.Size(242, 21);
            this.layouts.TabIndex = 5;
            this.layouts.TextChanged += new System.EventHandler(this.Layout_selecionado);
            // 
            // btnSelectFile
            // 
            this.btnSelectFile.Location = new System.Drawing.Point(9, 91);
            this.btnSelectFile.Margin = new System.Windows.Forms.Padding(2);
            this.btnSelectFile.Name = "btnSelectFile";
            this.btnSelectFile.Size = new System.Drawing.Size(63, 19);
            this.btnSelectFile.TabIndex = 1;
            this.btnSelectFile.Text = "Escolher";
            this.btnSelectFile.UseVisualStyleBackColor = true;
            this.btnSelectFile.Click += new System.EventHandler(this.Escolher_click);
            // 
            // txtFilePath
            // 
            this.txtFilePath.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtFilePath.ForeColor = System.Drawing.SystemColors.InactiveCaption;
            this.txtFilePath.Location = new System.Drawing.Point(76, 91);
            this.txtFilePath.Margin = new System.Windows.Forms.Padding(2);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.ReadOnly = true;
            this.txtFilePath.Size = new System.Drawing.Size(174, 19);
            this.txtFilePath.TabIndex = 2;
            // 
            // LayoutLabel
            // 
            this.LayoutLabel.AutoSize = true;
            this.LayoutLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LayoutLabel.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.LayoutLabel.Location = new System.Drawing.Point(6, 18);
            this.LayoutLabel.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.LayoutLabel.Name = "LayoutLabel";
            this.LayoutLabel.Size = new System.Drawing.Size(62, 17);
            this.LayoutLabel.TabIndex = 0;
            this.LayoutLabel.Text = "Layout:";
            // 
            // ArquivoLabel
            // 
            this.ArquivoLabel.AutoSize = true;
            this.ArquivoLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ArquivoLabel.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.ArquivoLabel.Location = new System.Drawing.Point(7, 70);
            this.ArquivoLabel.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.ArquivoLabel.Name = "ArquivoLabel";
            this.ArquivoLabel.Size = new System.Drawing.Size(68, 17);
            this.ArquivoLabel.TabIndex = 0;
            this.ArquivoLabel.Text = "Arquivo:";
            // 
            // panelmid
            // 
            this.panelmid.Controls.Add(this.Normal);
            this.panelmid.Controls.Add(this.Preto);
            this.panelmid.Controls.Add(this.Ouro);
            this.panelmid.Controls.Add(this.Prata);
            this.panelmid.Controls.Add(this.Verde);
            this.panelmid.Controls.Add(this.Vermelho);
            this.panelmid.Controls.Add(this.CabecalhoLabel);
            this.panelmid.Controls.Add(this.Cabecalho);
            this.panelmid.Controls.Add(this.Mateador);
            this.panelmid.Controls.Add(this.MensagemErro);
            this.panelmid.Controls.Add(this.NiveisCombo);
            this.panelmid.Controls.Add(this.Niveis);
            this.panelmid.Controls.Add(this.NivelCombo);
            this.panelmid.Controls.Add(this.Nivel);
            this.panelmid.Controls.Add(this.pictureBox2);
            this.panelmid.Controls.Add(this.versao);
            this.panelmid.Controls.Add(this.zoom);
            this.panelmid.Controls.Add(this.btnZoomOut);
            this.panelmid.Controls.Add(this.btnZoomIn);
            this.panelmid.Controls.Add(this.LC);
            this.panelmid.Controls.Add(this.labellog);
            this.panelmid.Controls.Add(this.grid);
            this.panelmid.Controls.Add(this.depuracao);
            this.panelmid.Controls.Add(this.possuiCabecalho);
            this.panelmid.Controls.Add(this.progressBar);
            this.panelmid.Controls.Add(this.excel);
            this.panelmid.Controls.Add(this.ArquivoLabel);
            this.panelmid.Controls.Add(this.validar);
            this.panelmid.Controls.Add(this.layouts);
            this.panelmid.Controls.Add(this.txtFilePath);
            this.panelmid.Controls.Add(this.LayoutLabel);
            this.panelmid.Controls.Add(this.btnSelectFile);
            this.panelmid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelmid.Location = new System.Drawing.Point(0, 0);
            this.panelmid.Margin = new System.Windows.Forms.Padding(2);
            this.panelmid.Name = "panelmid";
            this.panelmid.Size = new System.Drawing.Size(717, 414);
            this.panelmid.TabIndex = 6;
            // 
            // CabecalhoLabel
            // 
            this.CabecalhoLabel.AutoSize = true;
            this.CabecalhoLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CabecalhoLabel.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.CabecalhoLabel.Location = new System.Drawing.Point(8, 119);
            this.CabecalhoLabel.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.CabecalhoLabel.Name = "CabecalhoLabel";
            this.CabecalhoLabel.Size = new System.Drawing.Size(155, 17);
            this.CabecalhoLabel.TabIndex = 28;
            this.CabecalhoLabel.Text = "Contém cabeçalho?:";
            // 
            // Cabecalho
            // 
            this.Cabecalho.BackColor = System.Drawing.SystemColors.Control;
            this.Cabecalho.Cursor = System.Windows.Forms.Cursors.Default;
            this.Cabecalho.FormattingEnabled = true;
            this.Cabecalho.Location = new System.Drawing.Point(10, 140);
            this.Cabecalho.Margin = new System.Windows.Forms.Padding(2);
            this.Cabecalho.Name = "Cabecalho";
            this.Cabecalho.Size = new System.Drawing.Size(242, 21);
            this.Cabecalho.TabIndex = 27;
            // 
            // Mateador
            // 
            this.Mateador.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.Mateador.Image = ((System.Drawing.Image)(resources.GetObject("Mateador.Image")));
            this.Mateador.Location = new System.Drawing.Point(382, 103);
            this.Mateador.Margin = new System.Windows.Forms.Padding(2);
            this.Mateador.Name = "Mateador";
            this.Mateador.Size = new System.Drawing.Size(178, 197);
            this.Mateador.TabIndex = 26;
            this.Mateador.TabStop = false;
            this.Mateador.Visible = false;
            // 
            // MensagemErro
            // 
            this.MensagemErro.Location = new System.Drawing.Point(11, 328);
            this.MensagemErro.Margin = new System.Windows.Forms.Padding(2);
            this.MensagemErro.Name = "MensagemErro";
            this.MensagemErro.Size = new System.Drawing.Size(240, 20);
            this.MensagemErro.TabIndex = 25;
            this.MensagemErro.Visible = false;
            // 
            // NiveisCombo
            // 
            this.NiveisCombo.FormattingEnabled = true;
            this.NiveisCombo.Items.AddRange(new object[] {
            "2 (Grupo/SubGrupo)",
            "3 (Grupo/Subgrupo/Segmento)",
            "4 (Grupo/Subgrupo/Segmento/SubSegmento)"});
            this.NiveisCombo.Location = new System.Drawing.Point(9, 190);
            this.NiveisCombo.Margin = new System.Windows.Forms.Padding(2);
            this.NiveisCombo.Name = "NiveisCombo";
            this.NiveisCombo.Size = new System.Drawing.Size(242, 21);
            this.NiveisCombo.TabIndex = 24;
            this.NiveisCombo.Visible = false;
            this.NiveisCombo.TextChanged += new System.EventHandler(this.NiveisCombo_selecionado);
            // 
            // Niveis
            // 
            this.Niveis.AutoSize = true;
            this.Niveis.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Niveis.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.Niveis.Location = new System.Drawing.Point(7, 170);
            this.Niveis.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Niveis.Name = "Niveis";
            this.Niveis.Size = new System.Drawing.Size(148, 17);
            this.Niveis.TabIndex = 23;
            this.Niveis.Text = "Níveis da Empresa:";
            this.Niveis.Visible = false;
            // 
            // NivelCombo
            // 
            this.NivelCombo.FormattingEnabled = true;
            this.NivelCombo.Items.AddRange(new object[] {
            "SubGrupo",
            "Segmento",
            "SubSegmento"});
            this.NivelCombo.Location = new System.Drawing.Point(9, 240);
            this.NivelCombo.Margin = new System.Windows.Forms.Padding(2);
            this.NivelCombo.Name = "NivelCombo";
            this.NivelCombo.Size = new System.Drawing.Size(242, 21);
            this.NivelCombo.TabIndex = 22;
            this.NivelCombo.Visible = false;
            // 
            // Nivel
            // 
            this.Nivel.AutoSize = true;
            this.Nivel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Nivel.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.Nivel.Location = new System.Drawing.Point(7, 219);
            this.Nivel.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.Nivel.Name = "Nivel";
            this.Nivel.Size = new System.Drawing.Size(132, 17);
            this.Nivel.TabIndex = 21;
            this.Nivel.Text = "Nível do Arquivo:";
            this.Nivel.Visible = false;
            // 
            // pictureBox2
            // 
            this.pictureBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
            this.pictureBox2.Location = new System.Drawing.Point(-10, 345);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(195, 48);
            this.pictureBox2.TabIndex = 19;
            this.pictureBox2.TabStop = false;
            // 
            // versao
            // 
            this.versao.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.versao.AutoSize = true;
            this.versao.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.versao.Location = new System.Drawing.Point(9, 395);
            this.versao.Name = "versao";
            this.versao.Size = new System.Drawing.Size(40, 13);
            this.versao.TabIndex = 18;
            this.versao.Text = "Versão";
            // 
            // zoom
            // 
            this.zoom.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.zoom.AutoSize = true;
            this.zoom.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.zoom.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.zoom.Location = new System.Drawing.Point(626, 18);
            this.zoom.Name = "zoom";
            this.zoom.Size = new System.Drawing.Size(44, 16);
            this.zoom.TabIndex = 17;
            this.zoom.Text = "zoom";
            this.zoom.Visible = false;
            // 
            // btnZoomOut
            // 
            this.btnZoomOut.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnZoomOut.AutoSize = true;
            this.btnZoomOut.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnZoomOut.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnZoomOut.Location = new System.Drawing.Point(598, 12);
            this.btnZoomOut.Name = "btnZoomOut";
            this.btnZoomOut.Size = new System.Drawing.Size(17, 24);
            this.btnZoomOut.TabIndex = 16;
            this.btnZoomOut.Text = "-";
            this.btnZoomOut.Visible = false;
            this.btnZoomOut.Click += new System.EventHandler(this.ZoomOut_click);
            // 
            // btnZoomIn
            // 
            this.btnZoomIn.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnZoomIn.AutoSize = true;
            this.btnZoomIn.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnZoomIn.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnZoomIn.Location = new System.Drawing.Point(680, 15);
            this.btnZoomIn.Name = "btnZoomIn";
            this.btnZoomIn.Size = new System.Drawing.Size(22, 24);
            this.btnZoomIn.TabIndex = 15;
            this.btnZoomIn.Text = "+";
            this.btnZoomIn.Visible = false;
            this.btnZoomIn.Click += new System.EventHandler(this.ZoomIn_click);
            // 
            // LC
            // 
            this.LC.AutoSize = true;
            this.LC.BackColor = System.Drawing.SystemColors.WindowFrame;
            this.LC.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LC.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
            this.LC.Location = new System.Drawing.Point(134, 112);
            this.LC.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.LC.Name = "LC";
            this.LC.Size = new System.Drawing.Size(29, 17);
            this.LC.TabIndex = 12;
            this.LC.Text = "csv";
            this.LC.Visible = false;
            // 
            // grid
            // 
            this.grid.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.grid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grid.Location = new System.Drawing.Point(266, 39);
            this.grid.Margin = new System.Windows.Forms.Padding(2);
            this.grid.Name = "grid";
            this.grid.RowHeadersWidth = 51;
            this.grid.RowTemplate.Height = 24;
            this.grid.Size = new System.Drawing.Size(441, 336);
            this.grid.TabIndex = 11;
            this.grid.DataSourceChanged += new System.EventHandler(this.Grid_datasource_alterado);
            // 
            // depuracao
            // 
            this.depuracao.AutoSize = true;
            this.depuracao.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.depuracao.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.depuracao.Location = new System.Drawing.Point(9, 312);
            this.depuracao.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.depuracao.Name = "depuracao";
            this.depuracao.Size = new System.Drawing.Size(91, 13);
            this.depuracao.TabIndex = 10;
            this.depuracao.Text = "DEPURAÇÃO: ";
            this.depuracao.Visible = false;
            // 
            // possuiCabecalho
            // 
            this.possuiCabecalho.AutoSize = true;
            this.possuiCabecalho.Checked = true;
            this.possuiCabecalho.CheckState = System.Windows.Forms.CheckState.Checked;
            this.possuiCabecalho.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.possuiCabecalho.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.possuiCabecalho.Location = new System.Drawing.Point(167, 117);
            this.possuiCabecalho.Margin = new System.Windows.Forms.Padding(2);
            this.possuiCabecalho.Name = "possuiCabecalho";
            this.possuiCabecalho.Size = new System.Drawing.Size(136, 19);
            this.possuiCabecalho.TabIndex = 8;
            this.possuiCabecalho.Text = "Contém cabeçalho?";
            this.possuiCabecalho.UseVisualStyleBackColor = true;
            this.possuiCabecalho.Visible = false;
            // 
            // progressBar
            // 
            this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar.Location = new System.Drawing.Point(266, 379);
            this.progressBar.Margin = new System.Windows.Forms.Padding(2);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(442, 24);
            this.progressBar.TabIndex = 7;
            this.progressBar.Visible = false;
            // 
            // erroTela
            // 
            this.erroTela.ContainerControl = this;
            this.erroTela.Icon = ((System.Drawing.Icon)(resources.GetObject("erroTela.Icon")));
            // 
            // Vermelho
            // 
            this.Vermelho.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.Vermelho.Image = ((System.Drawing.Image)(resources.GetObject("Vermelho.Image")));
            this.Vermelho.Location = new System.Drawing.Point(382, 103);
            this.Vermelho.Margin = new System.Windows.Forms.Padding(2);
            this.Vermelho.Name = "Vermelho";
            this.Vermelho.Size = new System.Drawing.Size(178, 197);
            this.Vermelho.TabIndex = 29;
            this.Vermelho.TabStop = false;
            this.Vermelho.Visible = false;
            // 
            // Verde
            // 
            this.Verde.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.Verde.Image = ((System.Drawing.Image)(resources.GetObject("Verde.Image")));
            this.Verde.Location = new System.Drawing.Point(382, 101);
            this.Verde.Margin = new System.Windows.Forms.Padding(2);
            this.Verde.Name = "Verde";
            this.Verde.Size = new System.Drawing.Size(178, 197);
            this.Verde.TabIndex = 30;
            this.Verde.TabStop = false;
            this.Verde.Visible = false;
            // 
            // Prata
            // 
            this.Prata.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.Prata.Image = ((System.Drawing.Image)(resources.GetObject("Prata.Image")));
            this.Prata.Location = new System.Drawing.Point(382, 101);
            this.Prata.Margin = new System.Windows.Forms.Padding(2);
            this.Prata.Name = "Prata";
            this.Prata.Size = new System.Drawing.Size(178, 197);
            this.Prata.TabIndex = 31;
            this.Prata.TabStop = false;
            this.Prata.Visible = false;
            // 
            // Ouro
            // 
            this.Ouro.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.Ouro.Image = ((System.Drawing.Image)(resources.GetObject("Ouro.Image")));
            this.Ouro.Location = new System.Drawing.Point(382, 101);
            this.Ouro.Margin = new System.Windows.Forms.Padding(2);
            this.Ouro.Name = "Ouro";
            this.Ouro.Size = new System.Drawing.Size(178, 197);
            this.Ouro.TabIndex = 32;
            this.Ouro.TabStop = false;
            this.Ouro.Visible = false;
            // 
            // Preto
            // 
            this.Preto.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.Preto.Image = ((System.Drawing.Image)(resources.GetObject("Preto.Image")));
            this.Preto.Location = new System.Drawing.Point(382, 101);
            this.Preto.Margin = new System.Windows.Forms.Padding(2);
            this.Preto.Name = "Preto";
            this.Preto.Size = new System.Drawing.Size(178, 197);
            this.Preto.TabIndex = 33;
            this.Preto.TabStop = false;
            this.Preto.Visible = false;
            // 
            // Normal
            // 
            this.Normal.BackColor = System.Drawing.SystemColors.AppWorkspace;
            this.Normal.Image = ((System.Drawing.Image)(resources.GetObject("Normal.Image")));
            this.Normal.Location = new System.Drawing.Point(382, 101);
            this.Normal.Margin = new System.Windows.Forms.Padding(2);
            this.Normal.Name = "Normal";
            this.Normal.Size = new System.Drawing.Size(178, 197);
            this.Normal.TabIndex = 34;
            this.Normal.TabStop = false;
            this.Normal.Visible = false;
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.WindowFrame;
            this.ClientSize = new System.Drawing.Size(717, 414);
            this.Controls.Add(this.panelmid);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "Main";
            this.Text = "Validar CSV";
            this.panelmid.ResumeLayout(false);
            this.panelmid.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Mateador)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.grid)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.erroTela)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Vermelho)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Verde)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Prata)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Ouro)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Preto)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.Normal)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button excel;
        private System.Windows.Forms.Label labellog;
        private System.Windows.Forms.Button validar;
        private System.Windows.Forms.ComboBox layouts;
        private System.Windows.Forms.Label LayoutLabel;
        private System.Windows.Forms.Button btnSelectFile;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.Label ArquivoLabel;
        private System.Windows.Forms.Panel panelmid;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.CheckBox possuiCabecalho;
        private System.Windows.Forms.Label depuracao;
        private System.Windows.Forms.DataGridView grid;
        private System.Windows.Forms.Label LC;
        private System.Windows.Forms.Label btnZoomOut;
        private System.Windows.Forms.Label btnZoomIn;
        private System.Windows.Forms.Label zoom;
        private System.Windows.Forms.Label versao;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.ComboBox NivelCombo;
        private System.Windows.Forms.Label Nivel;
        private System.Windows.Forms.ComboBox NiveisCombo;
        private System.Windows.Forms.Label Niveis;
        private System.Windows.Forms.TextBox MensagemErro;
        private ErrorProvider erroTela;
        private PictureBox Mateador;
        private Label CabecalhoLabel;
        private ComboBox Cabecalho;
        private PictureBox Preto;
        private PictureBox Ouro;
        private PictureBox Prata;
        private PictureBox Verde;
        private PictureBox Vermelho;
        private PictureBox Normal;
    }
}

