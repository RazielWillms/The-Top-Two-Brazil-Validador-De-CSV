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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Main));
            this.excel = new System.Windows.Forms.Button();
            this.labellog = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ComboBox();
            this.btnSelectFile = new System.Windows.Forms.Button();
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.panelmid = new System.Windows.Forms.Panel();
            this.LC = new System.Windows.Forms.Label();
            this.grid = new System.Windows.Forms.DataGridView();
            this.depuracao = new System.Windows.Forms.Label();
            this.possuiCabecalho = new System.Windows.Forms.CheckBox();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.btnZoomIn = new System.Windows.Forms.Label();
            this.btnZoomOut = new System.Windows.Forms.Label();
            this.zoom = new System.Windows.Forms.Label();
            this.versao = new System.Windows.Forms.Label();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.panelmid.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grid)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            this.SuspendLayout();
            // 
            // excel
            // 
            this.excel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.excel.Location = new System.Drawing.Point(79, 260);
            this.excel.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
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
            this.labellog.Location = new System.Drawing.Point(263, 7);
            this.labellog.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.labellog.Name = "labellog";
            this.labellog.Size = new System.Drawing.Size(74, 17);
            this.labellog.TabIndex = 5;
            this.labellog.Text = "Registro:";
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(9, 260);
            this.button1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(65, 24);
            this.button1.TabIndex = 6;
            this.button1.Text = "Validar";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Validar_click);
            // 
            // listBox1
            // 
            this.listBox1.BackColor = System.Drawing.SystemColors.Control;
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Items.AddRange(new object[] {
            "Máquinas",
            "Saldos Máquinas",
            "Adiantamentos",
            "Orçamento Balcão",
            "Orçamento Oficina",
            "Estatísticas",
            "Veículos Clientes",
            "Imobilizado Itens",
            "Imobilizado Saldos",
            "Legado Financeiro",
            "Legado Pagamentos",
            "Legado Pedidos",
            "Legado Pedidos Itens",
            "Legado Movimentacao",
            "Grupos",
            "SubGrupos"});
            this.listBox1.Location = new System.Drawing.Point(11, 93);
            this.listBox1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(242, 21);
            this.listBox1.TabIndex = 5;
            // 
            // btnSelectFile
            // 
            this.btnSelectFile.Location = new System.Drawing.Point(11, 145);
            this.btnSelectFile.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
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
            this.txtFilePath.Location = new System.Drawing.Point(79, 145);
            this.txtFilePath.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.ReadOnly = true;
            this.txtFilePath.Size = new System.Drawing.Size(174, 19);
            this.txtFilePath.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label2.Location = new System.Drawing.Point(8, 75);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(62, 17);
            this.label2.TabIndex = 0;
            this.label2.Text = "Layout:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label1.Location = new System.Drawing.Point(9, 126);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(68, 17);
            this.label1.TabIndex = 0;
            this.label1.Text = "Arquivo:";
            // 
            // panelmid
            // 
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
            this.panelmid.Controls.Add(this.label1);
            this.panelmid.Controls.Add(this.button1);
            this.panelmid.Controls.Add(this.listBox1);
            this.panelmid.Controls.Add(this.txtFilePath);
            this.panelmid.Controls.Add(this.label2);
            this.panelmid.Controls.Add(this.btnSelectFile);
            this.panelmid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelmid.Location = new System.Drawing.Point(0, 0);
            this.panelmid.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.panelmid.Name = "panelmid";
            this.panelmid.Size = new System.Drawing.Size(850, 417);
            this.panelmid.TabIndex = 6;
            // 
            // LC
            // 
            this.LC.AutoSize = true;
            this.LC.BackColor = System.Drawing.SystemColors.WindowFrame;
            this.LC.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LC.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
            this.LC.Location = new System.Drawing.Point(136, 166);
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
            this.grid.Location = new System.Drawing.Point(266, 30);
            this.grid.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.grid.Name = "grid";
            this.grid.RowHeadersWidth = 51;
            this.grid.RowTemplate.Height = 24;
            this.grid.Size = new System.Drawing.Size(574, 348);
            this.grid.TabIndex = 11;
            // 
            // depuracao
            // 
            this.depuracao.AutoSize = true;
            this.depuracao.Location = new System.Drawing.Point(16, 319);
            this.depuracao.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.depuracao.Name = "depuracao";
            this.depuracao.Size = new System.Drawing.Size(149, 13);
            this.depuracao.TabIndex = 10;
            this.depuracao.Text = "Mensagem Exibir (Depuração)";
            this.depuracao.Visible = false;
            // 
            // possuiCabecalho
            // 
            this.possuiCabecalho.AutoSize = true;
            this.possuiCabecalho.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.possuiCabecalho.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.possuiCabecalho.Location = new System.Drawing.Point(12, 168);
            this.possuiCabecalho.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.possuiCabecalho.Name = "possuiCabecalho";
            this.possuiCabecalho.Size = new System.Drawing.Size(136, 19);
            this.possuiCabecalho.TabIndex = 8;
            this.possuiCabecalho.Text = "Contém cabeçalho?";
            this.possuiCabecalho.UseVisualStyleBackColor = true;
            // 
            // progressBar
            // 
            this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar.Location = new System.Drawing.Point(266, 383);
            this.progressBar.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(575, 24);
            this.progressBar.TabIndex = 7;
            this.progressBar.Visible = false;
            // 
            // btnZoomIn
            // 
            this.btnZoomIn.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnZoomIn.AutoSize = true;
            this.btnZoomIn.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnZoomIn.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnZoomIn.Location = new System.Drawing.Point(813, 3);
            this.btnZoomIn.Name = "btnZoomIn";
            this.btnZoomIn.Size = new System.Drawing.Size(22, 24);
            this.btnZoomIn.TabIndex = 15;
            this.btnZoomIn.Text = "+";
            this.btnZoomIn.Visible = false;
            this.btnZoomIn.Click += new System.EventHandler(this.ZoomIn_click);
            // 
            // btnZoomOut
            // 
            this.btnZoomOut.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnZoomOut.AutoSize = true;
            this.btnZoomOut.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnZoomOut.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnZoomOut.Location = new System.Drawing.Point(731, 0);
            this.btnZoomOut.Name = "btnZoomOut";
            this.btnZoomOut.Size = new System.Drawing.Size(17, 24);
            this.btnZoomOut.TabIndex = 16;
            this.btnZoomOut.Text = "-";
            this.btnZoomOut.Visible = false;
            this.btnZoomOut.Click += new System.EventHandler(this.ZoomOut_click);
            // 
            // zoom
            // 
            this.zoom.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.zoom.AutoSize = true;
            this.zoom.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.zoom.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.zoom.Location = new System.Drawing.Point(760, 6);
            this.zoom.Name = "zoom";
            this.zoom.Size = new System.Drawing.Size(44, 16);
            this.zoom.TabIndex = 17;
            this.zoom.Text = "zoom";
            this.zoom.Visible = false;
            // 
            // versao
            // 
            this.versao.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.versao.AutoSize = true;
            this.versao.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.versao.Location = new System.Drawing.Point(3, 403);
            this.versao.Name = "versao";
            this.versao.Size = new System.Drawing.Size(28, 13);
            this.versao.TabIndex = 18;
            this.versao.Text = "v0.3";
            // 
            // pictureBox2
            // 
            this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
            this.pictureBox2.Location = new System.Drawing.Point(-9, 0);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(195, 69);
            this.pictureBox2.TabIndex = 19;
            this.pictureBox2.TabStop = false;
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.WindowFrame;
            this.ClientSize = new System.Drawing.Size(850, 417);
            this.Controls.Add(this.panelmid);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "Main";
            this.Text = "Validar CSV";
            this.panelmid.ResumeLayout(false);
            this.panelmid.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.grid)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button excel;
        private System.Windows.Forms.Label labellog;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ComboBox listBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnSelectFile;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.Label label1;
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
    }
}

