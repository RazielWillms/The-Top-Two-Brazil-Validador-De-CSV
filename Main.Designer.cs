namespace The_Top_Two_Brazil_Validador_De_CSV
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
            this.log = new System.Windows.Forms.ListBox();
            this.listBox1 = new System.Windows.Forms.ComboBox();
            this.btnSelectFile = new System.Windows.Forms.Button();
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.panelmid = new System.Windows.Forms.Panel();
            this.depuracao = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.possuiCabecalho = new System.Windows.Forms.CheckBox();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.grid = new System.Windows.Forms.DataGridView();
            this.LC = new System.Windows.Forms.Label();
            this.panelmid.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.grid)).BeginInit();
            this.SuspendLayout();
            // 
            // excel
            // 
            this.excel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.excel.Location = new System.Drawing.Point(105, 271);
            this.excel.Name = "excel";
            this.excel.Size = new System.Drawing.Size(85, 29);
            this.excel.TabIndex = 1;
            this.excel.Text = "Exportar";
            this.excel.UseVisualStyleBackColor = true;
            this.excel.Click += new System.EventHandler(this.Exportar_click);
            // 
            // labellog
            // 
            this.labellog.AutoSize = true;
            this.labellog.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labellog.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.labellog.Location = new System.Drawing.Point(351, 9);
            this.labellog.Name = "labellog";
            this.labellog.Size = new System.Drawing.Size(86, 20);
            this.labellog.TabIndex = 5;
            this.labellog.Text = "Registro:";
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(12, 271);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(87, 29);
            this.button1.TabIndex = 6;
            this.button1.Text = "Validar";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Validar_click);
            // 
            // log
            // 
            this.log.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.log.BackColor = System.Drawing.SystemColors.ScrollBar;
            this.log.FormattingEnabled = true;
            this.log.ItemHeight = 16;
            this.log.Location = new System.Drawing.Point(355, 110);
            this.log.Name = "log";
            this.log.Size = new System.Drawing.Size(766, 356);
            this.log.TabIndex = 0;
            this.log.Visible = false;
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
            "Legado Movimentacao"});
            this.listBox1.Location = new System.Drawing.Point(15, 114);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(321, 24);
            this.listBox1.TabIndex = 5;
            // 
            // btnSelectFile
            // 
            this.btnSelectFile.Location = new System.Drawing.Point(15, 179);
            this.btnSelectFile.Name = "btnSelectFile";
            this.btnSelectFile.Size = new System.Drawing.Size(84, 23);
            this.btnSelectFile.TabIndex = 1;
            this.btnSelectFile.Text = "Escolher";
            this.btnSelectFile.UseVisualStyleBackColor = true;
            this.btnSelectFile.Click += new System.EventHandler(this.Escolher_click);
            // 
            // txtFilePath
            // 
            this.txtFilePath.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtFilePath.ForeColor = System.Drawing.SystemColors.InactiveCaption;
            this.txtFilePath.Location = new System.Drawing.Point(105, 179);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.ReadOnly = true;
            this.txtFilePath.Size = new System.Drawing.Size(231, 22);
            this.txtFilePath.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label2.Location = new System.Drawing.Point(11, 92);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(71, 20);
            this.label2.TabIndex = 0;
            this.label2.Text = "Layout:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label1.Location = new System.Drawing.Point(12, 155);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(78, 20);
            this.label1.TabIndex = 0;
            this.label1.Text = "Arquivo:";
            // 
            // panelmid
            // 
            this.panelmid.Controls.Add(this.LC);
            this.panelmid.Controls.Add(this.labellog);
            this.panelmid.Controls.Add(this.grid);
            this.panelmid.Controls.Add(this.depuracao);
            this.panelmid.Controls.Add(this.pictureBox1);
            this.panelmid.Controls.Add(this.possuiCabecalho);
            this.panelmid.Controls.Add(this.progressBar);
            this.panelmid.Controls.Add(this.excel);
            this.panelmid.Controls.Add(this.log);
            this.panelmid.Controls.Add(this.label1);
            this.panelmid.Controls.Add(this.button1);
            this.panelmid.Controls.Add(this.listBox1);
            this.panelmid.Controls.Add(this.txtFilePath);
            this.panelmid.Controls.Add(this.label2);
            this.panelmid.Controls.Add(this.btnSelectFile);
            this.panelmid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelmid.Location = new System.Drawing.Point(0, 0);
            this.panelmid.Name = "panelmid";
            this.panelmid.Size = new System.Drawing.Size(1133, 513);
            this.panelmid.TabIndex = 6;
            // 
            // depuracao
            // 
            this.depuracao.AutoSize = true;
            this.depuracao.Location = new System.Drawing.Point(22, 393);
            this.depuracao.Name = "depuracao";
            this.depuracao.Size = new System.Drawing.Size(75, 16);
            this.depuracao.TabIndex = 10;
            this.depuracao.Text = "Depuração";
            this.depuracao.Visible = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(12, 14);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(223, 64);
            this.pictureBox1.TabIndex = 9;
            this.pictureBox1.TabStop = false;
            // 
            // possuiCabecalho
            // 
            this.possuiCabecalho.AutoSize = true;
            this.possuiCabecalho.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.possuiCabecalho.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.possuiCabecalho.Location = new System.Drawing.Point(16, 207);
            this.possuiCabecalho.Name = "possuiCabecalho";
            this.possuiCabecalho.Size = new System.Drawing.Size(109, 22);
            this.possuiCabecalho.TabIndex = 8;
            this.possuiCabecalho.Text = "Cabeçalho?";
            this.possuiCabecalho.UseVisualStyleBackColor = true;
            // 
            // progressBar
            // 
            this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar.Location = new System.Drawing.Point(355, 471);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(767, 30);
            this.progressBar.TabIndex = 7;
            this.progressBar.Visible = false;
            // 
            // grid
            // 
            this.grid.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.grid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grid.Location = new System.Drawing.Point(355, 37);
            this.grid.Name = "grid";
            this.grid.RowHeadersWidth = 51;
            this.grid.RowTemplate.Height = 24;
            this.grid.Size = new System.Drawing.Size(766, 428);
            this.grid.TabIndex = 11;
            // 
            // LC
            // 
            this.LC.AutoSize = true;
            this.LC.BackColor = System.Drawing.SystemColors.WindowFrame;
            this.LC.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LC.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
            this.LC.Location = new System.Drawing.Point(181, 204);
            this.LC.Name = "LC";
            this.LC.Size = new System.Drawing.Size(35, 20);
            this.LC.TabIndex = 12;
            this.LC.Text = "csv";
            this.LC.Visible = false;
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.WindowFrame;
            this.ClientSize = new System.Drawing.Size(1133, 513);
            this.Controls.Add(this.panelmid);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Main";
            this.Text = "Validar CSV";
            this.panelmid.ResumeLayout(false);
            this.panelmid.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.grid)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button excel;
        private System.Windows.Forms.Label labellog;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ListBox log;
        private System.Windows.Forms.ComboBox listBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnSelectFile;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panelmid;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.CheckBox possuiCabecalho;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label depuracao;
        private System.Windows.Forms.DataGridView grid;
        private System.Windows.Forms.Label LC;
    }
}

