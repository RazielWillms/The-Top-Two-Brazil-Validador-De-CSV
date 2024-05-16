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
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.possuiCabecalho = new System.Windows.Forms.CheckBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.panelmid.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // excel
            // 
            this.excel.Location = new System.Drawing.Point(129, 257);
            this.excel.Name = "excel";
            this.excel.Size = new System.Drawing.Size(107, 28);
            this.excel.TabIndex = 1;
            this.excel.Text = "EXPORTAR";
            this.excel.UseVisualStyleBackColor = true;
            this.excel.Click += new System.EventHandler(this.Exportar_click);
            // 
            // labellog
            // 
            this.labellog.AutoSize = true;
            this.labellog.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labellog.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.labellog.Location = new System.Drawing.Point(9, 303);
            this.labellog.Name = "labellog";
            this.labellog.Size = new System.Drawing.Size(70, 16);
            this.labellog.TabIndex = 5;
            this.labellog.Text = "Registro:";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(15, 257);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(108, 28);
            this.button1.TabIndex = 6;
            this.button1.Text = "VALIDAR";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Validar_click);
            // 
            // log
            // 
            this.log.BackColor = System.Drawing.SystemColors.ScrollBar;
            this.log.FormattingEnabled = true;
            this.log.ItemHeight = 16;
            this.log.Location = new System.Drawing.Point(14, 322);
            this.log.Name = "log";
            this.log.Size = new System.Drawing.Size(426, 292);
            this.log.TabIndex = 0;
            // 
            // listBox1
            // 
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
            this.listBox1.Location = new System.Drawing.Point(15, 101);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(426, 24);
            this.listBox1.TabIndex = 5;
            // 
            // btnSelectFile
            // 
            this.btnSelectFile.Location = new System.Drawing.Point(15, 185);
            this.btnSelectFile.Name = "btnSelectFile";
            this.btnSelectFile.Size = new System.Drawing.Size(107, 27);
            this.btnSelectFile.TabIndex = 1;
            this.btnSelectFile.Text = "ESCOLHER";
            this.btnSelectFile.UseVisualStyleBackColor = true;
            this.btnSelectFile.Click += new System.EventHandler(this.Escolher_click);
            // 
            // txtFilePath
            // 
            this.txtFilePath.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtFilePath.ForeColor = System.Drawing.SystemColors.InactiveCaption;
            this.txtFilePath.Location = new System.Drawing.Point(15, 157);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.ReadOnly = true;
            this.txtFilePath.Size = new System.Drawing.Size(425, 22);
            this.txtFilePath.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label2.Location = new System.Drawing.Point(12, 82);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(140, 16);
            this.label2.TabIndex = 0;
            this.label2.Text = "Selecione o layout:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.label1.Location = new System.Drawing.Point(9, 138);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(150, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "Selecione o arquivo:";
            // 
            // panelmid
            // 
            this.panelmid.Controls.Add(this.pictureBox1);
            this.panelmid.Controls.Add(this.possuiCabecalho);
            this.panelmid.Controls.Add(this.progressBar);
            this.panelmid.Controls.Add(this.labellog);
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
            this.panelmid.Size = new System.Drawing.Size(452, 664);
            this.panelmid.TabIndex = 6;
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(15, 623);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(425, 30);
            this.progressBar.TabIndex = 7;
            this.progressBar.Visible = false;
            // 
            // possuiCabecalho
            // 
            this.possuiCabecalho.AutoSize = true;
            this.possuiCabecalho.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.possuiCabecalho.Location = new System.Drawing.Point(15, 218);
            this.possuiCabecalho.Name = "possuiCabecalho";
            this.possuiCabecalho.Size = new System.Drawing.Size(131, 20);
            this.possuiCabecalho.TabIndex = 8;
            this.possuiCabecalho.Text = "Tem cabeçalho?";
            this.possuiCabecalho.UseVisualStyleBackColor = true;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(12, 14);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(202, 56);
            this.pictureBox1.TabIndex = 9;
            this.pictureBox1.TabStop = false;
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.WindowFrame;
            this.ClientSize = new System.Drawing.Size(452, 664);
            this.Controls.Add(this.panelmid);
            this.Name = "Main";
            this.Text = "The Top Two Brazil Csv Validator";
            this.panelmid.ResumeLayout(false);
            this.panelmid.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
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
    }
}

