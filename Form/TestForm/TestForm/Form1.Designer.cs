namespace TestForm
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.databseToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.connnectToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.queryToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.addMockingDataToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.clearDatabaseToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.crystalReportToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exportPDFToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.result = new System.Windows.Forms.DataGridView();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.result)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.databseToolStripMenuItem,
            this.crystalReportToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(800, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "File";
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(93, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            // 
            // databseToolStripMenuItem
            // 
            this.databseToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.connnectToolStripMenuItem,
            this.queryToolStripMenuItem,
            this.addMockingDataToolStripMenuItem,
            this.clearDatabaseToolStripMenuItem});
            this.databseToolStripMenuItem.Name = "databseToolStripMenuItem";
            this.databseToolStripMenuItem.Size = new System.Drawing.Size(61, 20);
            this.databseToolStripMenuItem.Text = "Databse";
            // 
            // connnectToolStripMenuItem
            // 
            this.connnectToolStripMenuItem.Name = "connnectToolStripMenuItem";
            this.connnectToolStripMenuItem.Size = new System.Drawing.Size(173, 22);
            this.connnectToolStripMenuItem.Text = "Connect";
            this.connnectToolStripMenuItem.Click += new System.EventHandler(this.connnectToolStripMenuItem_Click);
            // 
            // queryToolStripMenuItem
            // 
            this.queryToolStripMenuItem.Name = "queryToolStripMenuItem";
            this.queryToolStripMenuItem.Size = new System.Drawing.Size(173, 22);
            this.queryToolStripMenuItem.Text = "Query";
            this.queryToolStripMenuItem.Click += new System.EventHandler(this.queryToolStripMenuItem_Click);
            // 
            // addMockingDataToolStripMenuItem
            // 
            this.addMockingDataToolStripMenuItem.Name = "addMockingDataToolStripMenuItem";
            this.addMockingDataToolStripMenuItem.Size = new System.Drawing.Size(173, 22);
            this.addMockingDataToolStripMenuItem.Text = "Add Mocking Data";
            this.addMockingDataToolStripMenuItem.Click += new System.EventHandler(this.addMockingDataToolStripMenuItem_Click);
            // 
            // clearDatabaseToolStripMenuItem
            // 
            this.clearDatabaseToolStripMenuItem.Name = "clearDatabaseToolStripMenuItem";
            this.clearDatabaseToolStripMenuItem.Size = new System.Drawing.Size(173, 22);
            this.clearDatabaseToolStripMenuItem.Text = "Clear Database";
            this.clearDatabaseToolStripMenuItem.Click += new System.EventHandler(this.clearDatabaseToolStripMenuItem_Click);
            // 
            // crystalReportToolStripMenuItem
            // 
            this.crystalReportToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.exportPDFToolStripMenuItem});
            this.crystalReportToolStripMenuItem.Name = "crystalReportToolStripMenuItem";
            this.crystalReportToolStripMenuItem.Size = new System.Drawing.Size(93, 20);
            this.crystalReportToolStripMenuItem.Text = "Crystal Report";
            // 
            // exportPDFToolStripMenuItem
            // 
            this.exportPDFToolStripMenuItem.Name = "exportPDFToolStripMenuItem";
            this.exportPDFToolStripMenuItem.Size = new System.Drawing.Size(132, 22);
            this.exportPDFToolStripMenuItem.Text = "Export PDF";
            this.exportPDFToolStripMenuItem.Click += new System.EventHandler(this.exportPDFToolStripMenuItem_Click);
            // 
            // result
            // 
            this.result.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.result.Location = new System.Drawing.Point(0, 27);
            this.result.Name = "result";
            this.result.Size = new System.Drawing.Size(800, 422);
            this.result.TabIndex = 1;
            this.result.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.result);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "Form1";
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.result)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem databseToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem connnectToolStripMenuItem;
        private System.Windows.Forms.DataGridView result;
        private System.Windows.Forms.ToolStripMenuItem queryToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem addMockingDataToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem crystalReportToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exportPDFToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem clearDatabaseToolStripMenuItem;
    }
}

