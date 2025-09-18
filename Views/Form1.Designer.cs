namespace MultasLectura.Views
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
       // private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        /* protected override void Dispose(bool disposing)
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
             this.components = new System.ComponentModel.Container();
             this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
             this.ClientSize = new System.Drawing.Size(800, 450);
             this.Text = "Form1";
         }

         #endregion*/
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.ListBox listBoxFiles;
        private System.Windows.Forms.Button btnSelectFiles;
        private System.Windows.Forms.Button btnRemoveFile;
        private System.Windows.Forms.ListBox listBoxColumns;
        private System.Windows.Forms.TextBox textBoxNewColumn;
        private System.Windows.Forms.Button btnAddColumn;
        private System.Windows.Forms.Button btnRemoveColumn;
        private System.Windows.Forms.Button btnProcess;
        private System.Windows.Forms.TextBox textBoxLog;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null)) components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.listBoxFiles = new System.Windows.Forms.ListBox();
            this.btnSelectFiles = new System.Windows.Forms.Button();
            this.btnRemoveFile = new System.Windows.Forms.Button();
            this.listBoxColumns = new System.Windows.Forms.ListBox();
            this.textBoxNewColumn = new System.Windows.Forms.TextBox();
            this.btnAddColumn = new System.Windows.Forms.Button();
            this.btnRemoveColumn = new System.Windows.Forms.Button();
            this.btnProcess = new System.Windows.Forms.Button();
            this.textBoxLog = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // listBoxFiles
            this.listBoxFiles.FormattingEnabled = true;
            this.listBoxFiles.Location = new System.Drawing.Point(12, 12);
            this.listBoxFiles.Size = new System.Drawing.Size(400, 150);
            // 
            // btnSelectFiles
            this.btnSelectFiles.Location = new System.Drawing.Point(420, 12);
            this.btnSelectFiles.Size = new System.Drawing.Size(120, 30);
            this.btnSelectFiles.Text = "Agregar Archivos";
            this.btnSelectFiles.Click += new System.EventHandler(this.btnSelectFiles_Click);
            // 
            // btnRemoveFile
            this.btnRemoveFile.Location = new System.Drawing.Point(420, 48);
            this.btnRemoveFile.Size = new System.Drawing.Size(120, 30);
            this.btnRemoveFile.Text = "Eliminar Archivo";
            this.btnRemoveFile.Click += new System.EventHandler(this.btnRemoveFile_Click);
            // 
            // listBoxColumns
            this.listBoxColumns.FormattingEnabled = true;
            this.listBoxColumns.Location = new System.Drawing.Point(12, 180);
            this.listBoxColumns.Size = new System.Drawing.Size(200, 150);
            // 
            // textBoxNewColumn
            this.textBoxNewColumn.Location = new System.Drawing.Point(220, 180);
            this.textBoxNewColumn.Size = new System.Drawing.Size(120, 22);
            // 
            // btnAddColumn
            this.btnAddColumn.Location = new System.Drawing.Point(220, 210);
            this.btnAddColumn.Size = new System.Drawing.Size(120, 30);
            this.btnAddColumn.Text = "Agregar Columna";
            this.btnAddColumn.Click += new System.EventHandler(this.btnAddColumn_Click);
            // 
            // btnRemoveColumn
            this.btnRemoveColumn.Location = new System.Drawing.Point(220, 250);
            this.btnRemoveColumn.Size = new System.Drawing.Size(120, 30);
            this.btnRemoveColumn.Text = "Eliminar Columna";
            this.btnRemoveColumn.Click += new System.EventHandler(this.btnRemoveColumn_Click);
            // 
            // btnProcess
            this.btnProcess.Location = new System.Drawing.Point(420, 180);
            this.btnProcess.Size = new System.Drawing.Size(120, 60);
            this.btnProcess.Text = "Procesar";
            this.btnProcess.Click += new System.EventHandler(this.btnProcess_Click);
            // 
            // textBoxLog
            this.textBoxLog.Location = new System.Drawing.Point(12, 340);
            this.textBoxLog.Multiline = true;
            this.textBoxLog.ScrollBars = ScrollBars.Vertical;
            this.textBoxLog.Size = new System.Drawing.Size(528, 200);
            this.textBoxLog.ReadOnly = true;
            this.textBoxLog.BackColor = SystemColors.Window;
            // 
            // Form1
            this.ClientSize = new System.Drawing.Size(552, 550);
            this.Controls.Add(this.listBoxFiles);
            this.Controls.Add(this.btnSelectFiles);
            this.Controls.Add(this.btnRemoveFile);
            this.Controls.Add(this.listBoxColumns);
            this.Controls.Add(this.textBoxNewColumn);
            this.Controls.Add(this.btnAddColumn);
            this.Controls.Add(this.btnRemoveColumn);
            this.Controls.Add(this.btnProcess);
            this.Controls.Add(this.textBoxLog);
            this.Text = "Excel Processor";
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}