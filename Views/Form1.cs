using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace MultasLectura.Views
{
    public partial class Form1 : Form
    {
        private List<ExcelFileInfo> excelFiles = new List<ExcelFileInfo>();
        private List<string> requiredColumns = new List<string> { "Nombre", "Apellido", "Edad", "Email", "Ciudad" };

        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            this.AllowDrop = true;
            this.DragEnter += Form1_DragEnter;
            this.DragDrop += Form1_DragDrop;

            listBoxFiles.DrawMode = DrawMode.OwnerDrawFixed;
            listBoxFiles.DrawItem += ListBoxFiles_DrawItem;

            RefreshColumnsList();
        }

        #region Drag & Drop y selección
        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            AddFiles(files);
        }

        private void btnSelectFiles_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls",
                Multiselect = true
            };
            if (ofd.ShowDialog() == DialogResult.OK)
                AddFiles(ofd.FileNames);
        }

        private void AddFiles(string[] files)
        {
            foreach (var file in files)
            {
                if (!File.Exists(file) || (!file.EndsWith(".xls") && !file.EndsWith(".xlsx")))
                    continue;

                if (!excelFiles.Any(x => x.FilePath == file))
                {
                    var info = ValidateFileColumns(file);
                    excelFiles.Add(info);
                    listBoxFiles.Items.Add(info);
                    AppendLog($"Archivo agregado: {Path.GetFileName(file)} ({(info.IsValid ? "OK" : "Faltan columnas")})");
                }
            }
        }
        #endregion

        #region Validación columnas
        private ExcelFileInfo ValidateFileColumns(string file)
        {
            var info = new ExcelFileInfo { FilePath = file, MissingColumns = new List<string>() };
            using (var package = new ExcelPackage(new FileInfo(file)))
            {
                var ws = package.Workbook.Worksheets.First();
                var headers = new List<string>();
                for (int col = 1; col <= ws.Dimension.End.Column; col++)
                    headers.Add(ws.Cells[1, col].Text.Trim());

                foreach (var column in requiredColumns)
                {
                    if (!headers.Contains(column))
                        info.MissingColumns.Add(column);
                }
            }
            info.IsValid = info.MissingColumns.Count == 0;
            return info;
        }
        #endregion

        #region ListBox visual
        private void ListBoxFiles_DrawItem(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();
            if (e.Index < 0) return;

            var info = (ExcelFileInfo)listBoxFiles.Items[e.Index];
            Color color = info.IsValid ? Color.Green : Color.Red;

            using (SolidBrush brush = new SolidBrush(color))
            {
                e.Graphics.DrawString(Path.GetFileName(info.FilePath), e.Font, brush, e.Bounds);
            }
            e.DrawFocusRectangle();
        }
        #endregion

        #region Columnas dinámicas
        private void RefreshColumnsList()
        {
            listBoxColumns.Items.Clear();
            foreach (var col in requiredColumns)
                listBoxColumns.Items.Add(col);
        }

        private void btnAddColumn_Click(object sender, EventArgs e)
        {
            string newCol = textBoxNewColumn.Text.Trim();
            if (!string.IsNullOrEmpty(newCol) && !requiredColumns.Contains(newCol))
            {
                requiredColumns.Add(newCol);
                RefreshColumnsList();
                textBoxNewColumn.Clear();
                AppendLog($"Columna agregada: {newCol}");
            }
        }

        private void btnRemoveColumn_Click(object sender, EventArgs e)
        {
            if (listBoxColumns.SelectedItem != null)
            {
                string col = listBoxColumns.SelectedItem.ToString();
                requiredColumns.Remove(col);
                RefreshColumnsList();
                AppendLog($"Columna eliminada: {col}");
            }
        }
        #endregion

        #region Archivos dinámicos
        private void btnRemoveFile_Click(object sender, EventArgs e)
        {
            if (listBoxFiles.SelectedItem != null)
            {
                var fileInfo = (ExcelFileInfo)listBoxFiles.SelectedItem;
                excelFiles.Remove(fileInfo);
                listBoxFiles.Items.Remove(fileInfo);
                AppendLog($"Archivo eliminado: {Path.GetFileName(fileInfo.FilePath)}");
            }
        }
        #endregion

        #region Procesamiento
        private void btnProcess_Click(object sender, EventArgs e)
        {
            if (excelFiles.Count < 5)
            {
                MessageBox.Show("Debe seleccionar al menos 5 archivos.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            AppendLog("Iniciando procesamiento...");

            var overallMissingColumns = new HashSet<string>(requiredColumns);

            foreach (var fileInfo in excelFiles)
            {
                using (var package = new ExcelPackage(new FileInfo(fileInfo.FilePath)))
                {
                    var ws = package.Workbook.Worksheets.First();
                    var headers = new List<string>();
                    for (int col = 1; col <= ws.Dimension.End.Column; col++)
                        headers.Add(ws.Cells[1, col].Text.Trim());

                    foreach (var column in requiredColumns)
                    {
                        if (headers.Contains(column))
                        {
                            overallMissingColumns.Remove(column);
                            CallFunctionByColumn(column, ws);
                            AppendLog($"{Path.GetFileName(fileInfo.FilePath)} -> Procesada columna: {column}");
                        }
                    }
                }
            }

            if (overallMissingColumns.Count > 0)
            {
                string missingCols = string.Join(", ", overallMissingColumns);
                AppendLog($"ERROR: No se encontraron las siguientes columnas en ningún archivo: {missingCols}");
                MessageBox.Show("No se encontraron las siguientes columnas en ningún archivo: " + missingCols, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            AppendLog("Procesamiento finalizado correctamente.");
            MessageBox.Show("Procesamiento finalizado correctamente.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void CallFunctionByColumn(string columnName, ExcelWorksheet ws)
        {
            switch (columnName)
            {
                case "Nombre": ProcessNombre(ws); break;
                case "Apellido": ProcessApellido(ws); break;
                case "Edad": ProcessEdad(ws); break;
                case "Email": ProcessEmail(ws); break;
                case "Ciudad": ProcessCiudad(ws); break;
                default: ProcessCustomColumn(columnName, ws); break;
            }
        }

        private void ProcessNombre(ExcelWorksheet ws) { /* Código */ }
        private void ProcessApellido(ExcelWorksheet ws) { /* Código */ }
        private void ProcessEdad(ExcelWorksheet ws) { /* Código */ }
        private void ProcessEmail(ExcelWorksheet ws) { /* Código */ }
        private void ProcessCiudad(ExcelWorksheet ws) { /* Código */ }
        private void ProcessCustomColumn(string columnName, ExcelWorksheet ws) { /* Código */ }
        #endregion

        #region Log en tiempo real
        private void AppendLog(string message)
        {
            textBoxLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
        }
        #endregion
    }

    public class ExcelFileInfo
    {
        public string FilePath { get; set; }
        public bool IsValid { get; set; }
        public List<string> MissingColumns { get; set; }
        public override string ToString() => Path.GetFileName(FilePath);
    }
}
