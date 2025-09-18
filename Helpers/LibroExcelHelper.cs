using Aspose.Cells;
using ExcelDataReader;
using MathNet.Numerics.Distributions;
using MultasLectura.Enums;
using MultasLectura.Models;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Net.WebRequestMethods;
using File = System.IO.File;

namespace MultasLectura.Helpers
{
    public class LibroExcelHelper
    {
        public static void AutoFitColumnsWithLimits(
        ExcelWorksheet worksheet,
        double minWidth = 10,
        double maxWidth = 50,
        double extraPadding = 2)
        {
            if (worksheet.Dimension == null) return;

            // Ajustar automáticamente todas las columnas usadas
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

            // Aplicar límites de ancho
            for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
            {
                double currentWidth = worksheet.Column(col).Width + extraPadding;

                if (currentWidth < minWidth)
                    worksheet.Column(col).Width = minWidth;
                else if (currentWidth > maxWidth)
                    worksheet.Column(col).Width = maxWidth;
                else
                    worksheet.Column(col).Width = currentWidth;
            }
        }

        static public void IniciarProcesoCarga(TextBox txt)
        {
            string filePath = CargarLibroExcel();

            if (string.IsNullOrEmpty(filePath))
            {
                txt.Text = string.Empty;
                MostrarMensaje("Ocurrió un error al intentar cargar el archivo. Por favor inténtelo nuevamente", true);
            }
            else
            {
                txt.Text = filePath;
            }
        }

        static public string CargarLibroExcel()
        {
            try
            {
                OpenFileDialog openFileDialog = new()
                {
                    InitialDirectory = "c:\\",
                    Filter = "Archivos Excel (*.xls)|*.xls|Archivos Excel (*.xlsx)|*.xlsx",
                    FilterIndex = 1,
                    RestoreDirectory = true
                };

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return openFileDialog.FileName;
                }
                else
                {
                    return "";
                }
            }
            catch
            {
                return "";
            }

        }

        static public string ObtenerValorPorClave(List<Dictionary<string, string>> lista, string clave)
        {
            foreach (var diccionario in lista)
            {
                if (diccionario.ContainsKey(clave))
                {
                    return diccionario[clave];
                }
            }
            return ""; 
        }

        static public List<Dictionary<string, string>> ProcesarPathArchivos(
            List<Dictionary<string, string>> rutas,
            string nombreXDefecto
            )
        {
            List<Dictionary<string, string>> rutasFinales = new();

            try
            {
                for (int i = 0; i < rutas.Count; i++)
                {
                    var clave = rutas[i].Keys.FirstOrDefault();

                    string rutaValida = LibroExcelHelper.ValidarFormato(rutas[i][clave!]);

                    if (!string.IsNullOrEmpty(rutaValida))
                    {
                        Dictionary<string, string> rutaNueva = new()
                        {
                            { clave!, rutaValida }
                        };

                        rutasFinales.Add(rutaNueva);
                          
                    }
                }


                if (rutas.Count != rutasFinales.Count)
                {
                    LibroExcelHelper.MostrarMensaje("Error al cargar los archivos. Intente nuevamente.", true);
                } else
                {
                    string rutaArchivo = LibroExcelHelper.DialogoGuardarArchivo(nombreXDefecto);

                    if (string.IsNullOrEmpty(rutaArchivo))
                    {
                        LibroExcelHelper.MostrarMensaje("Tarea cancelada por el usuario.", true);
                    }
                    else
                    {
                        Dictionary<string, string> rutaNueva = new()
                        {
                            { RutaArchivo.Guardar.ToString(), rutaArchivo }
                        };

                        rutasFinales.Add(rutaNueva);
                    }
                }


            }
            catch (Exception e)
            {
                LibroExcelHelper.MostrarMensaje(e.Message, true);
            }

            return rutasFinales;


        }

        static public void AsignarValorFormulaACelda<V>(ExcelWorksheet hoja, string celda, V valor, TipoOpCelda tipo)
        {
            ExcelRange rango = hoja.Cells[$"{celda}"];

            switch (tipo)
            {
                case TipoOpCelda.Value:
                    rango.Value = valor;
                    break;
                case TipoOpCelda.Formula:
                    rango.Formula = valor!.ToString();
                    break;
            }
        }

        static public void ColorFondoLetra(ExcelWorksheet hoja, char letraCelda, int numCelda1, ColorModel color)
        {
            hoja.Cells[$"{letraCelda.ToString().ToUpper()}{numCelda1}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
           hoja.Cells[$"{letraCelda.ToString().ToUpper()}{numCelda1}"].Style.Fill.BackgroundColor.SetColor(color.Fondo);
            hoja.Cells[$"{letraCelda.ToString().ToUpper()}{numCelda1}"].Style.Font.Color.SetColor(color.Letra);
        }

        static public void FormatoMergeCelda(ExcelWorksheet hoja, string rango)
        {
            hoja.Cells[rango].Merge = true;
        }

        static public void CentrarContenidoCelda(ExcelWorksheet hoja, string rango)
        {
            hoja.Cells[rango].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            hoja.Cells[rango].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }

        static public void AplicarBordeFinoARango(ExcelRangeBase rango)
        {
            rango.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            rango.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            rango.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            rango.Style.Border.Right.Style = ExcelBorderStyle.Thin;
        }

        static public void AplicarBordeGruesoARango(ExcelRangeBase rango)
        {
            rango.Style.Border.Top.Style = ExcelBorderStyle.Thick;
            rango.Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            rango.Style.Border.Left.Style = ExcelBorderStyle.Thick;
            rango.Style.Border.Right.Style = ExcelBorderStyle.Thick;
        }

        static public void AplicarBordeParcialARango(ExcelWorksheet hoja, string rango, TipoBordeParcial tipo)
        {
            ExcelRange rangoFinal = hoja.Cells[rango];

            switch (tipo)
            {
                case TipoBordeParcial.SupDerIzq:
                    {
                        rangoFinal.Style.Border.Top.Style = ExcelBorderStyle.Thick;
                        rangoFinal.Style.Border.Left.Style = ExcelBorderStyle.Thick;
                        rangoFinal.Style.Border.Right.Style = ExcelBorderStyle.Thick;
                    } 
                   
                    break;
                case TipoBordeParcial.CentroDerIzq:
                    {
                        rangoFinal.Style.Border.Left.Style = ExcelBorderStyle.Thick;
                        rangoFinal.Style.Border.Right.Style = ExcelBorderStyle.Thick;
                    }
                    break;
                case TipoBordeParcial.InfDerIzq:
                    {
                        rangoFinal.Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                        rangoFinal.Style.Border.Left.Style = ExcelBorderStyle.Thick;
                        rangoFinal.Style.Border.Right.Style = ExcelBorderStyle.Thick;
                    }
                    break;
            }
           
        }

        static public void FormatoMoneda(ExcelRange rango)
        {
            using ExcelRange celda = rango;
            celda.Style.Numberformat.Format = "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)";
        }

        static public void FormatoPorcentaje(ExcelRange rango)
        {
            rango.Style.Numberformat.Format = "0.00%";
        }

        static public void FormatoNegrita(ExcelRange rango)
        {
            rango.Style.Font.Bold = true;
        }

        static public void FondoSolido(ExcelRange rango, Color color)
        {
            rango.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rango.Style.Fill.BackgroundColor.SetColor(color);
        }

        static public int ObtenerNumeroColumna(ExcelWorksheet hoja, string encabezado)
        {
            int contadorColumnas = hoja.Dimension.End.Column;
            for (int col = 1; col <= contadorColumnas; col++)
            {
                if (hoja.Cells[1, col].Text == encabezado)
                {
                    return col;
                }
            }
            return -1;
        }


        static public string ValidarFormato(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string fileExtension = Path.GetExtension(filePath);

            if (fileExtension.ToLower() == ".xlsx")
            {
                return filePath;
            }
            else if (fileExtension.ToLower() == ".xls")
            {
                
                string convertedFilePath = ConvertirXlsAXlsx(filePath);
                if (!string.IsNullOrEmpty(convertedFilePath))
                {
                    return convertedFilePath;
                }
                else
                {
                    return string.Empty;
                }
            }
            else
            {
                return string.Empty;
            }
        }

       static private string ConvertirXlsAXlsx(string inputXlsPath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string outputXlsxPath = Path.Combine(
                Path.GetDirectoryName(inputXlsPath),
                Path.GetFileNameWithoutExtension(inputXlsPath) + "_convertido.xlsx");

            using (var stream = File.Open(inputXlsPath, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                var conf = new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = false
                    }
                };

                using (var result = reader.AsDataSet(conf))
                using (var package = new ExcelPackage())
                {
                    foreach (DataTable table in result.Tables)
                    {
                        var hoja = package.Workbook.Worksheets.Add(table.TableName);

                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            for (int j = 0; j < table.Columns.Count; j++)
                            {
                                hoja.Cells[i + 1, j + 1].Value = table.Rows[i][j];
                            }
                        }
                    }

                    package.SaveAs(new FileInfo(outputXlsxPath));
                }
            }

            return outputXlsxPath;
        }

        static private string ConvertirXlsAXlsx2(string xlsFilePath)
        {
           
            try
            {

                // Cargar archivo .xls
                using FileStream fs = new(xlsFilePath, FileMode.Open, FileAccess.Read);
                HSSFWorkbook hssfwb = new(fs); // Crear instancia de libro .xls
                XSSFWorkbook workbook = new(); // Crear instancia de libro .xlsx


                // Copiar hojas de .xls a .xlsx
                for (int i = 0; i < hssfwb.NumberOfSheets; i++)
                {
                    ISheet sheet = hssfwb.GetSheetAt(i);
                    XSSFSheet newSheet = (XSSFSheet)workbook.CreateSheet(sheet.SheetName);

                    // Copiar filas y celdas
                    for (int j = 0; j <= sheet.LastRowNum; j++)
                    {
                        IRow row = sheet.GetRow(j);
                        XSSFRow newRow = (XSSFRow)newSheet.CreateRow(j);

                        if (row != null)
                        {
                            for (int k = 0; k < row.LastCellNum; k++)
                            {
                                ICell cell = row.GetCell(k);
                                if (cell != null)
                                {
                                    XSSFCell newCell = (XSSFCell)newRow.CreateCell(k);
                                    newCell.SetCellValue(cell.ToString());
                                }
                            }
                        }
                    }
                }

                var partesPath = xlsFilePath.Split('\\');
                string pathGuardarConversion = "";

                foreach (string parte in partesPath)
                {
                    if (partesPath[partesPath.Length - 1] != parte)
                    {
                        pathGuardarConversion += parte + '\\';
                    }
                }

                string nombreOriginal = Path.GetFileNameWithoutExtension(xlsFilePath);

                //   Random random = new Random();

                // Generar un número aleatorio entre 1 y 100 (ambos inclusive)
                //  int numeroAleatorio = random.Next(1, 101);

                //  pathGuardarConversion += $"libro-conversion{numeroAleatorio}.xlsx";
                // DateTime.Now.ToString("yyyyMMdd") + "-" + DateTime.Now.ToString("HHmmss")
                pathGuardarConversion += $"{nombreOriginal}_{DateTime.Now.ToString("yyyyMMdd")}{DateTime.Now.ToString("HHmmss")}.xlsx";

                //  MessageBox.Show(partesPath[partesPath.Length - 1]);

                // Guardar como .xlsx
                using (FileStream fileOut = new FileStream(pathGuardarConversion, FileMode.Create))
                {
                    workbook.Write(fileOut);
                }

                return pathGuardarConversion;


            }
            catch
            {
                return string.Empty;
            }
        }

        

        static public void ConvertirTextoANumero(ExcelRange rango)
        {
            foreach (var celda in rango)
            {
                if (double.TryParse(celda.Value?.ToString(), out double valor))
                {
                    // Asignar el valor convertido de vuelta a la celda
                    celda.Value = valor;
                }
            }
        }

        static public string DialogoGuardarArchivo(string nombreXDefecto)
        {
            string rutaArchivo = string.Empty;

            /* SaveFileDialog guardar = new SaveFileDialog();
            
            guardar.Filter = "Archivos Excel (*.xlsx)|*.xlsx";
            guardar.FilterIndex = 1;
            guardar.RestoreDirectory = true;

            //guardar.FileName = "+ Calidad_Lectura.xlsx";
            guardar.FileName = nombreXDefecto;

            if (guardar.ShowDialog() == DialogResult.OK)
            {
                rutaArchivo = guardar.FileName;
            }

            return rutaArchivo;*/

            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Seleccioná la carpeta donde guardar los archivos";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    rutaArchivo = dialog.SelectedPath;
                }


                //MessageBox.Show(rutaArchivo);

                return rutaArchivo;
            }




                }

                static public void MostrarMensaje(string mensaje, bool esError)
        {
            MessageBox.Show(mensaje, esError ? "ERROR" : "ATENCIÓN", MessageBoxButtons.OK, esError ? MessageBoxIcon.Error : MessageBoxIcon.Warning);
        }

        static public int SumarColumnaInt(ExcelWorksheet hoja, int colASumar, int filaInicial)
        {
            int totalSuma = 0;

            // int columnToSum = 5; // Columna que quieres sumar (por ejemplo, A=1, B=2, etc.)
            //int startRow = 2; // Fila inicial donde comienzan los datos (puede variar según el archivo)
            int filaFinal = hoja.Dimension.End.Row; // Última fila con datos en la hoja

            // double total = 0;

            for (int fila = filaInicial; fila <= filaFinal; fila++)
            {
                var celda = hoja.Cells[fila, colASumar].Value;
                if (celda != null)
                {
                    int datoCelda;
                    if (int.TryParse(celda.ToString(), out datoCelda))
                    {
                        totalSuma += datoCelda;
                    }
                    // Si los datos no son numéricos, se pueden ignorar o manejar según sea necesario
                }
            }

            return totalSuma;
        }


    }





    public static class ExcelConverter
   {
        /// <summary>
        /// Convierte un archivo .xls a .xlsx utilizando NPOI.
        /// </summary>
        public static void ConvertXlsToXlsx(string xlsPath, string xlsxPath)
        {
            if (!File.Exists(xlsPath))
                throw new FileNotFoundException("El archivo .xls no existe.", xlsPath);

            using (var xlsStream = new FileStream(xlsPath, FileMode.Open, FileAccess.Read))
            {
                var hssfWorkbook = new HSSFWorkbook(xlsStream);
                var xssfWorkbook = new XSSFWorkbook();

                for (int i = 0; i < hssfWorkbook.NumberOfSheets; i++)
                {
                    ISheet sourceSheet = hssfWorkbook.GetSheetAt(i);
                    ISheet newSheet = xssfWorkbook.CreateSheet(sourceSheet.SheetName);

                    for (int row = 0; row <= sourceSheet.LastRowNum; row++)
                    {
                        IRow sourceRow = sourceSheet.GetRow(row);
                        if (sourceRow == null) continue;

                        IRow newRow = newSheet.CreateRow(row);

                        for (int col = 0; col < sourceRow.LastCellNum; col++)
                        {
                            ICell sourceCell = sourceRow.GetCell(col);
                            if (sourceCell == null) continue;

                            ICell newCell = newRow.CreateCell(col);

                            switch (sourceCell.CellType)
                            {
                                case CellType.String:
                                    newCell.SetCellValue(sourceCell.StringCellValue);
                                    break;
                                case CellType.Numeric:
                                    newCell.SetCellValue(sourceCell.NumericCellValue);
                                    break;
                                case CellType.Boolean:
                                    newCell.SetCellValue(sourceCell.BooleanCellValue);
                                    break;
                                case CellType.Formula:
                                    newCell.SetCellValue(sourceCell.ToString()); // Fórmulas como texto
                                    break;
                                default:
                                    newCell.SetCellValue(sourceCell.ToString());
                                    break;
                            }
                        }
                    }
                }

                using (var outStream = new FileStream(xlsxPath, FileMode.Create, FileAccess.Write))
                {
                    xssfWorkbook.Write(outStream);
                }
            }
        }
    }

    public static class ExcelHelper
    {
        /// <summary>
        /// Si el archivo es .xls, lo convierte a .xlsx. Si ya es .xlsx, lo devuelve tal cual.
        /// Ideal para preparar el archivo antes de usarlo con EPPlus.
        /// </summary>
        public static string PrepararArchivoParaEpplus(string archivoEntrada)
        {
            if (!File.Exists(archivoEntrada))
                throw new FileNotFoundException("El archivo no existe.", archivoEntrada);

            string extension = Path.GetExtension(archivoEntrada).ToLowerInvariant();

            if (extension == ".xlsx")
            {
                return archivoEntrada;
            }
            else if (extension == ".xls")
            {
                string nuevoArchivo = Path.Combine(
                    Path.GetDirectoryName(archivoEntrada),
                    Path.GetFileNameWithoutExtension(archivoEntrada) + "_convertido.xlsx");

                ExcelConverter.ConvertXlsToXlsx(archivoEntrada, nuevoArchivo);
                return nuevoArchivo;
            }
            else
            {
                throw new NotSupportedException("Formato de archivo no soportado: " + extension);
            }
        }
    }


    

}
