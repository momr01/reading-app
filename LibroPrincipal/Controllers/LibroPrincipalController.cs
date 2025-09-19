using MathNet.Numerics.Distributions;
using MultasLectura.Helpers;
using MultasLectura.LibroCalidad.Controllers;
using MultasLectura.LibroCalidad.Interfaces;
using MultasLectura.LibroPlazos.Controllers;
using MultasLectura.LibroPlazos.Interfaces;
using MultasLectura.LibroPrincipal.Interfaces;
using MultasLectura.Models;
using NPOI.SS.Formula.Functions;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Diagnostics.Metrics;
using System.Globalization;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using static MultasLectura.Views.GenerarTodo;

namespace MultasLectura.LibroPrincipal.Controllers
{
   


    public class LibroPrincipalController : ILibroPrincipalController
    {
        private readonly BaremoModel _baremos;
        private readonly MetaModel _metas;
        private readonly IPlazosHojaResumenController _hojaPlazosController;
        private readonly ICalidadHojaResumenController _hojaCalidadController;
        private readonly ICalidadHojaCuadrosController _hojaCuadrosController;
        private readonly ICalidadHojaResLecturistaController _hojaResLecturistaController;
        private readonly DateTime _fechaDesde;
        private readonly DateTime _fechaHasta;



        public LibroPrincipalController(BaremoModel baremos, MetaModel metas, DateTime fechaDesde, DateTime fechaHasta)
        {
            _baremos = baremos;
            _hojaPlazosController = new PlazosHojaResumenController(_baremos);
            _hojaCalidadController = new CalidadHojaResumenController();
            _hojaCuadrosController = new CalidadHojaCuadrosController();
            _hojaResLecturistaController = new CalidadHojaResLecturistaController();
            // _baremos = baremos;
            _metas = metas;
            _fechaDesde = fechaDesde;
            _fechaHasta = fechaHasta;
        }

        public void GenerarLibros(string rutaCertificacion, string rutaResumen, string rutaGuardar, string rutaPlazos, string rutaCalOperario, string rutaReclamos, string rutaCalidad)
        {
            /* DateTime fecha = DateTime.Today; // o cualquier otra fecha
             string fechaFormateada = fecha.ToString("dd_MM_yyyy");

             string mes = "MARZO";*/

            int diaDesde = _fechaDesde.Date.Day;
            int mesDesde = _fechaDesde.Date.Month;
            int anioDesde = _fechaDesde.Date.Year;

            int diaHasta = _fechaHasta.Date.Day;
            int mesHasta = _fechaHasta.Date.Month;
            int anioHasta = _fechaHasta.Date.Year;

            // Obtiene el nombre completo del mes en español
            string nombreMes = _fechaHasta.ToString("MMMM", new CultureInfo("es-ES"));

            // Primera letra en mayúscula
            nombreMes = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(nombreMes);






            /*  List<Ftl> ftls = new List<Ftl>();
              ftls.Add(new Ftl(-3, 1549));
              ftls.Add(new Ftl(-2, 9915));
              ftls.Add(new Ftl(-1, 46724));
              ftls.Add(new Ftl(0, 105535));
              ftls.Add(new Ftl(1, 50493));
              ftls.Add(new Ftl(2, 12262));
              ftls.Add(new Ftl(3, 843));*/

            /* List<Tipo> tipos = new List<Tipo>();
             tipos.Add(new Tipo(1, "error de lectura", 100, 10000));
             tipos.Add(new Tipo(2, "lecturas faltantes", 200, 20000));
             tipos.Add(new Tipo(3, "novedades mal y no informadas", 300, 30000));
             tipos.Add(new Tipo(4, "reclamos", 400, 40000));
             tipos.Add(new Tipo(5, "lecturas fuera plazo", 500, 50000));
             tipos.Add(new Tipo(6, "bonificación lectura", 600, -60000));*/







            /*  string rutaGuardarCertificacion = Path.Combine(Path.GetDirectoryName(rutaGuardar)!, $"Certificacion {fechaFormateada} al {fechaFormateada}.xlsx");
              string rutaGuardarResumen = Path.Combine(Path.GetDirectoryName(rutaGuardar)!, $"Resumen_Multa_{mes}_{fecha.Year}.xlsx");*/
            string rutaGuardarCertificacion = Path.Combine(Path.GetDirectoryName(rutaGuardar)!, $"Certificacion {diaDesde}_{mesDesde}_{anioDesde} al {diaHasta}_{mesHasta}_{anioHasta}.xlsx");
            string rutaGuardarResumen = Path.Combine(Path.GetDirectoryName(rutaGuardar)!, $"Resumen_Multa_{nombreMes.ToUpper()}_{anioHasta}.xlsx");
            string rutaGuardarPlazos = Path.Combine(Path.GetDirectoryName(rutaGuardar)!, $"+ Plazos Lectura.xlsx");
            string rutaGuardarCalidad= Path.Combine(Path.GetDirectoryName(rutaGuardar)!, $"+ Calidad Lectura.xlsx");
            string rutaGuardarReclamos = Path.Combine(Path.GetDirectoryName(rutaGuardar)!, $"+ Error_Lect_Reclamos.xlsx");

            double importeCertificacion = LibroCertificacion(rutaCertificacion, rutaGuardarCertificacion, $"{diaDesde}_{mesDesde}_{anioDesde}", $"{diaHasta}_{mesHasta}_{anioHasta}");
           
           Dictionary<string, dynamic> valores = LibroPlazos(rutaPlazos, rutaGuardarPlazos);
            //  LibroResumenMulta(rutaResumen, rutaGuardarResumen, ftls, improcedencias, mes);
            //  List<Grupo> improcedencias = LibroCalidad(rutaCalOperario, rutaReclamos, rutaCalidad, rutaGuardarCalidad, importeCertificacion);
           var resultado = LibroCalidad(rutaCalOperario, rutaReclamos, rutaCalidad, rutaGuardarCalidad, importeCertificacion);

            /*
            List<Grupo> improcedencias = new List<Grupo>
{
    new Grupo
    {
        Etiqueta = "Error de lectura",
        Datos = new List<Dato>
        {
           new Dato { Etiqueta = "Cant. Generada", Valor = tipos.Where(tipo => tipo.Id == 1).First().Cantidad},
            new Dato { Etiqueta = "% Generacion", Valor = 0.040},
             new Dato { Etiqueta = "$ multa", Valor = tipos.Where(tipo => tipo.Id == 1).First().Importe},
        }
    },
     new Grupo
    {
        Etiqueta = "Lecturas Faltantes",
        Datos = new List<Dato>
         {
           new Dato { Etiqueta = "Cant. Generada", Valor = tipos.Where(tipo => tipo.Id == 2).First().Cantidad},
            new Dato { Etiqueta = "% Generacion", Valor = 0.040},
             new Dato { Etiqueta = "$ multa", Valor = tipos.Where(tipo => tipo.Id == 2).First().Importe},
        }
    },
      new Grupo
    {
        Etiqueta = "Novedades Mal y No Informadas",
        Datos = new List<Dato>
         {
           new Dato { Etiqueta = "Cant. Generada", Valor = tipos.Where(tipo => tipo.Id == 3).First().Cantidad},
            new Dato { Etiqueta = "% Generacion", Valor = 0.040},
             new Dato { Etiqueta = "$ multa", Valor = tipos.Where(tipo => tipo.Id == 3).First().Importe},
        }
    },
       new Grupo
    {
        Etiqueta = "Reclamos",
        Datos = new List<Dato>
        {
           new Dato { Etiqueta = "Cant. Generada", Valor = tipos.Where(tipo => tipo.Id == 4).First().Cantidad},
            new Dato { Etiqueta = "% Generacion", Valor = 0.040},
             new Dato { Etiqueta = "$ multa", Valor = tipos.Where(tipo => tipo.Id == 4).First().Importe},
        }
    },
       new Grupo
    {
        Etiqueta = "Lecturas Fuera Plazo",
        Datos = new List<Dato>
        {
           new Dato { Etiqueta = "Cantidad", Valor = tipos.Where(tipo => tipo.Id == 5).First().Cantidad},
            new Dato { Etiqueta = "% Fuera Plazo", Valor = 0.040},
             new Dato { Etiqueta = "$ multa", Valor = tipos.Where(tipo => tipo.Id == 5).First().Importe},
        }
    },
       new Grupo
    {
        Etiqueta = "Bonificación Lectura",
        Datos = new List<Dato>
        {
           new Dato { Etiqueta = "Cantidad", Valor = tipos.Where(tipo => tipo.Id == 6).First().Cantidad},
            new Dato { Etiqueta = "% Bonificado", Valor = 0.040},
             new Dato { Etiqueta = "$ multa", Valor = tipos.Where(tipo => tipo.Id == 6).First().Importe},
        }
    },

};*/
            //  LibroResumenMulta(rutaResumen, rutaGuardarResumen, cantidadesFtl, improcedencias, mes);
            resultado.improcedencias.Add(new Grupo
            {
                Etiqueta = "Lecturas Fuera Plazo",
                Datos = new List<Dato>
        {
           new Dato { Etiqueta = "Cantidad", Valor =0},
            new Dato { Etiqueta = "% Fuera Plazo", Valor = 0.040},
             new Dato { Etiqueta = "$ multa", Valor = valores["multa"]},
        }
            });
            resultado.improcedencias.Add(new Grupo
            {
                Etiqueta = "Bonificación Lectura",
                Datos = new List<Dato>
        {
           new Dato { Etiqueta = "Cantidad", Valor = 0},
            new Dato { Etiqueta = "% Bonificado", Valor = 0.040},
             new Dato { Etiqueta = "$ multa", Valor = valores["bonificacion"]},
        }
            });

            LibroResumenMulta(rutaResumen, rutaGuardarResumen, valores["ftl"], resultado.improcedencias, nombreMes.ToUpper());

            LibroReclamos(rutaReclamos, rutaGuardarReclamos);

          //  MessageBox.Show(rutaGuardar);

            // Comprimir los archivos en un ZIP
            //  string zipPath = Path.Combine(Path.GetDirectoryName(rutaGuardar)!, "Exportados.zip");
            string zipPath = Path.Combine(rutaGuardar, $"multas_{resultado.nombreContratista}_{nombreMes}_{anioHasta}.zip");
            if (File.Exists(zipPath)) File.Delete(zipPath);

            using (var zip = ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                zip.CreateEntryFromFile(rutaGuardarCertificacion, Path.GetFileName(rutaGuardarCertificacion));
                zip.CreateEntryFromFile(rutaGuardarResumen, Path.GetFileName(rutaGuardarResumen));
                zip.CreateEntryFromFile(rutaGuardarPlazos, Path.GetFileName(rutaGuardarPlazos));
                zip.CreateEntryFromFile(rutaGuardarCalidad, Path.GetFileName(rutaGuardarCalidad));
                zip.CreateEntryFromFile(rutaGuardarReclamos, Path.GetFileName(rutaGuardarReclamos));
                //  zip.CreateEntryFromFile(file3, Path.GetFileName(file3));
            }

            // Opcional: borrar los Excel sueltos
            File.Delete(rutaGuardarCertificacion);
            File.Delete(rutaGuardarResumen);
            File.Delete(rutaGuardarPlazos);
            File.Delete(rutaGuardarCalidad);
            File.Delete(rutaGuardarReclamos);
            // File.Delete(file3);

           // MessageBox.Show("Archivos exportados y comprimidos correctamente.");



            // Eliminar archivos temporales "_convertido.xlsx"
            string carpeta = Path.GetDirectoryName(rutaGuardar)!;
            var archivosConvertidos = Directory.GetFiles(carpeta, "*_convertido.xlsx");

            foreach (var archivo in archivosConvertidos)
            {
                try
                {
                    File.Delete(archivo);
                }
                catch (Exception ex)
                {
                    // Si algún archivo está en uso o no se puede borrar, lo ignoramos
                    Console.WriteLine($"No se pudo borrar {archivo}: {ex.Message}");
                }
            }


        }

        private double LibroCertificacion(string rutaCertificacion, string rutaGuardar, string fechaDesde, string fechaHasta) {

            using ExcelPackage libroCertificacion = new(new FileInfo(rutaCertificacion));
            ExcelWorksheet hojaBaseCertificacion = libroCertificacion.Workbook.Worksheets[0];
            hojaBaseCertificacion.Name = "detalle_lect";

            //creamos hojas nuevas del libro
            ExcelWorksheet hojaResumen = libroCertificacion.Workbook.Worksheets.Add("Resumen");

            //ubicacion de hojas
            libroCertificacion.Workbook.Worksheets.MoveBefore(hojaResumen.Name, hojaBaseCertificacion.Name);

            //Obtener rangos de las hojas que utilizaremos
            var rangoCertificacion = hojaBaseCertificacion.Cells[hojaBaseCertificacion.Dimension.Address];

            //Convertir a número la columna de la hoja plazos detalles
            LibroExcelHelper.ConvertirTextoANumero(rangoCertificacion);

            //Agregar contenido
            //AgregarContenidoHojaResumen(hojaResumen, hojaBasePlazosDet, rangoPlazosDetalles);


            double importeFinal = GenerarExcelCertificacion(hojaResumen, hojaBaseCertificacion, fechaDesde, fechaHasta);


            DefinirHojaActiva(libroCertificacion, hojaResumen);

            //guardar libro calidad
            libroCertificacion.SaveAs(new FileInfo(rutaGuardar));
            // GenerarExcelInicial(rutaGuardar);
            //CrearExcelResumenFinal(rutaGuardar);
            // CreateExcel5(rutaGuardar);

            return importeFinal;
        }

        private void LibroResumenMulta(string rutaResumen, string rutaGuardar, List<Ftl> ftls, List<Grupo> datos, string mes) {

            using ExcelPackage libroResumen = new(new FileInfo(rutaResumen));
            ExcelWorksheet hojaBaseResumen = libroResumen.Workbook.Worksheets[0];
            hojaBaseResumen.Name = "detalle";

            //creamos hojas nuevas del libro
            ExcelWorksheet hojaResumen = libroResumen.Workbook.Worksheets.Add("Resumen");

            //ubicacion de hojas
            libroResumen.Workbook.Worksheets.MoveBefore(hojaResumen.Name, hojaBaseResumen.Name);

            //Obtener rangos de las hojas que utilizaremos
            var rangoPlazos = hojaBaseResumen.Cells[hojaBaseResumen.Dimension.Address];

            //Convertir a número la columna de la hoja plazos detalles
            LibroExcelHelper.ConvertirTextoANumero(rangoPlazos);

            //Agregar contenido
            //AgregarContenidoHojaResumen(hojaResumen, hojaBasePlazosDet, rangoPlazosDetalles);


            //VIP
            GenerarExcelResumen(hojaResumen, hojaBaseResumen, ftls, datos, mes);


            // --- Hacer que "Resumen" sea la hoja visible por defecto ---
            /* hojaResumen.Select(); // marcarla como seleccionada
             libroResumen.Workbook.View.ActiveTab = hojaResumen.Index - 1; // ajustar pestaña activa*/
            // --- Hacer que "Resumen" sea la hoja visible por defecto ---
            /*   hojaResumen.Select();
               int pos = libroResumen.Workbook.Worksheets.IndexOf(hojaResumen);
               libroResumen.Workbook.View.ActiveTab = pos;*/
            // Ajustar hoja visible por defecto (Resumen)
            /* int posResumen = 0;
             for (int i = 0; i < libroResumen.Workbook.Worksheets.Count; i++)
             {
                 if (libroResumen.Workbook.Worksheets[i].Name == hojaResumen.Name)
                 {
                     posResumen = i;
                     break;
                 }
             }
             libroResumen.Workbook.View.ActiveTab = posResumen;*/

            DefinirHojaActiva(libroResumen, hojaResumen);


            //guardar libro calidad
            libroResumen.SaveAs(new FileInfo(rutaGuardar));
            // GenerarExcelInicial(rutaGuardar);
            //CrearExcelResumenFinal(rutaGuardar);
            // CreateExcel5(rutaGuardar);
        }

        //  private void LibroPlazos(string rutaPlazos, string rutaGuardar)
        private Dictionary<string, dynamic> LibroPlazos(string rutaPlazos, string rutaGuardar)
        {

            using ExcelPackage libroPlazos = new(new FileInfo(rutaPlazos));
            ExcelWorksheet hojaBasePlazos = libroPlazos.Workbook.Worksheets[0];
            hojaBasePlazos.Name = "Plazos_Lectura";

            //creamos hojas nuevas del libro
            ExcelWorksheet hojaResumen = libroPlazos.Workbook.Worksheets.Add("Resumen");

            //ubicacion de hojas
            libroPlazos.Workbook.Worksheets.MoveBefore(hojaResumen.Name, hojaBasePlazos.Name);

            //Obtener rangos de las hojas que utilizaremos
            var rangoPlazos = hojaBasePlazos.Cells[hojaBasePlazos.Dimension.Address];

            //Convertir a número la columna de la hoja plazos detalles
            LibroExcelHelper.ConvertirTextoANumero(rangoPlazos);

            //Agregar contenido
            //AgregarContenidoHojaResumen(hojaResumen, hojaBasePlazosDet, rangoPlazosDetalles);


            //VIP
            //GenerarExcelInicial(hojaResumen, hojaBaseCertificacion);

           Dictionary<string, dynamic> valores = AgregarContenidoHojaResumen(hojaResumen, hojaBasePlazos, rangoPlazos);

            DefinirHojaActiva(libroPlazos, hojaResumen);


            //guardar libro calidad
            libroPlazos.SaveAs(new FileInfo(rutaGuardar));
            // GenerarExcelInicial(rutaGuardar);
            //CrearExcelResumenFinal(rutaGuardar);
            // CreateExcel5(rutaGuardar);

            return valores;






          /*  using ExcelPackage libroPlazosDetalles = new(new FileInfo(rutaPlazosDetalles));
            ExcelWorksheet hojaBasePlazosDet = libroPlazosDetalles.Workbook.Worksheets[0];

            //creamos hojas nuevas del libro
            ExcelWorksheet hojaResumen = libroPlazosDetalles.Workbook.Worksheets.Add("Resumen");

            //ubicacion de hojas
            libroPlazosDetalles.Workbook.Worksheets.MoveBefore(hojaResumen.Name, hojaBasePlazosDet.Name);

            //Obtener rangos de las hojas que utilizaremos
            var rangoPlazosDetalles = hojaBasePlazosDet.Cells[hojaBasePlazosDet.Dimension.Address];

            //Convertir a número la columna de la hoja plazos detalles
            LibroExcelHelper.ConvertirTextoANumero(rangoPlazosDetalles);

            //Agregar contenido
            AgregarContenidoHojaResumen(hojaResumen, hojaBasePlazosDet, rangoPlazosDetalles);


            //guardar libro calidad
            libroPlazosDetalles.SaveAs(new FileInfo(rutaGuardar));*/
        }

        //private void AgregarContenidoHojaResumen(ExcelWorksheet hojaResumen, ExcelWorksheet hojaPlazosDetalles, ExcelRange rango)
        private Dictionary<string, dynamic> AgregarContenidoHojaResumen(ExcelWorksheet hojaResumen, ExcelWorksheet hojaPlazosDetalles, ExcelRange rango)
        {
            int numFila = 1;

            _hojaPlazosController.CrearTablaDinCertAtrasoTotal(hojaResumen, rango);
            //Dictionary<string, double> importesFinales = _hojaPlazosController.CrearTablaDatosPorTarifa(hojaResumen, hojaPlazosDetalles, ref numFila);
            // List<Ftl> cantidadesFtl = _hojaPlazosController.CrearTablaDatosPorTarifa(hojaResumen, hojaPlazosDetalles, ref numFila);
            Dictionary<string, dynamic> valores = _hojaPlazosController.CrearTablaDatosPorTarifa(hojaResumen, hojaPlazosDetalles, ref numFila);
            //_hojaPlazosController.CrearTablaImportesFinales(hojaResumen, importesFinales["multaT1"], importesFinales["multaT2"], ref numFila);
            _hojaPlazosController.CrearTablaImportesFinales(hojaResumen, ref numFila);


            return valores;

        }

        public static string GetFirstValueFromColumn(ExcelWorksheet worksheet, string columnLetter)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet), "La hoja no puede ser nula.");

            int colIndex = ColumnLetterToNumber(columnLetter);

            int startRow = worksheet.Dimension.Start.Row;
            int endRow = worksheet.Dimension.End.Row;

            for (int row = startRow; row <= endRow; row++)
            {
                if(row == startRow + 1)
                {
                    var value = worksheet.Cells[row, colIndex].Text;
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        if (value.ToLower().Contains("symesa"))
                        {
                            return "SYMESA";
                        }
                        else if (value.ToLower().Contains("pamar"))
                        {
                            return "PAMAR";
                        }
                        else
                        {
                            return "ERLYFSA";
                        }
                    }
                    //return value;
                }

            }

            return "contratista";
        }

        /// <summary>
        /// Convierte letra de columna en número (ej: "A"=1, "B"=2, "AA"=27).
        /// </summary>
        private static int ColumnLetterToNumber(string columnLetter)
        {
            columnLetter = columnLetter.ToUpperInvariant();
            int sum = 0;
            foreach (char c in columnLetter)
            {
                sum *= 26;
                sum += (c - 'A' + 1);
            }
            return sum;
        }


        //private List<Grupo> LibroCalidad(string rutaCalXOper, string rutaReclDetalles, string rutaCalDetalles, string rutaGuardar, double importeCertificacion)
             private (List<Grupo> improcedencias, string nombreContratista) LibroCalidad(string rutaCalXOper, string rutaReclDetalles, string rutaCalDetalles, string rutaGuardar, double importeCertificacion)
        {
            using ExcelPackage libroCalXOperario = new(new FileInfo(rutaCalXOper));
            ExcelWorksheet hojaBaseCalXOp = libroCalXOperario.Workbook.Worksheets[0];

            using ExcelPackage libroReclDetalles = new(new FileInfo(rutaReclDetalles));
            ExcelWorksheet hojaBaseReclDetalles = libroReclDetalles.Workbook.Worksheets[0];

            using ExcelPackage libroCalDetalles = new(new FileInfo(rutaCalDetalles));
            ExcelWorksheet hojaBaseCalDetalles = libroCalDetalles.Workbook.Worksheets[0];
            hojaBaseCalDetalles.Name = "Calidad";

            //creamos hojas nuevas del libro
            ExcelWorksheet hojaResumen = libroCalDetalles.Workbook.Worksheets.Add("Resumen");
            ExcelWorksheet hojaResLecturista = libroCalDetalles.Workbook.Worksheets.Add("Res-Lecturista");
            ExcelWorksheet hojaCantXOperario = libroCalDetalles.Workbook.Worksheets.Add("Cant_x_Oper", hojaBaseCalXOp);
            ExcelWorksheet hojaCuadros = libroCalDetalles.Workbook.Worksheets.Add("Cuadros");
            ExcelWorksheet hojaEliminados = libroCalDetalles.Workbook.Worksheets.Add("ELIMINADOS");

            //ubicacion de hojas
            /*  libroCalDetalles.Workbook.Worksheets.MoveBefore("Resumen", "calidad_detalle");
              libroCalDetalles.Workbook.Worksheets.MoveBefore("Res-Lecturista", "calidad_detalle");*/
            libroCalDetalles.Workbook.Worksheets.MoveBefore("Resumen", "Calidad");
            libroCalDetalles.Workbook.Worksheets.MoveBefore("Res-Lecturista", "Calidad");


            //Obtener rangos de las hojas que utilizaremos
            var rangoHojaCantXOperario = hojaCantXOperario.Cells[hojaCantXOperario.Dimension.Address];
            var rangoCalidadDetalles = hojaBaseCalDetalles.Cells[hojaBaseCalDetalles.Dimension.Address];
            var rangoCalXOperario = hojaCantXOperario.Cells[hojaCantXOperario.Dimension.Address];

            //Convertir a número la columna de la hoja cal x operario
            LibroExcelHelper.ConvertirTextoANumero(rangoHojaCantXOperario);

            Dictionary<string, int> reclamosValores = new();

            int numeroColumna = LibroExcelHelper.ObtenerNumeroColumna(hojaBaseReclDetalles, "desc_tar");
            if (numeroColumna != -1)
            {
                reclamosValores = ReclamosPorTarifa(hojaBaseReclDetalles, numeroColumna);
            }
            else
            {
                throw new Exception();
            }

            //Agregar contenido
           // AgregarContenidoHojaResumen2(hojaBaseCalDetalles, hojaResumen, rangoCalidadDetalles, _baremos, _metas, hojaBaseCalXOp, importeCertificacion, reclamosValores);
            List<Grupo> improcedencias = AgregarContenidoHojaResumen2(hojaBaseCalDetalles, hojaResumen, rangoCalidadDetalles, _baremos, _metas, hojaBaseCalXOp, importeCertificacion, reclamosValores);
            AgregarContenidoHojaCuadros(hojaCuadros, rangoCalidadDetalles, rangoCalXOperario);
            AgregarContenidoHojaResLecturista(hojaCantXOperario, hojaBaseCalDetalles, hojaResLecturista);

            string nombreContratista = GetFirstValueFromColumn(hojaBaseCalXOp, "B");


            /* int posResumen = 0;
             for (int i = 0; i < libroCalDetalles.Workbook.Worksheets.Count; i++)
             {
                 if (libroCalDetalles.Workbook.Worksheets[i].Name == hojaResumen.Name)
                 {
                     posResumen = i;
                     break;
                 }
             }
             libroCalDetalles.Workbook.View.ActiveTab = posResumen;*/
            DefinirHojaActiva(libroCalDetalles, hojaResumen);


            //guardar libro calidad
            libroCalDetalles.SaveAs(new FileInfo(rutaGuardar));


            return ( improcedencias, nombreContratista );

        }

        private void DefinirHojaActiva(ExcelPackage libro, ExcelWorksheet hoja)
        {
            int posResumen = 0;
            for (int i = 0; i < libro.Workbook.Worksheets.Count; i++)
            {
                if (libro.Workbook.Worksheets[i].Name == hoja.Name)
                {
                    posResumen = i;
                    break;
                }
            }
            libro.Workbook.View.ActiveTab = posResumen;
        }


        private Dictionary<string, int> ReclamosPorTarifa(ExcelWorksheet hoja, int numeroColumna)
        {
            int contFilas = hoja.Dimension.Rows;

            int totalReclT1 = 0;
            int totalReclT2 = 0;

            for (int row = 1; row <= contFilas; row++)
            {
                object cellValue = hoja.Cells[row, numeroColumna].Value;
                if (cellValue != null)
                {

                    if (cellValue.ToString()!.ToLower().Contains("t1"))
                    {
                        totalReclT1++;
                    }
                    else if (cellValue.ToString()!.ToLower().Contains("t2"))
                    {
                        totalReclT2++;
                    }

                }
            }

            return new()
            {
                ["t1"] = totalReclT1,
                ["t2"] = totalReclT2
            };
        }





        private List<Grupo> AgregarContenidoHojaResumen2(
            ExcelWorksheet hojaBase,
            ExcelWorksheet hojaResumen,
            ExcelRange rango,
            BaremoModel baremos,
            MetaModel metas,
            ExcelWorksheet hojaCalXOperario,
            double importeCertificacion,
            Dictionary<string, int> reclamosValores

            )
        {
            //  _hojaResumenController.CrearTablaDinTipoEstado(hojaResumen, rango);
            _hojaCalidadController.CrearTablaDinTipoEstado(hojaResumen, rango);
            // Dictionary<string, double> totales = _hojaResumenController.CrearTablaMetodoLineal(hojaResumen, hojaBase, baremos);
            Dictionary<string, double> totales = _hojaCalidadController.CrearTablaMetodoLineal(hojaResumen, hojaBase, baremos);
            //Dictionary<string, double> propInMasImpMetLineal = _hojaResumenController.CrearTablaTotales(hojaResumen, totales, reclamosValores, baremos, hojaCalXOperario, importeCertificacion);
            Dictionary<string, double> propInMasImpMetLineal = _hojaCalidadController.CrearTablaTotales(hojaResumen, totales, reclamosValores, baremos, hojaCalXOperario, importeCertificacion);
           /* _hojaResumenController.CrearTablaValorFinalMulta(hojaResumen, propInMasImpMetLineal["propInconformidades"], propInMasImpMetLineal["totalMetLineal"], importeCertificacion, metas, hojaBase, propInMasImpMetLineal["totalReclamos"]);
            _hojaResumenController.CrearTablaBaremosMetas(hojaResumen, baremos, metas, propInMasImpMetLineal["propInconformidades"]);*/
          List<Grupo> improcedencias =  _hojaCalidadController.CrearTablaValorFinalMulta(hojaResumen, propInMasImpMetLineal["propInconformidades"], propInMasImpMetLineal["totalMetLineal"], importeCertificacion, metas, hojaBase, propInMasImpMetLineal["totalReclamos"]);
            _hojaCalidadController.CrearTablaBaremosMetas(hojaResumen, baremos, metas, propInMasImpMetLineal["propInconformidades"]);
            hojaResumen.Cells.AutoFitColumns();
            hojaResumen.Column(2).AutoFit();

            return improcedencias;


        }

        private void AgregarContenidoHojaCuadros(ExcelWorksheet hojaCuadros, ExcelRange rangoCalidadDetalles, ExcelRange rangoCalXOperario)
        {
            _hojaCuadrosController.CrearTablaDinEmpleadoTotal(hojaCuadros, rangoCalXOperario);
            _hojaCuadrosController.CrearTablaDinLectorTotal(hojaCuadros, rangoCalidadDetalles);
            hojaCuadros.Cells.AutoFitColumns();

        }

        private void AgregarContenidoHojaResLecturista(ExcelWorksheet hojaCantXOper, ExcelWorksheet hojaCalidadDetalles, ExcelWorksheet hojaDestino)
        {
            _hojaResLecturistaController.CrearTablaLecturistaInconformidades(hojaCantXOper, hojaCalidadDetalles, hojaDestino);
            hojaDestino.Cells.AutoFitColumns();
        }

        private void AgregarContenidoHojaReclamos(ExcelWorksheet hoja, ExcelRange rango, BaremoModel baremos)
        {
            var pivotTable = hoja.PivotTables.Add(hoja.Cells["A9"], rango, "TablaDinamicaErrorLecturas");
            pivotTable.RowFields.Add(pivotTable.Fields["desc_tar"]);
           // var atraso = pivotTable.RowFields.Add(pivotTable.Fields["atraso"]);
            pivotTable.DataFields.Add(pivotTable.Fields["nic"]);
            pivotTable.DataFields[0].Function = DataFieldFunctions.Count;

            //atraso.Sort = eSortType.Ascending;


            string fechaStr = "";

            // Cortamos la parte antes de "(" porque molesta el parseo
            int idx = baremos.FechaEdicion.IndexOf("(");
            if (idx > 0)
                fechaStr = baremos.FechaEdicion.Substring(0, idx).Trim();

            DateTime fecha = DateTime.ParseExact(
                fechaStr,
                "ddd MMM dd yyyy HH:mm:ss 'GMT'K",
                CultureInfo.InvariantCulture
            );

            // Ahora podés extraer lo que quieras
            int dia = fecha.Day;
            int mes = fecha.Month;
            int anio = fecha.Year;




         //   hoja.Cells["F1"].Value = $"Baremo Lectura desde el {dia}/{mes}/{anio}";



            hoja.Cells["A1"].Value = $"Baremo Lectura desde el {dia}/{mes}/{anio}";
            //hoja.Cells["A1"].Style.Border.BorderAround(ExcelBorderStyle.Thick);
            hoja.Cells["A1"].Style.Font.Bold = true;

            hoja.Cells["A2"].Value = "T1 y T3";
            hoja.Cells["A2"].Style.Border.BorderAround(ExcelBorderStyle.Thick);
            hoja.Cells["A2"].Style.Font.Bold = true;

            hoja.Cells["B2"].Value = 732.48;
            hoja.Cells["B2"].Style.Font.Bold = true;
            hoja.Cells["B2"].Style.Numberformat.Format = "$ #,##0.00";
            hoja.Cells["B2"].Style.Border.BorderAround(ExcelBorderStyle.Thick);

            hoja.Cells["A3"].Value = "T2";
            hoja.Cells["A3"].Style.Border.BorderAround(ExcelBorderStyle.Thick);
            hoja.Cells["A3"].Style.Font.Bold = true;


            hoja.Cells["B3"].Value = 6761.35;
            hoja.Cells["B3"].Style.Font.Bold = true;
            hoja.Cells["B3"].Style.Numberformat.Format = "$ #,##0.00";
            hoja.Cells["B3"].Style.Border.BorderAround(ExcelBorderStyle.Thick);


            hoja.Cells["A4"].Value = "Altura T1 y T3";
            hoja.Cells["A4"].Style.Border.BorderAround(ExcelBorderStyle.Thick);
            hoja.Cells["A4"].Style.Font.Bold = true;


            hoja.Cells["B4"].Value = 8582.59;
            hoja.Cells["B4"].Style.Font.Bold = true;
            hoja.Cells["B4"].Style.Numberformat.Format = "$ #,##0.00";
            hoja.Cells["B4"].Style.Border.BorderAround(ExcelBorderStyle.Thick);

            hoja.Cells.AutoFitColumns();


        }

        private void LibroReclamos(string rutaReclamos, string rutaGuardar) {

            using ExcelPackage libroReclamos = new(new FileInfo(rutaReclamos));
            ExcelWorksheet hojaBaseReclamos = libroReclamos.Workbook.Worksheets[0];
            hojaBaseReclamos.Name = "Reclamos";

            //creamos hojas nuevas del libro
            ExcelWorksheet hojaResumen = libroReclamos.Workbook.Worksheets.Add("Resumen");

            //ubicacion de hojas
            libroReclamos.Workbook.Worksheets.MoveBefore(hojaResumen.Name, hojaBaseReclamos.Name);

            //Obtener rangos de las hojas que utilizaremos
            var rangoPlazos = hojaBaseReclamos.Cells[hojaBaseReclamos.Dimension.Address];

            //Convertir a número la columna de la hoja plazos detalles
            LibroExcelHelper.ConvertirTextoANumero(rangoPlazos);

            //Agregar contenido
            //AgregarContenidoHojaResumen(hojaResumen, hojaBasePlazosDet, rangoPlazosDetalles);


            //VIP
            //GenerarExcelInicial(hojaResumen, hojaBaseCertificacion);

            AgregarContenidoHojaReclamos(hojaResumen, rangoPlazos, _baremos);

            //creamos hoja de eliminados
           libroReclamos.Workbook.Worksheets.Add("Eliminados");

            DefinirHojaActiva(libroReclamos, hojaResumen);


            //guardar libro calidad
            libroReclamos.SaveAs(new FileInfo(rutaGuardar));
            // GenerarExcelInicial(rutaGuardar);
            //CrearExcelResumenFinal(rutaGuardar);
            // CreateExcel5(rutaGuardar);






            /*  using ExcelPackage libroPlazosDetalles = new(new FileInfo(rutaPlazosDetalles));
              ExcelWorksheet hojaBasePlazosDet = libroPlazosDetalles.Workbook.Worksheets[0];

              //creamos hojas nuevas del libro
              ExcelWorksheet hojaResumen = libroPlazosDetalles.Workbook.Worksheets.Add("Resumen");

              //ubicacion de hojas
              libroPlazosDetalles.Workbook.Worksheets.MoveBefore(hojaResumen.Name, hojaBasePlazosDet.Name);

              //Obtener rangos de las hojas que utilizaremos
              var rangoPlazosDetalles = hojaBasePlazosDet.Cells[hojaBasePlazosDet.Dimension.Address];

              //Convertir a número la columna de la hoja plazos detalles
              LibroExcelHelper.ConvertirTextoANumero(rangoPlazosDetalles);

              //Agregar contenido
              AgregarContenidoHojaResumen(hojaResumen, hojaBasePlazosDet, rangoPlazosDetalles);


              //guardar libro calidad
              libroPlazosDetalles.SaveAs(new FileInfo(rutaGuardar));*/

        }

        private void LibroResumenFinal()
        {

        }

        public class Dato
        {
           /* public double Cantidad { get; set; }
            public double Porcentaje { get; set; }
            public double MontoMulta { get; set; }*/
           public string Etiqueta { get; set; }
            public double Valor { get; set; }

           /* public Dato(string etiqueta, double Valor) {
            }*/
        }

        public class Grupo
        {
            public string Etiqueta { get; set; }
            public List<Dato> Datos { get; set; }
        }

        public class Ftl
        {
            public int Dia { get; set; }
            public int Cantidad { get; set; }

            public Ftl(int dia, int cantidad) {
            this.Dia = dia;
                this.Cantidad = cantidad;
            }
        }

        public class Tipo
        {
            public int Id { get; set; }
            public string Descripcion { get; set; }
            public int Cantidad { get; set; }
            public double Importe { get; set; }

            public Tipo(int id, string descripcion, int cantidad, double importe)
            {
                this.Id = id;
                this.Descripcion = descripcion;
                this.Importe = importe;
                this.Cantidad = cantidad;
            }
        }


        public static void GenerarExcelResumen(ExcelWorksheet ws, ExcelWorksheet libroDatos, List<Ftl> ftls, List<Grupo> datos, string mes)
        {
          




            // ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //  using var package = new ExcelPackage();
            //  var ws = package.Workbook.Worksheets.Add("Detalle T1");

            // Cabecera
            ws.Cells["A1"].Value = "Detalle T1";
            ws.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["A1"].Style.Font.Bold = true;
            ws.Cells["A1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(204, 255, 229));
            ws.Cells["A1:B1"].Merge = true;
            ws.Cells["A1"].Style.Border.BorderAround(ExcelBorderStyle.Thick);


            ws.Cells["C1"].Value = "NIC";
            ws.Cells["C1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["C1"].Style.Fill.BackgroundColor.SetColor(Color.Red);
            ws.Cells["C1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["C1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["C1"].Style.Border.BorderAround(ExcelBorderStyle.Thick);
            ws.Cells["C1"].Style.Font.Color.SetColor(Color.White);
            ws.Cells["C1"].Style.Font.Bold = true;


            ws.Cells["E1"].Value = "FTL";
            ws.Cells["E1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["E1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["E1"].Style.Font.Bold = true;
            ws.Cells["E1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["E1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(204, 255, 229));
           


            ws.Cells["F1"].Value = mes;
            ws.Cells["F1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["F1"].Style.Fill.BackgroundColor.SetColor(Color.Red);
            ws.Cells["F1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["F1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["F1"].Style.Font.Color.SetColor(Color.White);
            ws.Cells["F1"].Style.Font.Bold = true;


            List<string> ef = new List<string>()
            {
                "E", "F"
            };

            foreach (string t in ef)
            {
                using (var range = ws.Cells[$"{t}2:{t}31"])
                {
                    range.Style.Border.BorderAround(ExcelBorderStyle.Thick);
                    // range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                }

                ws.Cells[$"{t}32"].Style.Border.BorderAround(ExcelBorderStyle.Thick);
                ws.Cells[$"{t}33"].Style.Border.BorderAround(ExcelBorderStyle.Thick);
                ws.Cells[$"{t}1"].Style.Border.BorderAround(ExcelBorderStyle.Thick);

            }
           



            ws.Cells["B2"].Value = "Total";
            ws.Cells["B2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["B2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["B2"].Style.Border.BorderAround(ExcelBorderStyle.Thick);




            ws.Cells["C2"].Formula = "SUM(F2:F31)";
            ws.Cells["C2"].Style.Numberformat.Format = "#,##0";
            ws.Cells["C2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["C2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["C2"].Style.Border.BorderAround(ExcelBorderStyle.Thick);



            int fila = 3;
            int primeraFila = fila;
            foreach (var item in datos)
            {
              
                ws.Cells[fila, 1].Value = item.Etiqueta;


               foreach(var item2 in item.Datos)
                {
                    ws.Cells[fila, 2].Value = item2.Etiqueta;
                    ws.Cells[fila, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[fila, 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    

                    if (item2.Etiqueta.Contains("%"))
                    {
                        ws.Cells[fila, 3].Formula = $"C{fila - 1}/C2";
                        ws.Cells[fila, 3].Style.Numberformat.Format = "0.00%";
                    } else if(item2.Etiqueta.Contains('$')) 
                    {
                        if(item.Etiqueta.ToLower().Contains("bonificación lectura"))
                        {
                            ws.Cells[fila, 3].Value = (item2.Valor * -1);
                        } else
                        {
                            ws.Cells[fila, 3].Value = item2.Valor;
                        }
                       
                        ws.Cells[fila, 3].Style.Numberformat.Format = "$ #,##0.00";


                    } else if(item.Etiqueta.ToLower().Contains( "lecturas fuera plazo"))
                    {
                        ws.Cells[fila, 3].Formula = "SUM(F2:F9)+SUM(F15:F31)";
                    } else if(item.Etiqueta.ToLower().Contains("bonificación lectura"))
                    {
                        ws.Cells[fila, 3].Formula = "SUM(F10:F14)";
                    }
                    else
                    {

                        ws.Cells[fila, 3].Value = item2.Valor;
                    }

                    ws.Cells[fila, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[fila, 3].Style.VerticalAlignment = ExcelVerticalAlignment.Center;


                    fila++;


                }
                

               List<string> abc = new List<string>() { "A", "B", "C"};

                foreach (string letra  in abc)
                {
                    using (var range = ws.Cells[$"{letra}{primeraFila}:{letra}{fila - 1}"])
                    {
                        range.Style.Border.BorderAround(ExcelBorderStyle.Thick);
                       
                    }
                }

               
            }

            ws.Cells["A3:A5"].Merge = true;
            ws.Cells["A3:A5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A3:A5"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["A3:A5"].Style.WrapText = true;

            ws.Cells["A6:A8"].Merge = true;
            ws.Cells["A6:A8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A6:A8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["A6:A8"].Style.WrapText = true;

            ws.Cells["A9:A11"].Merge = true;
            ws.Cells["A9:A11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A9:A11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["A9:A11"].Style.WrapText = true;

            ws.Cells["A12:A14"].Merge = true;
            ws.Cells["A12:A14"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A12:A14"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["A12:A14"].Style.WrapText = true;

            ws.Cells["A15:A17"].Merge = true;
            ws.Cells["A15:A17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A15:A17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["A15:A17"].Style.WrapText = true;

            ws.Cells["A18:A20"].Merge = true;
            ws.Cells["A18:A20"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A18:A20"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["A18:A20"].Style.WrapText = true;






            int filaSubtotales = 21;
            List<string> subtotalesEtiquetas = new List<string>()
            {
                "multa", "bonificación", "descuento", "resultado"
            };

            foreach (string item in subtotalesEtiquetas)
            {
                ws.Cells[$"A{filaSubtotales}"].Value = item.ToUpper();
                ws.Cells[$"A{filaSubtotales}:B{filaSubtotales}"].Merge = true;
                ws.Cells[$"A{filaSubtotales}:B{filaSubtotales}"].Style.Border.BorderAround(ExcelBorderStyle.Thick);
                ws.Cells[$"A{filaSubtotales}:B{filaSubtotales}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells[$"A{filaSubtotales}:B{filaSubtotales}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                ws.Cells[$"A{filaSubtotales}:B{filaSubtotales}"].Style.Font.Bold = true;
                filaSubtotales++;
            }







            

            void AplicarFormatoSubtotalValor(string celda, string formula, Color colorFondo, Color colorLetra)
            {
                ws.Cells[$"{celda}"].Formula = formula;
                ws.Cells[$"{celda}"].Style.Font.Bold = true;
                ws.Cells[$"{celda}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells[$"{celda}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[$"{celda}"].Style.Numberformat.Format = "$ #,##0.00";
                ws.Cells[$"{celda}"].Style.Border.BorderAround(ExcelBorderStyle.Thick);
                ws.Cells[$"{celda}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[$"{celda}"].Style.Fill.BackgroundColor.SetColor(colorFondo);
                ws.Cells[$"{celda}"].Style.Font.Color.SetColor(colorLetra);

            }


           
            AplicarFormatoSubtotalValor("C21", "C5+C8+C11+C14+C17", Color.LightSalmon, Color.Black);
            AplicarFormatoSubtotalValor("C22", "C20", Color.LightGreen, Color.Black);
            AplicarFormatoSubtotalValor("C23", "0", Color.Red, Color.White);
            AplicarFormatoSubtotalValor("C24", "SUM(C21:C23)", Color.LightYellow, Color.Black);
           

            // FTL (Días) y datos
            int ftlInicio = 2;
            for (int i = -11; i <= 18; i++)
            {
                Ftl valor = ftls.FirstOrDefault(x => x.Dia == i);
                if (valor != null)
                {
                    ws.Cells[ftlInicio, 6].Value = valor.Cantidad;
                    ws.Cells[ftlInicio, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    ws.Cells[ftlInicio, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                }

                ws.Cells[ftlInicio, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[ftlInicio, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells[ftlInicio++, 5].Value = i;
               


            }


            // Gráfico
            var chart = ws.Drawings.AddChart("FTLChart", eChartType.Line) as ExcelLineChart;
            chart.Title.Text = "Mes Par";
            chart.SetPosition(1, 0, 7, 0);
            chart.SetSize(700, 400);
           /* var xRange = ws.Cells["E3:E9"];
            var yRange = ws.Cells["F3:F9"];*/
            var xRange = ws.Cells["E2:E31"];
            var yRange = ws.Cells["F2:F31"];
            var series = chart.Series.Add(yRange, xRange);
            chart.YAxis.Title.Text = "Cant. clientes";
            chart.XAxis.Title.Text = "Desfasaje del día de lectura (FTL en Días Hábiles)";
            chart.XAxis.Title.Font.Color = Color.Red;
            chart.XAxis.Title.Font.Size = 9;
            series.Header = "";


           // int filaCodigosHojas = 31;

            void CodigosHojas(int numCelda, string valor)
            {
                ws.Cells[$"A{numCelda}"].Value = valor;
                ws.Cells[$"A{numCelda}:B{numCelda}"].Merge = true;
                ws.Cells[$"A{numCelda}:B{numCelda}"].Style.Border.BorderAround(ExcelBorderStyle.Thick);
                ws.Cells[$"A{numCelda}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[$"A{numCelda}"].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                ws.Cells[$"A{numCelda}"].Style.Font.Bold = true;
                ws.Cells[$"A{numCelda}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[$"A{numCelda}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            }



            // Notas
            CodigosHojas(31, "Hoja de Pedido");
            CodigosHojas(32, "Hoja de Entrada");
           
            ws.Cells["C31:C32"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["C31:C32"].Style.Fill.BackgroundColor.SetColor(Color.Red);
            ws.Cells["C31"].Style.Border.BorderAround(ExcelBorderStyle.Thick);
            ws.Cells["C32"].Style.Border.BorderAround(ExcelBorderStyle.Thick);
           

            ws.Cells["E32"].Value = "Dentro Plazo";
            ws.Cells["E32"].Style.Font.Bold = true;
            ws.Cells["E32"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["E32"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(204, 255, 229));

          
            ws.Cells["F32"].Formula = "SUM(F10:F14)/SUM(F2:F31)";
            ws.Cells["F32"].Style.Numberformat.Format = "0.00%";
            ws.Cells["F32"].Style.Font.Bold = true;

            ws.Cells["E33"].Value = "Fuera de plazo";
            ws.Cells["E33"].Style.Font.Bold = true;
            ws.Cells["E33"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["E33"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(204, 255, 229));
            //  ws.Cells["F33"].Value = "6%";
            ws.Cells["F33"].Formula = "(SUM(F2:F9)+SUM(F15:F31))/SUM(F2:F31)";
            ws.Cells["F33"].Style.Numberformat.Format = "0.00%";
            ws.Cells["F33"].Style.Font.Bold = true;

            // Estilos generales
           

            ws.Cells.AutoFitColumns();

           // package.SaveAs(new FileInfo(rutaArchivo));
        }



        private static List<int> GenerarValor(ExcelWorksheet hojaDatos, string codUnicom, string tarifa, string eliminarPalabra)
        {
            int contFilas = hojaDatos.Dimension.Rows;

            //  int totalReclT1 = 0;
            //  int totalReclT2 = 0;
            int cantLect = 0;
            int conAnom = 0;
            int sinAnom = 0;

            for (int row = 1; row <= contFilas; row++)
            {
                object cellValue = hojaDatos.Cells[row, 3].Value;

               
                if (cellValue != null)
                {

                    if (cellValue.ToString()!.ToLower().Contains(tarifa.ToLower()) && !cellValue.ToString()!.ToLower().Contains(eliminarPalabra.ToLower()))
                    {
                        if (hojaDatos.Cells[row, 2].Value.ToString().ToLower().Contains(codUnicom.ToLower()))
                        {
                            cantLect += int.Parse(hojaDatos.Cells[row, 4].Value.ToString());
                            conAnom += int.Parse(hojaDatos.Cells[row, 5].Value.ToString());
                            sinAnom += int.Parse(hojaDatos.Cells[row, 6].Value.ToString());
                        } else if(tarifa.ToLower().Contains("t3") && codUnicom.ToUpper().Contains("ZONA NORTE"))
                        {
                            if (hojaDatos.Cells[row, 2].Value.ToString().ToLower().Contains("lecturas t3"))
                            {
                                cantLect += int.Parse(hojaDatos.Cells[row, 4].Value.ToString());
                                conAnom += int.Parse(hojaDatos.Cells[row, 5].Value.ToString());
                                sinAnom += int.Parse(hojaDatos.Cells[row, 6].Value.ToString());
                            }

                        }
                    }


                }
            }

           

            return new List<int>() { cantLect, conAnom, sinAnom };

        }

        class OficinaComercial {

            public int codUnicom;
            public string oficina;

            public OficinaComercial(int codUnicom, string oficina)
            {
                this.codUnicom = codUnicom;
                this.oficina = oficina;
            }
        }

        public static List<object[]> DatosXOficina(ExcelWorksheet libroDatos, string tarifa, string eliminarPalabra)
        {

            List<OficinaComercial> oficinas = new List<OficinaComercial>();

            switch (tarifa.ToUpper())
            {
                case "T1":
                    oficinas = new List<OficinaComercial>() {
                        new OficinaComercial(1102, "CAPITAL"),
               new OficinaComercial(1202, "GUAYMALLEN"),
                new OficinaComercial(1302, "LAS HERAS"),
                 new OficinaComercial(1402, "LUJAN"),
                  new OficinaComercial(1502, "MAIPU"),
                   new OficinaComercial(1602, "LAVALLE"),
                    new OficinaComercial(2102, "SAN CARLOS"),
                     new OficinaComercial(2202, "TUNUYAN"),
                      new OficinaComercial(2302, "TUPUNGATO"),
                       new OficinaComercial(3102, "SAN RAFAEL"),
                        new OficinaComercial(3202, "MALARGUE"),

            };
                    break;
                case "T2":
                    oficinas = new List<OficinaComercial>() {
                        new OficinaComercial(1902, "ZONA NORTE"),
               new OficinaComercial(3902, "ZONA SUR"),
            };
                    break;
                case "T3":
                    oficinas = new List<OficinaComercial>() { 
                        
                        new OficinaComercial(1702, "ZONA NORTE"),
                         new OficinaComercial(2102, "SAN CARLOS"),
                     new OficinaComercial(2202, "TUNUYAN"),
                      new OficinaComercial(2302, "TUPUNGATO"),
                       new OficinaComercial(3102, "SAN RAFAEL"),
                        new OficinaComercial(3202, "MALARGUE"),


                    };
                    break;
                case "ALTURA T1":
                    oficinas = new List<OficinaComercial>() {

                         new OficinaComercial(1102, "CAPITAL"),
               new OficinaComercial(1202, "GUAYMALLEN"),
                new OficinaComercial(1302, "LAS HERAS"),
                 new OficinaComercial(1402, "LUJAN"),
                  new OficinaComercial(1502, "MAIPU"),
                   new OficinaComercial(1602, "LAVALLE"),
                      new OficinaComercial(2302, "TUPUNGATO"),
                       new OficinaComercial(3102, "SAN RAFAEL"),
                      

                    };
                    break;
                case "ALTURA T3":
                    oficinas = new List<OficinaComercial>() {

                         new OficinaComercial(1702, "ZONA NORTE"),
                    };
                    break;


            }
            
            

            List<object[]> objects = new List<object[]>();

            foreach (var item in oficinas)
            {
                objects.Add(new object[] { item.codUnicom, item.oficina.ToUpper(),
                    GenerarValor(libroDatos, item.oficina.ToUpper(), tarifa.ToUpper(), eliminarPalabra)[0],
                    GenerarValor(libroDatos, item.oficina.ToUpper(), tarifa.ToUpper(), eliminarPalabra )[1],
                    GenerarValor(libroDatos, item.oficina.ToUpper(), tarifa.ToUpper(), eliminarPalabra)[2] });
            }


            return objects;

        }

        public double GenerarExcelCertificacion(ExcelWorksheet ws, ExcelWorksheet libroDatos, string fechaDesde, string fechaHasta)
        {
           

            // Encabezado
            ws.Cells[1, 5].Value = $"CERTIFICACIÓN PERIODO {fechaDesde} al {fechaHasta}";
            ws.Cells[1, 5, 1, 10].Merge = true;
            ws.Cells[1, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells[1, 5].Style.Font.Bold = true;
            ws.Cells[1, 5].Style.Font.UnderLine = true;

            string[] encabezados = { "CÓDIGO SAP", "cod_unicom", "OFICINA", "Realizadas", "Sin Lectura - Con código de Anomalía", "Sin Lectura - Sin código de Anomalía", "CECO", "LECTURAS A CERTIFICAR", "IMPORTE A CERTIFICAR" };

            for (int i = 0; i < encabezados.Length; i++)
            {
                ws.Cells[3, i + 2].Value = encabezados[i];
                ws.Cells[3, i + 2].Style.Font.Bold = true;
                ws.Cells[3, i + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[3, i + 2].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                //  ws.Cells[3, i + 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                LibroExcelHelper.AplicarBordeGruesoARango(ws.Cells[3, i + 2]);
            }

            int fila = 4;

           // int capitalT1 = GenerarValor(libroDatos, "CAPITAL", "T1")[0];
            var secciones = new List<(string Titulo, string Codigo, Color Color, List<object[]> Filas, string ceco, string baremos)>
        {
            ("T1", "C000001", Color.LightBlue, 
           DatosXOficina(libroDatos, "t1", "altura")
            , "3220101",
            "732.48"
           
            
            ),

            ("T2", "C000013", Color.LightSkyBlue, 
            
             DatosXOficina(libroDatos, "t2", "altura")

            , "3220102",
            "6761.35"
            
            
            ),

            ("T3", "C000001", Color.Bisque, 
            
             DatosXOficina(libroDatos, "t3", "altura")
            , "3220101",
            "732.48"
           
            ),

            ("Altura", "C001000", Color.LightGreen, 
          
            DatosXOficina(libroDatos, "altura t1", "t2")

            , "3220101",
            "8582.59"
            
            ),

            ("T3 Altura", "C001000", Color.MediumPurple, 
           
             DatosXOficina(libroDatos, "altura t3", "t2")

            , "0",
            "8582.59"
           
            
            )
        };

          

            foreach (var (titulo, codigo, color, filasSeccion, ceco, baremos) in secciones)
            {
                int filaPrimerValor = fila;
                ws.Cells[fila, 1].Value = titulo;
                ws.Cells[fila, 1].Style.Font.Bold = true;
                ws.Cells[fila, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[fila, 1].Style.Fill.BackgroundColor.SetColor(color);
             
                ws.Cells[fila, 2].Value = codigo;
                ws.Cells[fila, 2].Style.Font.Bold = true;
                ws.Cells[fila, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[fila, 2].Style.Fill.BackgroundColor.SetColor(color);
                fila++;

                foreach (var datos in filasSeccion)
                {
                    for (int col = 0; col < datos.Length; col++)
                    {
                        var valor = datos[col];
                        var celda = ws.Cells[fila - 1, col + 3];
                        celda.Value = valor;

                       

                        if (col == 1)
                        {
                           
                            celda.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        }

                      

                        if (col == 8 && valor is double d)
                        {
                            celda.Style.Numberformat.Format = "$ #,##0.00";
                        }
                    }
                    fila++;
                }

                ws.Cells[fila - 1, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[fila - 1, 3].Style.Fill.BackgroundColor.SetColor(color);

                ws.Cells[fila -1, 4].Value = "TOTAL GENERAL";
                ws.Cells[fila - 1, 4].Style.Font.Bold = true;
                ws.Cells[fila - 1, 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                // ws.Cells[fila - 1, 4].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                ws.Cells[fila - 1, 4].Style.Fill.BackgroundColor.SetColor(color);

                ws.Cells[fila - 1, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[fila - 1, 5].Style.Fill.BackgroundColor.SetColor(color);

                //  ws.Cells[fila - 1, 6].Value = totalLecturas;
                ws.Cells[fila - 1, 5].Formula = $"SUM(E{filaPrimerValor}:F{fila - 2})";
                ws.Cells[fila - 1, 5].Style.Font.Bold = true;
                ws.Cells[fila - 1, 5].Style.Fill.PatternType = ExcelFillStyle.Solid;
               // ws.Cells[fila - 1, 6].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                ws.Cells[fila - 1, 5].Style.Fill.BackgroundColor.SetColor(color);


                ws.Cells[$"E{fila - 1}:F{fila - 1}"].Merge = true;



                ws.Cells[fila - 1, 7].Formula = $"SUM(G{filaPrimerValor}:G{fila - 2})";
                ws.Cells[fila - 1, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[fila - 1, 7].Style.Fill.BackgroundColor.SetColor(Color.Red);

                ws.Cells[fila - 1, 8].Formula = $"SUM(E{fila - 1}:G{fila - 1})";
                ws.Cells[fila - 1, 8].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                /* ws.Cells[fila - 1, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                 ws.Cells[fila - 1, 7].Style.Fill.BackgroundColor.SetColor(Color.Red);*/
                LibroExcelHelper.AplicarBordeGruesoARango(ws.Cells[$"C{fila - 1}:G{fila - 1}"]);


                ws.Cells[filaPrimerValor, 8].Value = ceco;
                ws.Cells[filaPrimerValor, 9].Formula = $"SUM(E{filaPrimerValor}:F{fila - 2})";
                ws.Cells[filaPrimerValor, 9].Style.Font.Bold = true;
                ws.Cells[filaPrimerValor, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;



                // var baremosTexto = baremos.ToString(CultureInfo.InvariantCulture).Replace('.', ',');

                ws.Cells[fila - 1, 10].Formula = $"I{filaPrimerValor}*{baremos}";
              //  ws.Cells[fila - 1, 10].Formula = $"I{filaPrimerValor}*{baremosTexto}";
                ws.Cells[fila - 1, 10].Style.Numberformat.Format = "$ #,##0.00";
                ws.Cells[fila - 1, 10].Style.Font.Bold = true;
                //fila += 1;

                LibroExcelHelper.AplicarBordeFinoARango(ws.Cells[$"H{filaPrimerValor}:H{fila - 1}"]);

              
            }

            // Total final
            ws.Cells[fila, 9].Value = "Total importe";
            ws.Cells[fila, 9].Style.Font.Bold = true;
            // ws.Cells[fila, 10].Value = totalImporte;
            ws.Cells[fila, 10].Formula = $"SUM(J4:J{fila - 1})"; ;
            ws.Cells[fila, 10].Style.Numberformat.Format = "$ #,##0.00";
            ws.Cells[fila, 10].Style.Font.Bold = true;
            LibroExcelHelper.AplicarBordeGruesoARango(ws.Cells[$"I{fila}:J{fila}"]);

            ws.Cells[1, 1, fila, 10].AutoFitColumns();


            List<string> letrasParaMerge1 = new List<string>();
            letrasParaMerge1.Add("A");
            letrasParaMerge1.Add("B");

            //Merge columna A, tarifas y columna B, códigos

            letrasParaMerge1.ForEach(x =>
            {
                LibroExcelHelper.FormatoMergeCelda(ws, $"{x}4:{x}15");
                LibroExcelHelper.CentrarContenidoCelda(ws, $"{x}4:{x}15");
                LibroExcelHelper.AplicarBordeGruesoARango(ws.Cells[$"{x}4:{x}15"]);

                LibroExcelHelper.FormatoMergeCelda(ws, $"{x}16:{x}18");
                LibroExcelHelper.CentrarContenidoCelda(ws, $"{x}16:{x}18");
                LibroExcelHelper.AplicarBordeGruesoARango(ws.Cells[$"{x}16:{x}18"]);

                LibroExcelHelper.FormatoMergeCelda(ws, $"{x}19:{x}25");
                LibroExcelHelper.CentrarContenidoCelda(ws, $"{x}19:{x}25");
                LibroExcelHelper.AplicarBordeGruesoARango(ws.Cells[$"{x}19:{x}25"]);

                LibroExcelHelper.FormatoMergeCelda(ws, $"{x}26:{x}34");
                LibroExcelHelper.CentrarContenidoCelda(ws, $"{x}26:{x}34");
                LibroExcelHelper.AplicarBordeGruesoARango(ws.Cells[$"{x}26:{x}34"]);

                LibroExcelHelper.FormatoMergeCelda(ws, $"{x}35:{x}36");
                LibroExcelHelper.CentrarContenidoCelda(ws, $"{x}35:{x}36");
                LibroExcelHelper.AplicarBordeGruesoARango(ws.Cells[$"{x}35:{x}36"]);
            });


            LibroExcelHelper.FondoSolido(ws.Cells["H4:H36"], Color.GreenYellow);

            ws.Cells[39, 6].Value = "HOJA DE PEDIDO";
            ws.Cells[39, 6].Style.Font.Bold = true;
            ws.Cells[39, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[39, 6].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);

            ws.Cells[39, 7].Value = "";
            ws.Cells[39, 7].Style.Font.Bold = true;
            ws.Cells[39, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[39, 7].Style.Fill.BackgroundColor.SetColor(Color.Red);
            //  LibroExcelHelper.ColorFondoLetra(ws, 'G', 39, new ColorModel();

            ws.Cells[40, 6].Value = "HOJA DE ENTRADA";
            ws.Cells[40, 6].Style.Font.Bold = true;
            ws.Cells[40, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[40, 6].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);

            ws.Cells[40, 7].Value = "";
            ws.Cells[40, 7].Style.Font.Bold = true;
            ws.Cells[40, 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[40, 7].Style.Fill.BackgroundColor.SetColor(Color.Red);

            LibroExcelHelper.AplicarBordeGruesoARango(ws.Cells[$"F39:G40"]);

           

            for(int i = 4; i < 37; i++)
            {
                ws.Cells[i, 9].Style.Border.Right.Style = ExcelBorderStyle.Thick;
                ws.Cells[i, 9].Style.Border.Left.Style = ExcelBorderStyle.Thick;
                ws.Cells[i, 10].Style.Border.Right.Style = ExcelBorderStyle.Thick;

                if(i == 15 || i == 18 || i == 25 || i == 34 || i == 36)
                {
                    ws.Cells[i, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                    ws.Cells[i, 10].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                }
            }



            // MessageBox.Show(Convert.ToDouble(ws.Cells[37, 10].Value ?? 0).ToString());

            ws.Calculate();
            var val = ws.Cells[37, 10].Value;
            double result = Convert.ToDouble(val);


            // return Convert.ToDouble(ws.Cells[37, 10].Value ?? 0);
            return result;
        }



        public void CreateExcel5(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            FileInfo file = new FileInfo(filePath);

            using (var package = new ExcelPackage(file))
            {
                // Crear hoja detalle_lect
                var detalleSheet = package.Workbook.Worksheets.Add("detalle_lect");
                CreateDetalleLectSheet(detalleSheet);

                // Crear hoja principal
                var mainSheet = package.Workbook.Worksheets.Add("Hoja1");
                CreateMainSheet(mainSheet);

                package.Save();
            }
        }

        private void CreateDetalleLectSheet(ExcelWorksheet worksheet)
        {
            // Encabezados
            worksheet.Cells["A1:F1"].LoadFromArrays(new object[][]
            {
            new[] {"nro_cert", "nom_unicom", "desc_tipo", "cant_lect", "lect0_con_an", "lect0_sin_an"}
            });

            // Datos actualizados
            var data = new[]
            {
            new object[] {4234, "CL - T2 ZONA NORTE", "Certificación Itinerario  T2", 6164, 234, 10},
            new object[] {4234, "CL - T2 ZONA SUR", "Certificación Itinerario  T2", 2466, 75, 1},
            new object[] {4234, "CL - LECTURAS T3", "Certificación Itinerario  T3", 2610, 239, 0},
            // ... agregar todas las filas restantes
            new object[] {4234, "CL - MALARGUE", "Certificación Itinerario T1", 5227, 44, 0}
        };

            worksheet.Cells["A2"].LoadFromArrays(data);

            // Formato
            worksheet.Cells[1, 1, data.Length + 1, 6].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[1, 1, data.Length + 1, 6].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[1, 1, data.Length + 1, 6].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[1, 1, data.Length + 1, 6].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[1, 1, 1, 6].Style.Font.Bold = true;
        }

        private void CreateMainSheet(ExcelWorksheet worksheet)
        {
            // Título principal actualizado
            worksheet.Cells["E1"].Value = "CERTIFICACIÓN PERIODO 31_03_2025 al 29_04_2025";
            worksheet.Cells["E1"].Style.Font.Bold = true;
            worksheet.Cells["E1:J1"].Merge = true;

            // Encabezados de tabla
            var headers = new[] { "CODIGO SAP", "cod_unicom", "OFICINA", "Realizadas",
                            "Sin Lectura - Con código de Anomalía",
                            "Sin Lectura - Sin código de Anomalía",
                            "CECO", "LECTURAS A CERTIFICAR", "IMPORTE A CERTIFICAR" };

            worksheet.Cells["A3:J3"].LoadFromArrays(new object[][] { headers });
            worksheet.Cells["A3:J3"].Style.Font.Bold = true;

            // Datos T1 con nuevos valores
            int row = 4;
            AddSection(worksheet, ref row, "T1", "C000001", new[]
            {
            new object[] {null, 1102, "CAPITAL", 33476, 1000, 0, 3220101, "SUM(E4:F14)", null},
            new object[] {null, 1202, "GUAYMALLEN", 46805, 813, 0, null, null, null},
            // ... completar todas las filas T1
        }, 773);

            // Datos T2 actualizados
            AddSection(worksheet, ref row, "T2", "C000013", new[]
            {
            new object[] {null, 1902, "ZONA NORTE", 6164, 234, 10, 3220102, "SUM(E16:F17)", null},
            new object[] {null, 3902, "ZONA SUR", 2466, 75, 1, null, null, null}
        }, 7138);

            // Sección Altura actualizada
            AddAlturaSection(worksheet, ref row);

            // Total general actualizado
            worksheet.Cells[$"J{row}"].Formula = "SUM(J15:J36)";
            worksheet.Cells[$"J{row}"].Style.Numberformat.Format = "#,##0.00";

            // Agregar sección final
            worksheet.Cells[$"F{row + 2}"].Value = "HOJA DE PEDIDO";
            worksheet.Cells[$"F{row + 2}"].Value = 45066841;
            worksheet.Cells[$"F{row + 3}"].Value = "HOJA DE ENTRADA";
            worksheet.Cells[$"F{row + 3}"].Value = 1000346085;

            // Formato general
            worksheet.Cells[3, 1, row + 3, 10].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[3, 1, row + 3, 10].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[3, 1, row + 3, 10].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            worksheet.Cells[3, 1, row + 3, 10].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            // Autoajustar columnas
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
        }

        private void AddSection(ExcelWorksheet worksheet, ref int row, string tipo, string codigoSap,
                              object[][] data, double multiplicador)
        {
            worksheet.Cells[row, 1].Value = tipo;
            worksheet.Cells[row, 2].Value = codigoSap;

            foreach (var rowData in data)
            {
                worksheet.Cells[row, 2].LoadFromArrays(new object[][] { rowData });
                row++;
            }

            // Fórmulas actualizadas
            worksheet.Cells[row, 5].Formula = $"SUM(E{row - data.Length + 1}:F{row - 1})";
            worksheet.Cells[row, 7].Formula = $"E{row} + G{row}";
            worksheet.Cells[row, 9].Formula = $"I{row - data.Length} * {multiplicador}";
            worksheet.Cells[row, 9].Style.Numberformat.Format = "#,##0.00";

            row++;
        }

        private void AddAlturaSection(ExcelWorksheet worksheet, ref int row)
        {
            // Sección Altura
            AddSection(worksheet, ref row, "Altura", "C001000", new[]
            {
            new object[] {null, 1102, "CAPITAL", 0, 0, 0, 3220101, "SUM(E26:F33)", null},
            // ... completar filas Altura
            new object[] {null, 3102, "SAN RAFAEL", 10, 0, 0, null, null, null}
        }, 9060);

            // Sección T3 Altura
            AddSection(worksheet, ref row, "T3 Altura", "C001000", new[]
            {
            new object[] {null, 1702, "ZONA NORTE", 66, 9, 0, 0, "SUM(E35:F35)", null}
        }, 9060);
        }



        public void CreateExcel3(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            FileInfo file = new FileInfo(filePath);

            using (var package = new ExcelPackage(file))
            {
                // Crear hoja principal
                var mainSheet = package.Workbook.Worksheets.Add("Hoja1");
                CreateMainSheet2(mainSheet);

                // Crear hojas vacías
                package.Workbook.Worksheets.Add("Hoja2");
                package.Workbook.Worksheets.Add("Hoja3");

                package.Save();
            }
        }

        private void CreateMainSheet2(ExcelWorksheet worksheet)
        {
            // Configurar anchos de columnas
            worksheet.Column(1).Width = 25;
            worksheet.Column(5).Width = 12;
            worksheet.Column(6).Width = 15;

            // Encabezado inicial
            worksheet.Cells["A1:L1"].LoadFromArrays(new object[][]
            {
            new[] {"Detalle T1", null, "NC", null, "FTL", "ENERO", null, null, null, null, null, null}
            });

            // Datos principales
            var data = new[]
            {
            new object[] {null, "Total", 227321, null, -11, null},
            new object[] {"Error de lecturas", "Cant. Generada", 92, null, -10, null},
            new object[] {null, "% Generacion", 0.000404714038738172, null, -9, null},
            new object[] {null, "$ multa", 0, null, -8, null},
            new object[] {"Lecturas Faltantes", "Cant. Generada", 19, null, -7, null},
            new object[] {null, "% Generacion", 8.358224713070944E-05, null, -6, null},
            new object[] {null, "$ multa", 0, null, -5, null},
            new object[] {"Novedades Mal y No Informadas", "Cant. Generada", 26, null, -4, null, null, null, null, null, 1549, null},
            new object[] {null, "% Generacion", 0.00011437570659991817, null, -3, null, null, null, null, null, 9915, null},
            new object[] {null, "$ multa", 0, null, -2, null, null, null, null, null, 46724, null},
            new object[] {"Reclamos", "Cant. Generada", 17, null, -1, null, null, null, null, null, 105535, null},
            new object[] {null, "% Generacion", 7.478411585379266E-05, null, 0, null, null, null, null, null, 50493, null},
            new object[] {null, "$ multa", 0, null, 1, null, null, null, null, null, 12262, null},
            new object[] {"Lecturas Fuera Plazo", "Cantidad", 13105, null, 2, null, null, null, null, null, 843, null},
            new object[] {null, "% Fuera Plazo", 0.05764975519199722, null, 3, null},
            new object[] {null, "$ multa", 230117.04300000003, null, 4, null},
            new object[] {"Bonificación Lectura", "Cantidad", 214216, null, 5, null},
            new object[] {null, "% Bonificado", 0.9423502448080028, null, 6, null},
            new object[] {null, "$ multa", -19247304.84600001, null, 7, null},
            new object[] {"MULTA", null, 230117.04300000003, null, 8, null},
            new object[] {"BONIFICACION", null, -19247304.84600001, null, 9, null},
            new object[] {"DESCUENTO", null, 0, null, 10, null},
            new object[] {"RESULTADO", null, -19017187.803000007, null, 11, null},
            new object[] {null, null, null, null, 12, null},
            new object[] {null, null, null, null, 13, null},
            new object[] {null, null, null, null, 14, null},
            new object[] {null, null, null, null, 15, null},
            new object[] {null, null, null, null, 16, null},
            new object[] {"Hoja de Pedido", null, null, null, 17, null},
            new object[] {"Hoja de Entrada", null, null, null, "Dentro Plazo", 0.9423502448080028},
            new object[] {null, null, null, null, "Fuera de plazo", 0.05764975519199722}
        };

            worksheet.Cells["A2"].LoadFromArrays(data);

            // Formato numérico
            worksheet.Cells["C2:C24"].Style.Numberformat.Format = "#,##0";
            worksheet.Cells["F30:F31"].Style.Numberformat.Format = "0.00%";
            worksheet.Cells["C16,C19,C20,C21,C22,C23"].Style.Numberformat.Format = "#,##0.00";

            // Formato de porcentaje
            worksheet.Cells["C4,C7,C10,C13,C15,C18"].Style.Numberformat.Format = "0.00%";

            // Formato de moneda
            worksheet.Cells["C6,C9,C12,C16,C19,C20,C21,C22,C23"].Style.Numberformat.Format = "[$$-en-US]#,##0.00";

            // Colores para valores negativos
            worksheet.Cells["C19,C21,C23"].Style.Font.Color.SetColor(Color.Red);

            // Bordes
            var borderRange = worksheet.Cells["A1:L31"];
            borderRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            borderRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            borderRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            borderRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;

            // Formato de texto
            worksheet.Cells["A1:A23"].Style.Font.Bold = true;
            worksheet.Cells["B2:B31"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

            // Combinar celdas
            worksheet.Cells["A1:L1"].Merge = true;
            worksheet.Cells["E30:F30"].Merge = true;
            worksheet.Cells["E31:F31"].Merge = true;

            // Autoajustar columnas
            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
        }


    }
}
