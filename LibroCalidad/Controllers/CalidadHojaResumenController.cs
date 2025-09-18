using Aspose.Cells.Charts;
using MultasLectura.Helpers;
using MultasLectura.LibroCalidad.Interfaces;
using MultasLectura.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static MultasLectura.LibroPrincipal.Controllers.LibroPrincipalController;

namespace MultasLectura.LibroCalidad.Controllers
{
    public class MetodoLineal
    {
        private string _descripcion;
        private int _cantidad;
        private double _importe;

        public string Descripcion { get { return _descripcion; } set { _descripcion = value; } }
        public int Cantidad { get { return _cantidad; } set { _cantidad = value; } }
        public double Importe { get { return _importe; } set { _importe = value; } }

        public MetodoLineal(string descripcion, int cantidad, double importe)
        {
            _descripcion = descripcion;
            _cantidad = cantidad;
            _importe = importe;
        }

        //public string Obtener

        public void CalcularImporteConBaremos(double baremo)
        {
            // if (!baremo.Equals(0))
            // {
            _importe = 2 * _cantidad * baremo;
            //  }

            // return _importe;
        }

        public void SumarCantidades(int cantidad2)
        {
            _cantidad += cantidad2;

        }

        public void SumarImportes(double importe2)
        {
            _importe += importe2;
        }

    }
    public class CalidadHojaResumenController : ICalidadHojaResumenController
    {

        public void CrearTablaBaremosMetas(ExcelWorksheet hoja, BaremoModel baremos, MetaModel metas, double propInconformidades)
        {
            Dictionary<string, double> datos = new()
            {
                ["T1 y T3"] = baremos.T1,
                ["T2"] = baremos.T2,
                ["Altura T1 y T3"] = baremos.AlturaT1,
                ["Meta"] = metas.Meta1,
                ["Meta 2"] = metas.Meta2,
                ["Obtenido"] = propInconformidades
            };

            var claves = datos.Keys;
            int primeraFilaEstilizar = 2;
            int numFila = 2;



            // string fechaStr = "Tue Jul 08 2025 21:53:06 GMT-0300 (hora estándar de Argentina)";
            // DateTimeOffset dto = DateTimeOffset.Parse(baremos.Fecha, System.Globalization.CultureInfo.InvariantCulture);
            // DateTime fecha = dto.DateTime;
            // Console.WriteLine(fecha.ToString("dd/MM/yyyy")); // 08/07/2025


            // DateTime fecha = DateTime.Parse(baremos.Fecha);
            // string fechaBaremos = fecha.ToString("dd/MM/yyyy");
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




            hoja.Cells["F1"].Value = $"Baremo Lectura desde el {dia}/{mes}/{anio}";

            for (int i = 0; i < datos.Count; i++)
            {
                hoja.Cells[$"F{numFila}"].Value = claves.ElementAt(i);

                if (i == datos.Count - 1)
                {
                    hoja.Cells[$"G{numFila}"].Formula = "B39/B40";
                } else
                {
                    hoja.Cells[$"G{numFila}"].Value = datos[claves.ElementAt(i)];
                }
               
              

                if (numFila >= 5 && numFila <= 7)
                {
                    hoja.Cells[$"G{numFila}"].Style.Numberformat.Format = "0.00%";
                }
                else
                {
                    LibroExcelHelper.FormatoMoneda(hoja.Cells[$"G{numFila}"]);
                }
                numFila++;
            }
            LibroExcelHelper.AplicarBordeGruesoARango(hoja.Cells[$"F{primeraFilaEstilizar}:G{numFila - 1}"]);
            LibroExcelHelper.FormatoNegrita(hoja.Cells[$"F1:G{numFila - 1}"]);

        }

        public void CrearTablaDinTipoEstado(ExcelWorksheet hoja, ExcelRange rango)
        {
            var pivotTable = hoja.PivotTables.Add(hoja.Cells["A1"], rango, "TablaDinamicaTipoEstado");
            pivotTable.RowFields.Add(pivotTable.Fields["tipo_certificacion"]);
            pivotTable.RowFields.Add(pivotTable.Fields["estado"]);
            pivotTable.DataFields.Add(pivotTable.Fields["nic"]);
            pivotTable.DataFields[0].Function = DataFieldFunctions.Count;
        }

        public Dictionary<string, double> CrearTablaMetodoLineal(ExcelWorksheet hojaDestino, ExcelWorksheet hojaOrigen, BaremoModel baremos)
        {
            List<MetodoLineal> datos = new() {
                new MetodoLineal("Certificación Itinerario T1", 0, 0),
                new MetodoLineal("Certificación Itinerario  T2", 0, 0),
                new MetodoLineal("Certificación Itinerario  T3", 0, 0),
                new MetodoLineal("Certificación Itinerario en Altura T1", 0, 0),
                new MetodoLineal("Certificación Itinerario en Altura T3", 0, 0),
            };

            int cantFilas = hojaOrigen.Dimension.Rows;
            int cantCol = hojaOrigen.Dimension.Columns;

            hojaDestino.Cells["A26"].Value = "Método Lineal";

            for (int row = 1; row <= cantFilas; row++)
            {
                for (int col = 1; col <= cantCol; col++)
                {
                    object cellValue = hojaOrigen.Cells[row, col].Value;
                    if (cellValue != null)
                    {
                        foreach (MetodoLineal dato in datos)
                        {
                            if (cellValue.ToString() == dato.Descripcion)
                            {
                                dato.Cantidad++;
                            }
                        }
                    }
                }
            }

            int comienzoTabla = 26;
            int numFila = 27;
            int totalCantidades = 0;
            double totalImportes = 0;

            foreach (MetodoLineal dato in datos)
            {
                if (dato.Descripcion.Contains("T1") || dato.Descripcion.Contains("T3"))
                {
                    if (dato.Descripcion.Contains("Altura"))
                    {
                        dato.CalcularImporteConBaremos(baremos.AlturaT1);
                    }
                    else
                    {
                        dato.CalcularImporteConBaremos(baremos.T1);
                    }
                }
                else
                {
                    dato.CalcularImporteConBaremos(baremos.T2);
                }

                totalCantidades += dato.Cantidad;
                totalImportes += dato.Importe;

                hojaDestino.Cells[$"A{numFila}"].Value = dato.Descripcion;
                // hojaDestino.Cells[$"B{numFila}"].Value = dato.Cantidad;
                hojaDestino.Cells[$"B{numFila}"].Formula = $"IFERROR(GETPIVOTDATA(\"nic\",$A$1,\"estado\",\"Error de Lectura\",\"tipo_certificacion\",\"{dato.Descripcion}\"),0)+IFERROR(GETPIVOTDATA(\"nic\",$A$1,\"estado\",\"Lectura Faltante\",\"tipo_certificacion\",\"{dato.Descripcion}\"),0)+IFERROR(GETPIVOTDATA(\"nic\",$A$1,\"estado\",\"Mal Informado\",\"tipo_certificacion\",\"{dato.Descripcion}\"),0)";
                // hojaDestino.Cells[$"C{numFila}"].Value = dato.Importe;
                hojaDestino.Cells[$"C{numFila}"].Formula = $"2*B{numFila}*{CeldaResBaremoXCantidad(dato)}";

                numFila++;
            }

            /* hojaDestino.Cells[$"B{numFila}"].Value = totalCantidades;
             hojaDestino.Cells[$"C{numFila}"].Value = totalImportes;*/
            hojaDestino.Cells[$"B{numFila}"].Formula = "+SUM(B27:B31)";
            hojaDestino.Cells[$"C{numFila}"].Formula = "+SUM(C27:C31)";

            LibroExcelHelper.FondoSolido(hojaDestino.Cells[$"C{numFila}"], Color.FromArgb(1, 252, 213, 180));
            LibroExcelHelper.FormatoMoneda(hojaDestino.Cells[$"C{comienzoTabla + 1}:C{numFila}"]);
            LibroExcelHelper.AplicarBordeFinoARango(hojaDestino.Cells[$"A{comienzoTabla}:C{numFila}"]);

            return new()
            {
                ["total"] = totalCantidades,
                ["importe"] = totalImportes
            };

        }

        string CeldaResBaremoXCantidad(MetodoLineal dato)
        {
            string celda = "G";

            if (dato.Descripcion.ToLower().Contains("altura"))
            {
                celda += "4"; 
            } else if (dato.Descripcion.ToLower().Contains("t2"))
            {
                celda += "3";
            } else
            {
                celda += "2";
            }

            return celda;
        }

        public Dictionary<string, double> CrearTablaTotales(ExcelWorksheet hoja, Dictionary<string, double> totales, Dictionary<string, int> reclamos, BaremoModel baremos, ExcelWorksheet hojaCalXOperario, double importeCertificacion)
        {
            int totalCertificado = LibroExcelHelper.SumarColumnaInt(hojaCalXOperario, 5, 2);

            List<MetodoLineal> datos = new() {
              // new MetodoLineal("Anomalias de Facturacion NC", int.Parse(totales["total"].ToString()), totales["importe"]),
                new MetodoLineal("Reclamos procedentes T1", reclamos["t1"], 0),
                new MetodoLineal("Reclamos procedentes T2", reclamos["t2"], 0),
                new MetodoLineal("Total de NC por Metodo Lineal (0,15% al 0,3%)", int.Parse(totales["total"].ToString()), totales["importe"]),
                new MetodoLineal("Totales Certificado", totalCertificado, importeCertificacion),
            };

            int totalCantidadesReclamos = reclamos["t1"] + reclamos["t2"];
            double totalImportesReclamos = 0;
            int numFila = 35;
            int filaInicial = 35;

            hoja.Cells[$"A{numFila}"].Value = "Descripción";
            hoja.Cells[$"B{numFila}"].Value = "TOTAL";
            hoja.Cells[$"C{numFila}"].Value = "IMPORTE";

            hoja.Cells[$"A{numFila + 1}"].Value = "Anomalias de Facturacion NC";
            hoja.Cells[$"B{numFila + 1}"].Formula = "+B32";
            hoja.Cells[$"C{numFila + 1}"].Formula = "+C32";

            foreach (MetodoLineal dato in datos)
            {
                if (dato.Descripcion.ToLower().Contains("reclamos"))
                {
                    if (dato.Descripcion.ToLower().Contains("t1"))
                    {
                        dato.CalcularImporteConBaremos(baremos.T1);
                    }
                    else
                    {
                        dato.CalcularImporteConBaremos(baremos.T2);
                    }

                    totalImportesReclamos += dato.Importe;
                }

            }


            numFila++;

            foreach (MetodoLineal dato in datos)
            {

                if (dato.Descripcion.ToLower().Contains("metodo lineal"))
                {
                    dato.SumarCantidades(totalCantidadesReclamos);
                    dato.SumarImportes(totalImportesReclamos);
                }

                 hoja.Cells[$"A{numFila + 1}"].Value = dato.Descripcion;

                if (dato.Descripcion.ToLower().Contains("totales certificado"))
                {
                    hoja.Cells[$"B{numFila + 1}"].Value = dato.Cantidad;
                    hoja.Cells[$"C{numFila + 1}"].Value = dato.Importe;
                }
                /* hoja.Cells[$"B{numFila + 1}"].Value = dato.Cantidad;
                 hoja.Cells[$"C{numFila + 1}"].Value = dato.Importe;*/

                 numFila++;


            }


          
            hoja.Cells["B37"].Value = reclamos["t1"];
            hoja.Cells["C37"].Formula = "+2*B37*G2";

           
            hoja.Cells["B38"].Value = reclamos["t2"];
            hoja.Cells["C38"].Formula = "+2*B38*G3";


            hoja.Cells["B39"].Formula = "+SUM(B36:B38)";
            hoja.Cells["C39"].Formula = "+SUM(C36:C38)";


          /*  hoja.Cells["B40"].Value = "+SUM(B36:B38)";
            hoja.Cells["C40"].Formula = "+SUM(C36:C38)";*/




            double propInconformidades = (double)datos.Where(dato => dato.Descripcion.ToLower().Contains("metodo lineal")).FirstOrDefault().Cantidad / datos.Where(dato => dato.Descripcion.ToLower().Contains("certificado")).FirstOrDefault().Cantidad;

            hoja.Cells[$"D{filaInicial + datos.Count + 1}"].Value = propInconformidades;

            LibroExcelHelper.FormatoPorcentaje(hoja.Cells[$"D{filaInicial + datos.Count + 1}"]);

            LibroExcelHelper.AplicarBordeFinoARango(hoja.Cells[$"A{filaInicial}:C{filaInicial + datos.Count + 1}"]);
            LibroExcelHelper.FormatoMoneda(hoja.Cells[$"C{filaInicial + 1}:C{filaInicial + datos.Count + 1}"]);
            LibroExcelHelper.FormatoNegrita(hoja.Cells[$"B{filaInicial + datos.Count + 1}:C{filaInicial + datos.Count + 1}"]);

            hoja.Cells.AutoFitColumns();

            return new()
            {

                ["propInconformidades"] = propInconformidades,
                ["totalMetLineal"] = datos.Where(dato => dato.Descripcion.ToLower().Contains("metodo lineal")).FirstOrDefault().Importe,
                ["totalReclamos"] = totalCantidadesReclamos
            };
        }

        public List<Grupo> CrearTablaValorFinalMulta(ExcelWorksheet hoja, double propInconformidades, double importeTotalMetLineal, double importeTotalCertificacion, MetaModel metas, ExcelWorksheet hojaDetalleCalidad, double totalReclamos)
        {
            double importeMultaFinal = 0;

            if (propInconformidades > metas.Meta1)
            {
                if (propInconformidades > metas.Meta2)
                {
                    double calcAuxiliar = (propInconformidades - metas.Meta1) / (0.01 - metas.Meta1);
                    importeMultaFinal = importeTotalCertificacion * Math.Pow(calcAuxiliar, 2);

                }
                else
                {
                    importeMultaFinal = propInconformidades * importeTotalMetLineal / propInconformidades;
                }
            }

            double propMultaSobreTotalCert = importeMultaFinal / importeTotalCertificacion;


            hoja.Cells["A44"].Value = "Multa";
            //  hoja.Cells["B44"].Value = importeMultaFinal;
            hoja.Cells["B44"].Formula = "IF((D40<=$G$5),0,IF(D40>$G$6,($C$40*((D40-$G$5)/(0.01-$G$5))^2),(D40*$C$39)/$G$7))";

            hoja.Cells["C44"].Value = propMultaSobreTotalCert;


            LibroExcelHelper.FormatoNegrita(hoja.Cells["A44:B44"]);
            LibroExcelHelper.FondoSolido(hoja.Cells["A44:B44"], Color.FromArgb(1, 255, 192, 0));
            LibroExcelHelper.FormatoMoneda(hoja.Cells["B44"]);
            LibroExcelHelper.FormatoPorcentaje(hoja.Cells["C44"]);
            LibroExcelHelper.AplicarBordeFinoARango(hoja.Cells["A44:B44"]);

            hoja.Cells.AutoFitColumns();

          List<Grupo> improcedencias =  CrearTablaFinalSegunInconf(hoja, hojaDetalleCalidad, importeMultaFinal, totalReclamos);

            return improcedencias;

        }

        private int CalcularCantidadInconf(ExcelWorksheet hojaBase, string desc)
        {
            int cantFilas = hojaBase.Dimension.Rows;


           int colEstado = LibroExcelHelper.ObtenerNumeroColumna(hojaBase, "estado");

            int cantFinal = 0;

            if (colEstado != -1)
            {
                for (int fila = 1; fila <= cantFilas; fila++)
                {
                    object cellValue = hojaBase.Cells[fila, colEstado].Value;
                    if (cellValue != null)
                    {
                        if (cellValue.ToString()!.ToLower().Contains(desc))
                        {
                            cantFinal++;



                        }
                    }
                }
            }

            return cantFinal;
        }

        private List<Grupo> CrearTablaFinalSegunInconf(ExcelWorksheet hoja, ExcelWorksheet hojaDetalleCalidad, double importeMultaFinal, double totalReclamos)
        {
            int celdaInicial = 26;
            int celdaSuma = 26;
            int totalInconf = 0;

            List<Tipo> tipos = new List<Tipo>();
            

            List<Dictionary<string, int>> lista = new List<Dictionary<string, int>>();
          
            Dictionary<string, int> dictLA = new()
                        {
                            { "Lectura Faltante", CalcularCantidadInconf(hojaDetalleCalidad, "lectura faltante")
        }
                        };
            Dictionary<string, int> dictEL = new()
                        {
                            { "Error de Lectura", CalcularCantidadInconf(hojaDetalleCalidad, "error de lectura" )   }
                        };
            Dictionary<string, int> dictMI = new()
                        {
                            { "Mal Informado", CalcularCantidadInconf(hojaDetalleCalidad, "mal informado" ) }
                        };
            Dictionary<string, int> dictR = new()
                        {
                            { "Reclamos", (int)totalReclamos }
                        };
            lista.Add(dictLA);
            lista.Add(dictEL);
            lista.Add(dictMI);
            lista.Add(dictR);

            for (int i= 0; i<lista.Count; i++)
            {
                totalInconf += lista[i].First().Value;
            }


            for (int i = 0; i < lista.Count; i++)
            {
                hoja.Cells[$"F{celdaSuma}"].Value = lista[i].Keys;



                if(lista[i].First().Key.ToLower() == "reclamos")
                {
                    hoja.Cells[$"G{celdaSuma}"].Formula = "B37+B38";
                } else
                {
                    hoja.Cells[$"G{celdaSuma}"].Formula = $"IFERROR(GETPIVOTDATA(\"nic\",$A$1,\"estado\",\"{lista[i].First().Key}\",\"tipo_certificacion\",\"Certificación Itinerario T1\"),0)+IFERROR(GETPIVOTDATA(\"nic\",$A$1,\"estado\",\"{lista[i].First().Key}\",\"tipo_certificacion\",\"Certificación Itinerario  T2\"),0)+IFERROR(GETPIVOTDATA(\"nic\",$A$1,\"estado\",\"{lista[i].First().Key}\",\"tipo_certificacion\",\"Certificación Itinerario  T3\"),0)+IFERROR(GETPIVOTDATA(\"nic\",$A$1,\"estado\",\"{lista[i].First().Key}\",\"tipo_certificacion\",\"Certificación Itinerario en Altura T1\"),0)+IFERROR(GETPIVOTDATA(\"nic\",$A$1,\"estado\",\"{lista[i].First().Key}\",\"tipo_certificacion\",\"Certificación Itinerario en Altura T3\"),0)";
                }

               
                hoja.Cells[$"H{celdaSuma}"].Formula = $"(G{celdaSuma}*B44)/B39";

                /*hoja.Cells[$"G{celdaSuma}"].Value = lista[i].Values;
                hoja.Cells[$"H{celdaSuma}"].Value = (lista[i].First().Value * importeMultaFinal) / totalInconf;*/

                tipos.Add(new Tipo(i + 1, lista[i].First().Key, lista[i].First().Value, (lista[i].First().Value * importeMultaFinal) / totalInconf));

                celdaSuma++;

            }

            hoja.Cells[$"F{celdaSuma}"].Value = "TOTAL";
            /* hoja.Cells[$"G{celdaSuma}"].Value = totalInconf;
             hoja.Cells[$"H{celdaSuma}"].Value = importeMultaFinal;*/

            hoja.Cells[$"G{celdaSuma}"].Formula = "+SUM(G26:G29)";
            hoja.Cells[$"H{celdaSuma}"].Formula = "+SUM(H26:H29)";

            LibroExcelHelper.FormatoNegrita(hoja.Cells[$"G{celdaSuma}:H{celdaSuma}"]);
            LibroExcelHelper.FormatoMoneda(hoja.Cells[$"H{celdaInicial}:H{celdaSuma}"]);



            var rangoAFormatear = hoja.Cells[$"F{celdaInicial}:H{celdaSuma}"];
            LibroExcelHelper.AplicarBordeFinoARango(rangoAFormatear);

            List<Grupo> improcedencias = new List<Grupo>
{
    new Grupo
    {
        Etiqueta = "Error de lectura",
        Datos = new List<Dato>
        {
          /* new Dato { Etiqueta = "Cant. Generada", Valor = tipos.Where(tipo => tipo.Id == 1).First().Cantidad},
            new Dato { Etiqueta = "% Generacion", Valor = 0.040},
             new Dato { Etiqueta = "$ multa", Valor = tipos.Where(tipo => tipo.Id == 1).First().Importe},*/
          new Dato { Etiqueta = "Cant. Generada", Valor = tipos.Where(tipo => tipo.Descripcion.ToLower().Contains("error") ).First().Cantidad},
            new Dato { Etiqueta = "% Generacion", Valor = 0.040},
             new Dato { Etiqueta = "$ multa", Valor = tipos.Where(tipo => tipo.Descripcion.ToLower().Contains("error") ).First().Importe},
        }
    },
     new Grupo
    {
        Etiqueta = "Lecturas Faltantes",
        Datos = new List<Dato>
         {
          /* new Dato { Etiqueta = "Cant. Generada", Valor = tipos.Where(tipo => tipo.Id == 2).First().Cantidad},
            new Dato { Etiqueta = "% Generacion", Valor = 0.040},
             new Dato { Etiqueta = "$ multa", Valor = tipos.Where(tipo => tipo.Id == 2).First().Importe},*/
             new Dato { Etiqueta = "Cant. Generada", Valor = tipos.Where(tipo => tipo.Descripcion.ToLower().Contains("faltante")).First().Cantidad},
            new Dato { Etiqueta = "% Generacion", Valor = 0.040},
             new Dato { Etiqueta = "$ multa", Valor = tipos.Where(tipo => tipo.Descripcion.ToLower().Contains("faltante")).First().Importe},
        }
    },
      new Grupo
    {
        Etiqueta = "Novedades Mal y No Informadas",
        Datos = new List<Dato>
         {
          /* new Dato { Etiqueta = "Cant. Generada", Valor = tipos.Where(tipo => tipo.Id == 3).First().Cantidad},
            new Dato { Etiqueta = "% Generacion", Valor = 0.040},
             new Dato { Etiqueta = "$ multa", Valor = tipos.Where(tipo => tipo.Id == 3).First().Importe},*/
           new Dato { Etiqueta = "Cant. Generada", Valor = tipos.Where(tipo => tipo.Descripcion.ToLower().Contains("mal")).First().Cantidad},
            new Dato { Etiqueta = "% Generacion", Valor = 0.040},
             new Dato { Etiqueta = "$ multa", Valor = tipos.Where(tipo => tipo.Descripcion.ToLower().Contains("mal")).First().Importe},
        }
    },
       new Grupo
    {
        Etiqueta = "Reclamos",
        Datos = new List<Dato>
        {
          /* new Dato { Etiqueta = "Cant. Generada", Valor = tipos.Where(tipo => tipo.Id == 4).First().Cantidad},
            new Dato { Etiqueta = "% Generacion", Valor = 0.040},
             new Dato { Etiqueta = "$ multa", Valor = tipos.Where(tipo => tipo.Id == 4).First().Importe},*/
           new Dato { Etiqueta = "Cant. Generada", Valor = tipos.Where(tipo => tipo.Descripcion.ToLower().Contains("reclamo")).First().Cantidad},
            new Dato { Etiqueta = "% Generacion", Valor = 0.040},
             new Dato { Etiqueta = "$ multa", Valor = tipos.Where(tipo => tipo.Descripcion.ToLower().Contains("reclamo")).First().Importe},
        }
    },
      /* new Grupo
    {
        Etiqueta = "Lecturas Fuera Plazo",
        Datos = new List<Dato>
        {
         //  new Dato { Etiqueta = "Cantidad", Valor = tipos.Where(tipo => tipo.Id == 5).First().Cantidad},
         new Dato { Etiqueta = "Cantidad", Valor =0},
            new Dato { Etiqueta = "% Fuera Plazo", Valor = 0.040},
           //  new Dato { Etiqueta = "$ multa", Valor = tipos.Where(tipo => tipo.Id == 5).First().Importe},
           new Dato { Etiqueta = "$ multa", Valor = 0},
        }
    },
       new Grupo
    {
        Etiqueta = "Bonificación Lectura",
        Datos = new List<Dato>
        {
       //    new Dato { Etiqueta = "Cantidad", Valor = tipos.Where(tipo => tipo.Id == 6).First().Cantidad},
            new Dato { Etiqueta = "Cantidad", Valor = 0},
            new Dato { Etiqueta = "% Bonificado", Valor = 0.040},
          //   new Dato { Etiqueta = "$ multa", Valor = tipos.Where(tipo => tipo.Id == 6).First().Importe},
           new Dato { Etiqueta = "$ multa", Valor = 0},
        }
    },*/

};

            return improcedencias;

        }




    }
}
