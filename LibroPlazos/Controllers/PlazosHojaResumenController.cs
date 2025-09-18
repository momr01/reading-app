using MultasLectura.Helpers;
using MultasLectura.LibroPlazos.Interfaces;
using MultasLectura.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using System.Reflection;
using static MultasLectura.LibroPrincipal.Controllers.LibroPrincipalController;

namespace MultasLectura.LibroPlazos.Controllers
{
    public class PlazosHojaResumenController : IPlazosHojaResumenController
    {
        private readonly BaremoModel _baremos;

        public PlazosHojaResumenController(BaremoModel baremos)
        {
            _baremos = baremos;
        }

        // public Dictionary<string, double> CrearTablaDatosPorTarifa(ExcelWorksheet hojaResumen, ExcelWorksheet hojaReclDetalles, ref int numFila)
        // public List<Ftl> CrearTablaDatosPorTarifa(ExcelWorksheet hojaResumen, ExcelWorksheet hojaReclDetalles, ref int numFila)
        public Dictionary<string, dynamic> CrearTablaDatosPorTarifa(ExcelWorksheet hojaResumen, ExcelWorksheet hojaReclDetalles, ref int numFila)
        {
            Dictionary<string, double> totales = new();
            List<Ftl> dias = new List<Ftl>();

            // Dictionary<string, dynamic> exportar = new Dictionary<string, dynamic>();
            Dictionary<string, dynamic> exportar = new()
{
    { "multa", 0.0 },
    { "bonificacion", 0.0 },
                {"ftl", new List<Ftl>() }
};


            List<TarifaPlazosModel> tarifas = new()
            {
                new("t1", "itinerario t1", "días hábiles t1", 'a', _baremos.T1, "Certificación Itinerario T1"),
                new("t2", "itinerario  t2", "días hábiles t2", 'f', _baremos.T2, "Certificación Itinerario  T2"),
                new("t3", "itinerario  t3", "días hábiles t3", 'k', _baremos.T3, "Certificación Itinerario  T3"),
                new("altura t1", "altura t1", "días hábiles altura t1", 'p', _baremos.AlturaT1, "Certificación Itinerario en Altura T1"),
                new("altura t3", "altura t3", "días hábiles altura t3", 'u', _baremos.AlturaT3, "Certificación Itinerario en Altura T3")
            };

            double totalBonifT1 = 0;
            double totalBonifT2 = 0;
            double totalMultaT1 = 0;
            double totalMultaT2 = 0;

            //int numFila = 1;

            foreach(TarifaPlazosModel tarifa in tarifas)
            {
                // Dictionary<string, double> valores = CrearTablaDatosPorTarifa(hojaResumen, hojaReclDetalles, tarifa, ref numFila);
               ValoresLibroPlazos valores = CrearTablaDatosPorTarifa(hojaResumen, hojaReclDetalles, tarifa, ref numFila);

                foreach(Ftl ftl in valores.DiasCantidades)
                {
                    if(dias.Count == 0)
                    {
                        dias.Add(ftl);
                    } else
                    {
                        //  if (dias.Where(dia => dia.Dia == ftl.Dia).First() != null)
                        if (dias.FirstOrDefault(dia => dia.Dia == ftl.Dia) != null)
                        {
                            dias.Where(dia => dia.Dia == ftl.Dia).First().Cantidad += ftl.Cantidad;

                        }
                        else
                        {
                            dias.Add(ftl);
                        }
                    }

                   
                }

                // totales["bonificacion"] = 

                /*  if (tarifa.Tarifa == "t2")
                  {
                      totalBonifT2 += valores["bonificacion"];
                      totalMultaT2 += valores["multa"];
                  } else
                  {
                      totalBonifT1 += valores["bonificacion"];
                      totalMultaT1 += valores["multa"];
                  }
              }*/

                /* if (tarifa.Tarifa == "t2")
                 {
                     totalBonifT2 += valores.ImporteBonificacion;
                     totalMultaT2 += valores["multa"];
                 }
                 else
                 {
                     totalBonifT1 += valores["bonificacion"];
                     totalMultaT1 += valores["multa"];
                 }*/


                //  CrearTablaFueraPlazoResumen(hojaResumen, totalMultaT1, totalMultaT2, ref numFila);
                //  CrearTablaBonificacionResumen();

                /* totales["bonificacionT1"] = totalBonifT1;
                 totales["bonificacionT2"] = totalBonifT2;
                 totales["multaT1"] = totalMultaT1;
                 totales["multaT2"] = totalMultaT2;*/

                //return totales;
                exportar["multa"] += valores.ImporteMulta;
                exportar["bonificacion"] += valores.ImporteBonificacion;
                 }

            exportar["ftl"] = dias;

            // return dias;
            return exportar;


        }
              

       /* private void CrearTablaFueraPlazoResumen(ExcelWorksheet hoja, double importeT1, double importeT2, ref int numFila)
        {
            double total = importeT1 + importeT2;
            numFila += 10;

            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"A{numFila}", "FUERA DE PLAZO DEL PERÍODO", Enums.TipoOpCelda.Value);
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"F{numFila}", "T1", Enums.TipoOpCelda.Value);
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"G{numFila}", importeT1, Enums.TipoOpCelda.Value);
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"F{numFila + 1}", "T2", Enums.TipoOpCelda.Value);
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"G{numFila + 1}", importeT2, Enums.TipoOpCelda.Value);
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"F{numFila + 2}", "Total", Enums.TipoOpCelda.Value);
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"G{numFila + 2}", total, Enums.TipoOpCelda.Value);

        }

        private void CrearTablaBonificacionResumen()
        {

        }*/

        private char ObtenerLetraSiguiente(char letra)
        {
            // Verificar si la letra es una letra del alfabeto
           // if (!char.IsLetter(letra))
           // {
          //      throw new ArgumentException("El carácter proporcionado no es una letra.");
          //  }

            // Obtener el código ASCII de la letra
            int asciiCode = (int)letra;

            // Incrementar el código ASCII y ajustar para el caso de 'z' y 'Z'
            if (letra == 'z')
            {
                asciiCode = (int)'a';
            }
            else if (letra == 'Z')
            {
                asciiCode = (int)'A';
            }
            else
            {
                asciiCode++;
            }

            // Convertir de nuevo a carácter
            return (char)asciiCode;
        }


        class ValoresLibroPlazos
        {
            public double ImporteBonificacion { get; set; }
            public double ImporteMulta { get; set; }
            public List<Ftl> DiasCantidades { get; set; }

            public ValoresLibroPlazos(double importeBonificacion, double importeMulta, List<Ftl> diasCantidades)
            {
                ImporteBonificacion = importeBonificacion;
                ImporteMulta = importeMulta;
                DiasCantidades = diasCantidades;
            }
        }



      //  private Dictionary<string,double> CrearTablaDatosPorTarifa(ExcelWorksheet hojaResumen, ExcelWorksheet hojaReclDetalles, TarifaPlazosModel tarifa, ref int numFila)
        private ValoresLibroPlazos CrearTablaDatosPorTarifa(ExcelWorksheet hojaResumen, ExcelWorksheet hojaReclDetalles, TarifaPlazosModel tarifa, ref int numFila)

        {
            Dictionary<string, double> importesFinales = new();
            int numFilaInicial = 3;
            //int numFila = 3;
            numFila = 3;

            string primLetra = tarifa.LetraInicial.ToString().ToUpper();
            char letra2 = ObtenerLetraSiguiente(tarifa.LetraInicial);
            string segLetra = letra2.ToString().ToUpper();
            char letra3 = ObtenerLetraSiguiente(letra2);
            string tercLetra = letra3.ToString().ToUpper();
            char letra4 = ObtenerLetraSiguiente(letra3);
            string cuarLetra = letra4.ToString().ToUpper();
            char letra5 = ObtenerLetraSiguiente(letra4);
            string quintaLetra = letra5.ToString().ToUpper();

            LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{primLetra}1", tarifa.Encabezado.ToUpper(), Enums.TipoOpCelda.Value);
          //  hojaResumen.Cells[$"{primLetra}1:{cuarLetra}1"].Merge = true;
            LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{primLetra}2", "FTL", Enums.TipoOpCelda.Value);
            LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{segLetra}2", "k", Enums.TipoOpCelda.Value);

            LibroExcelHelper.FormatoMergeCelda(hojaResumen, $"{primLetra}1:{cuarLetra}1");

            // hojaResumen.Cells[$"{primLetra}1"].Value = tarifa.Encabezado.ToUpper();

           
            //hojaResumen.Cells[$"{primLetra}2"].Value = "FTL";
            
           // hojaResumen.Cells[$"{segLetra}2"].Value = "k";
            LibroExcelHelper.FondoSolido(hojaResumen.Cells[$"{primLetra}2:{segLetra}2"], Color.FromArgb(1, 204, 255, 255 ));
           // hojaResumen.Cells[$"{tercLetra}2"].Value = "Qij";
            LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{tercLetra}2", "Qij", Enums.TipoOpCelda.Value);
            LibroExcelHelper.FormatoMergeCelda(hojaResumen, $"{tercLetra}2:{cuarLetra}2");
           // hojaResumen.Cells[$"{tercLetra}2:{cuarLetra}2"].Merge = true;
            LibroExcelHelper.FondoSolido(hojaResumen.Cells[$"{tercLetra}2"], Color.FromArgb(1, 255, 0, 0));

            LibroExcelHelper.AplicarBordeGruesoARango(hojaResumen.Cells[$"{primLetra}1:{cuarLetra}2"]);

            List<int> ftlLista = new()
          ;

           

            for (int i = -14; i < 18; i++)
            {
                ftlLista.Add(i);

            }


            int totalCantidades = 0;
            int totalEnPlazo = 0;
            int totalFueraDePlazo = 0;
            int totalK1 = 0;
            double importeK1 = 0;
            int totalK2 = 0;
            double importeK2 = 0;
            int totalK3oMas = 0;
            double importeK3oMas = 0;


            List<Ftl> ftlsQij = new List<Ftl>();
           /* ftlsQij.Add(new Ftl(-3, 1549));
            ftlsQij.Add(new Ftl(-2, 9915));
            ftlsQij.Add(new Ftl(-1, 46724));
            ftlsQij.Add(new Ftl(0, 105535));
            ftlsQij.Add(new Ftl(1, 50493));
            ftlsQij.Add(new Ftl(2, 12262));
            ftlsQij.Add(new Ftl(3, 843));*/



            foreach (int ftl in ftlLista)
            {
                LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{primLetra}{numFila}", ftl, Enums.TipoOpCelda.Value);
               // hojaResumen.Cells[$"{primLetra}{numFila}"].Value = ftl;

                int cantidad = CalcularCantAtrasos(hojaReclDetalles, ftl, tarifa.Descripcion);
                totalCantidades += cantidad;

               // hojaResumen.Cells[$"{tercLetra}{numFila}"].Value = cantidad;
                LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{tercLetra}{numFila}", cantidad, Enums.TipoOpCelda.Value);
                hojaResumen.Cells[$"{tercLetra}{numFila}"].Formula = $"IFERROR(GETPIVOTDATA(\"cant_sum\",Z1,\"tip_itin\",\"{tarifa.TipItin}\",\"atraso\",{ftl}),0)";


                ColorearRangoSegunNum(ftl, hojaResumen, numFila, primLetra, tercLetra);

                AplicarBordeParcialARango(numFila, numFilaInicial, ftlLista, hojaResumen, tarifa.LetraInicial);
                AplicarBordeParcialARango(numFila, numFilaInicial, ftlLista, hojaResumen, letra2);
                AplicarBordeParcialARango(numFila, numFilaInicial, ftlLista, hojaResumen, letra3);
                AplicarBordeParcialARango(numFila, numFilaInicial, ftlLista, hojaResumen, letra4);

                CargarDatosColK(ftl, hojaResumen, numFila, segLetra);

               // hojaResumen.Cells[$"{tercLetra}{numFila}:{cuarLetra}{numFila}"].Merge = true;
                LibroExcelHelper.FormatoMergeCelda(hojaResumen, $"{tercLetra}{numFila}:{cuarLetra}{numFila}");

                if (ftl >= -3 && ftl <= 1)
                {
                    totalEnPlazo += cantidad;
                } else
                {
                    totalFueraDePlazo += cantidad;
                }

                if(ftl == -4 || ftl == 2)
                {
                    totalK1 += cantidad;
                    //importeK1 += (double)cantidad * int.Parse(hojaResumen.Cells[$"{segLetra}{numFila}"].Value.ToString());
                    importeK1 += CalcularImporteK(cantidad, hojaResumen, segLetra, numFila);
                } else if(ftl == -5 || ftl == 3)
                {
                    totalK2 += cantidad;
                    importeK2 += CalcularImporteK(cantidad, hojaResumen, segLetra, numFila);
                } else if(ftl <= -6 || ftl >= 4)
                {
                    totalK3oMas += cantidad;
                    importeK3oMas += CalcularImporteK(cantidad, hojaResumen, segLetra, numFila);
                }


                if(cantidad > 0)
                {
                    ftlsQij.Add(new Ftl(ftl, cantidad));
                }




                numFila++;

            }

            // LibroExcelHelper.AplicarBordeGruesoARango(hojaResumen.Cells[$"A{numFilaInicial}:A{numFila}"]);
         //   hojaResumen.Cells[$"{primLetra}1:{cuarLetra}{numFila - 1}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
          //  hojaResumen.Cells[$"{primLetra}1:{cuarLetra}{numFila - 1}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            LibroExcelHelper.CentrarContenidoCelda(hojaResumen, $"{primLetra}1:{cuarLetra}{numFila - 1}");

            /* List<string> headersTabla2 = new()
             {
                 "Certificado", "Valor de Lectura", "Dentro Plazo", "Bonificación"
             };*/

            double porcentajeDentroPlazo = (double)totalEnPlazo / totalCantidades;

            double bonifica = porcentajeDentroPlazo >= 0.7 ? 1 : 0;
            double totalImpCert = (double)totalCantidades * tarifa.Baremos;

            List<double> dentroPlazo = new() { porcentajeDentroPlazo, double.Parse(totalEnPlazo.ToString()) };
            List<double> bonificacion = new() { bonifica,  0 };

            Dictionary<string, Generics> valoresTabla2 = new()
            {
                { "Certificado", new Generics( totalImpCert) },
                { "Valor de Lectura", new Generics(tarifa.Baremos) },
                { "Dentro Plazo", new Generics(dentroPlazo) },
                { "Bonificación", new Generics(bonificacion) }
            };

            double importeBonif = 0;

            var keys = valoresTabla2.Keys;

            for (int  i = 0; i < valoresTabla2.Count; i++)
            {
                var valor = valoresTabla2[keys.ElementAt(i)].Value;

                // hojaResumen.Cells[$"{primLetra}{numFila}"].Value = headersTabla2[i];
              //  hojaResumen.Cells[$"{primLetra}{numFila}"].Value = keys.ElementAt(i);

               // if()
                LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{primLetra}{numFila}", keys.ElementAt(i), Enums.TipoOpCelda.Value);
                //   LibroExcelHelper.FondoSolido(hojaResumen.Cells[$"{}"]);


                if(keys.ElementAt(i).ToLower().Contains("certificado"))
                {

                    hojaResumen.Cells[$"{tercLetra}{numFila}"].Formula = $"{tercLetra}{numFila + 1}*IFERROR(GETPIVOTDATA(\"cant_sum\",Z1,\"tip_itin\",\"{tarifa.TipItin}\"),0)";
                    LibroExcelHelper.FormatoMergeCelda(hojaResumen, $"{tercLetra}{numFila}:{cuarLetra}{numFila}");
                    hojaResumen.Cells[$"{tercLetra}{numFila}"].Style.Numberformat.Format = "$ #,##0";

                } else if(keys.ElementAt(i).ToLower().Contains("dentro plazo")) {
                    hojaResumen.Cells[$"{tercLetra}{numFila}"].Formula = $"IFERROR(SUM({tercLetra}14:{cuarLetra}18)/SUM({tercLetra}3:{cuarLetra}34),0)";
                    hojaResumen.Cells[$"{tercLetra}{numFila}"].Style.Font.Bold = true ;
                    hojaResumen.Cells[$"{tercLetra}{numFila}"].Style.Numberformat.Format = "0%";
                    hojaResumen.Cells[$"{cuarLetra}{numFila}"].Formula = $"SUM({tercLetra}14:{cuarLetra}18)";
                    hojaResumen.Cells[$"{cuarLetra}{numFila}"].Style.Numberformat.Format = "#,##0";
                    hojaResumen.Cells[$"{cuarLetra}{numFila}"].Style.Font.Bold = true;


                } else if(keys.ElementAt(i).ToLower().Contains("bonificación")) {

                    hojaResumen.Cells[$"{tercLetra}{numFila}"].Formula = $"IF({tercLetra}37>=0.7,\"Bonifica\",\"No Bonifica\")";
                    hojaResumen.Cells[$"{tercLetra}{numFila}"].Style.Font.Bold = true;

                    hojaResumen.Cells[$"{cuarLetra}{numFila}"].Formula = $"IF({tercLetra}37<0.7,0,(0.1*{tercLetra}35*(({tercLetra}37-0.7)/0.3)))";
                    hojaResumen.Cells[$"{cuarLetra}{numFila}"].Style.Font.Bold = true;
                    hojaResumen.Cells[$"{cuarLetra}{numFila}"].Style.Numberformat.Format = "$ #,##0";


                    if (valor is List<double>)
                    {
                        List<double> lista = valoresTabla2[keys.ElementAt(i)].GetValue<List<double>>();
                        if (i == 3)
                        {
                            string textoBonifica = "No Bonifica";
                            //  double importeBonif = 0;
                            ColorModel colores = new("no bonifica", Color.Red, Color.White);

                            if (lista[0] == 1)
                            {
                                textoBonifica = "Bonifica";
                                importeBonif = (0.1 * totalImpCert * ((porcentajeDentroPlazo - 0.7) / 0.3));
                                colores.Nombre = textoBonifica;
                                colores.Fondo = Color.FromArgb(1, 204, 255, 204);
                                colores.Letra = Color.Black;
                            }

                          /*  LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{tercLetra}{numFila}", textoBonifica,
                                Enums.TipoOpCelda.Value);
                            LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{cuarLetra}{numFila}", importeBonif,
                                Enums.TipoOpCelda.Value);*/
                            LibroExcelHelper.ColorFondoLetra(hojaResumen, letra3, numFila, colores);

                        }
                        else
                        {

                            // hojaResumen.Cells[$"{tercLetra}{numFila}"].Value = lista[0];
                            //hojaResumen.Cells[$"{cuarLetra}{numFila}"].Value = lista[1];
                          /*  LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{tercLetra}{numFila}", lista[0],
                                Enums.TipoOpCelda.Value);
                            LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{cuarLetra}{numFila}", lista[1],
                               Enums.TipoOpCelda.Value);*/
                        }


                    }

                }
                 else {     
                    

                   

                    if (valor is double)
                {
                    //double valorD = double.Parse(valor.ToString()!);
                    //  hojaResumen.Cells[$"{tercLetra}{numFila}"].Value = valoresTabla2[keys.ElementAt(i)].GetValue<double>();
                    LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{tercLetra}{numFila}", valoresTabla2[keys.ElementAt(i)].GetValue<double>(), 
                        Enums.TipoOpCelda.Value);
                    LibroExcelHelper.FormatoMergeCelda(hojaResumen, $"{tercLetra}{numFila}:{cuarLetra}{numFila}");
                } else if(valor is int)
                {
                    LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{tercLetra}{numFila}", valoresTabla2[keys.ElementAt(i)].GetValue<int>(), 
                        Enums.TipoOpCelda.Value);
                    LibroExcelHelper.FormatoMergeCelda(hojaResumen, $"{tercLetra}{numFila}:{cuarLetra}{numFila}");
                } else if(valor is List<double>)
                {
                    List<double> lista = valoresTabla2[keys.ElementAt(i)].GetValue<List<double>>();
                    if (i == 3)
                    {
                        string textoBonifica = "No Bonifica";
                      //  double importeBonif = 0;
                        ColorModel colores = new("no bonifica", Color.Red, Color.White);

                        if (lista[0] == 1)
                        {
                            textoBonifica = "Bonifica";
                            importeBonif = (0.1 * totalImpCert * ((porcentajeDentroPlazo - 0.7)/0.3));
                            colores.Nombre = textoBonifica;
                            colores.Fondo = Color.FromArgb(1, 204, 255, 204);
                            colores.Letra = Color.Black;
                        } 

                        LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{tercLetra}{numFila}", textoBonifica,
                            Enums.TipoOpCelda.Value);
                        LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{cuarLetra}{numFila}", importeBonif,
                            Enums.TipoOpCelda.Value);
                        LibroExcelHelper.ColorFondoLetra(hojaResumen, letra3, numFila, colores);

                    } else
                    {
                        
                        // hojaResumen.Cells[$"{tercLetra}{numFila}"].Value = lista[0];
                        //hojaResumen.Cells[$"{cuarLetra}{numFila}"].Value = lista[1];
                        LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{tercLetra}{numFila}", lista[0],
                            Enums.TipoOpCelda.Value);
                        LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{cuarLetra}{numFila}", lista[1],
                           Enums.TipoOpCelda.Value);
                    }
                   

                } else
                {
                    // hojaResumen.Cells[$"{tercLetra}{numFila}"].Value = "holiiis";
                    LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{tercLetra}{numFila}", 0,
                        Enums.TipoOpCelda.Value);
                    LibroExcelHelper.FormatoMergeCelda(hojaResumen, $"{tercLetra}{numFila}:{cuarLetra}{numFila}");

                }

                 }
               

                //ESTILOS CUADRO CERTIFICADO - VALOR DE LECTURA - DENTRO PLAZO - BONIFICACION 
                LibroExcelHelper.FormatoMergeCelda(hojaResumen, $"{primLetra}{numFila}:{segLetra}{numFila}");
               // LibroExcelHelper.FondoSolido(hojaResumen.Cells[$"{primLetra}{numFila}:{segLetra}{numFila}"], Color.FromArgb(1, 206, 255, 255));
                hojaResumen.Cells[$"{primLetra}{numFila}:{segLetra}{numFila}"].Style.Font.Bold = true;
                LibroExcelHelper.AplicarBordeGruesoARango(hojaResumen.Cells[$"{primLetra}{numFila}:{cuarLetra}{numFila}"]);
               // MessageBox.Show(hojaResumen.Cells[$"{primLetra}{numFila}"].Value.ToString());
                if(hojaResumen.Cells[$"{primLetra}{numFila}"].Value.ToString()!.ToLower() == "Dentro Plazo".ToLower())
                {
                    LibroExcelHelper.FondoSolido(hojaResumen.Cells[$"{primLetra}{numFila}:{segLetra}{numFila}"], Color.FromArgb(1, 255, 255, 153));
                } else
                {
                    LibroExcelHelper.FondoSolido(hojaResumen.Cells[$"{primLetra}{numFila}:{segLetra}{numFila}"], Color.FromArgb(1, 206, 255, 255));
                }





                if (i == 0 || i==1)
                {
                    LibroExcelHelper.FormatoMoneda(hojaResumen.Cells[$"{tercLetra}{numFila}"]);

                } else if(i== 2)
                {
                    LibroExcelHelper.FormatoPorcentaje(hojaResumen.Cells[$"{tercLetra}{numFila}"]);
                    //LibroExcelHelper.MIles
                } else
                {
                    LibroExcelHelper.FormatoMoneda(hojaResumen.Cells[$"{cuarLetra}{numFila}"]);
                }
                numFila++;
            }



            double porcentajeUnDia = 0.02;
            double porcentajeDosDias = porcentajeUnDia * 2;
            double porcentajeMasTresDias = porcentajeDosDias * 2.5;




            MultaPlazosDia tabla3 = new()
            {
                Tarifa = "t1",
               // ImporteFueraPlazo = 0,
                PorcentajeFueraPlazo = porcentajeUnDia + porcentajeDosDias + porcentajeMasTresDias,
              //  TotalMulta = 7889,

                Dia1 = new()
                {
                    Dia = 1,
                    PorcentajeIncremento = porcentajeUnDia,
                    PorcentajeObtenido = (double)totalK1 / totalCantidades,
                    // TotalMultaDia = (porcentajeUnDia * tarifa.Baremos) * importeK1,
                    TotalMultaDia = CalcularTotalMultaDia(porcentajeUnDia, tarifa.Baremos, importeK1),
                },

                 Dia2 = new()
                 {
                     Dia = 2,
                     PorcentajeIncremento = porcentajeDosDias,
                     PorcentajeObtenido = (double)totalK2/totalCantidades ,
                     TotalMultaDia = CalcularTotalMultaDia(porcentajeDosDias, tarifa.Baremos, importeK2),
                 },

                  Dia3Mas = new()
                  {
                      Dia = 3,
                      PorcentajeIncremento = porcentajeMasTresDias,
                      PorcentajeObtenido = (double)totalK3oMas/totalCantidades,
                      TotalMultaDia = CalcularTotalMultaDia(porcentajeMasTresDias, tarifa.Baremos, importeK3oMas),
                  }
            };

            // Obtener el tipo de la instancia
            Type tipo = tabla3.GetType();

            // Obtener todas las propiedades de la instancia
            PropertyInfo[] propiedades = tipo.GetProperties();

            // Recorrer las propiedades y sus valores

            numFila++;
            List<char> letras4 = new() { tarifa.LetraInicial, letra2, letra3, letra4 };

            for(int i= 1; i < 4; i++)
            {
                if(i== 1)
                {
                    hojaResumen.Cells[$"{primLetra}{numFila}"].Value = $"{i} día";
                    hojaResumen.Cells[$"{segLetra}{numFila}"].Value = 0.02;
                    hojaResumen.Cells[$"{tercLetra}{numFila}"].Formula = $"IFERROR(({tercLetra}19+{tercLetra}13)/SUM({tercLetra}3:{cuarLetra}34),0)";
                    hojaResumen.Cells[$"{cuarLetra}{numFila}"].Formula = $"({segLetra}40*{tercLetra}36)*(SUMPRODUCT({segLetra}13,{tercLetra}13)+SUMPRODUCT({segLetra}19,{tercLetra}19))";
                    hojaResumen.Cells[$"{cuarLetra}{numFila}"].Style.Numberformat.Format = "$ #,##0";

                    // hojaResumen.Cells[$"{segLetra}{numFila}"].Style.Numberformat.Format = "0%";
                    var rango = hojaResumen.Cells[$"{primLetra}{numFila}:{segLetra}{numFila}"];
                    rango.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rango.Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                } else if(i== 2)
                {
                    hojaResumen.Cells[$"{primLetra}{numFila}"].Value = $"{i} días";
                    hojaResumen.Cells[$"{segLetra}{numFila}"].Formula = $"{segLetra}{numFila - 1}*2";
                    hojaResumen.Cells[$"{tercLetra}{numFila}"].Formula = $"IFERROR(({tercLetra}20+{tercLetra}12)/SUM({tercLetra}3:{cuarLetra}34),0)";
                    hojaResumen.Cells[$"{cuarLetra}{numFila}"].Formula = $"({segLetra}41*{tercLetra}36)*(SUMPRODUCT({segLetra}12,{tercLetra}12)+SUMPRODUCT({segLetra}20,{tercLetra}20))";
                    hojaResumen.Cells[$"{cuarLetra}{numFila}"].Style.Numberformat.Format = "$ #,##0";

                    // hojaResumen.Cells[$"{segLetra}{numFila}"].Style.Numberformat.Format = "0%";
                    var rango = hojaResumen.Cells[$"{primLetra}{numFila}:{segLetra}{numFila}"];
                    rango.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rango.Style.Fill.BackgroundColor.SetColor(Color.Orange);
                } else
                {
                    hojaResumen.Cells[$"{primLetra}{numFila}"].Value = $">{i} días";
                    hojaResumen.Cells[$"{segLetra}{numFila}"].Formula = $"{segLetra}{numFila - 1}*2.5";
                    hojaResumen.Cells[$"{tercLetra}{numFila}"].Formula = $"IFERROR((SUM({tercLetra}3:{cuarLetra}11,{tercLetra}21:{cuarLetra}34))/SUM({tercLetra}3:{cuarLetra}34),0)";
                    hojaResumen.Cells[$"{cuarLetra}{numFila}"].Formula = $"({segLetra}42*{tercLetra}36)*(SUMPRODUCT({segLetra}3:{segLetra}11,{tercLetra}3:{tercLetra}11)+SUMPRODUCT({segLetra}21:{segLetra}34,{tercLetra}21:{tercLetra}34))";
                    hojaResumen.Cells[$"{cuarLetra}{numFila}"].Style.Numberformat.Format = "$ #,##0";

                    //  hojaResumen.Cells[$"{segLetra}{numFila}"].Style.Numberformat.Format = "0%";
                    var rango = hojaResumen.Cells[$"{primLetra}{numFila}:{segLetra}{numFila}"];
                    rango.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rango.Style.Fill.BackgroundColor.SetColor(Color.DarkOrange);

                }

                //LibroExcelHelper.AplicarBordeGruesoARango(hojaResumen.Cells[$"{letras4[i]}{numFila}"]);
                hojaResumen.Cells[$"{segLetra}{numFila}"].Style.Numberformat.Format = "0%";
                hojaResumen.Cells[$"{tercLetra}{numFila}"].Style.Numberformat.Format = "0.00%";
                hojaResumen.Cells[$"{cuarLetra}{numFila}"].Style.Numberformat.Format = "$ #,##0";

                AplicarBordes(hojaResumen,$"{primLetra}{numFila}",$"{primLetra}{numFila}");
                AplicarBordes(hojaResumen, $"{segLetra}{numFila}", $"{segLetra}{numFila}");
                AplicarBordes(hojaResumen, $"{tercLetra}{numFila}", $"{tercLetra}{numFila}");
                AplicarBordes(hojaResumen, $"{cuarLetra}{numFila}", $"{cuarLetra}{numFila}");

                var rangoNegrita = hojaResumen.Cells[$"{primLetra}{numFila}:{segLetra}{numFila}"];
                rangoNegrita.Style.Font.Bold = true;



                numFila++;
               
            }

            foreach (PropertyInfo propiedad in propiedades)
            {
               // if(propiedad.PropertyType == typeof(int))


                string nombre = propiedad.Name;
                object valor = propiedad.GetValue(tabla3)!;


                // MessageBox.Show(nombre);
                // object value = property.GetValue(person);
                // Console.WriteLine($"{name}: {value}");

                if (valor != null && valor.GetType() == typeof(DiaMulta))
                {
                    Type diaTipo = valor.GetType();
                    PropertyInfo[] diaProps = diaTipo.GetProperties();

                    foreach (PropertyInfo prop in diaProps)
                    {
                        //string nombreDia = prop.Name;
                       // object addressValue = prop.GetValue(valor);
                       // Console.WriteLine($"  {addressName}: {addressValue}");
                    }
                    for (int i = 0; i < diaProps.Length; i++)
                    {
                      //  LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{letras4[i]}{numFila}", diaProps[i].GetValue(valor), Enums.TipoOpCelda.Value);

                        //
                      //  LibroExcelHelper.AplicarBordeGruesoARango(hojaResumen.Cells[$"{letras4[i]}{numFila}"]);
                        if(i == 2)
                        {
                            hojaResumen.Cells[$"{letras4[i]}{numFila}"].Style.Numberformat.Format = "0.00%";
                        } else if(i== 3)
                        {
                            hojaResumen.Cells[$"{letras4[i]}{numFila}"].Style.Numberformat.Format = "$ #,##0.00";
                        }


                       

                    }

                  //  hojaResumen.Cells[$"{letras4[i]}{numFila}"]




                    //numFila++;
                } 



               /* if (nombre.ToLower().Contains("dia"))
                {
                    // Obtener el tipo de la instancia
                    Type tipoDia = propiedad.GetType();

                    // Obtener todas las propiedades de la instancia
                    PropertyInfo[] propsDia = tipoDia.GetProperties();

                  //  foreach(PropertyInfo prop in propsDia)
                  //  {
                   //     LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{}{fila}", valor, Enums.TipoOpCelda.Value);
                   // }
                   for(int i = 0; i < propsDia.Length; i++)
                    {
                        string jjj = propsDia[i].GetValue(tipoDia)!.ToString()!;
                        MessageBox.Show(jjj);
                        LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{letras4[i]}{numFila}", 0, Enums.TipoOpCelda.Value);
                    }


                   // FilaDiaTabla3(hojaResumen, new() { tarifa.LetraInicial, letra2, letra3, letra4}, numFila, 5);
                }*/

                
            }

            tabla3.CalcularImporteFueraPlazo();
            LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{primLetra}{numFila}", "Total Fuera Plazo", Enums.TipoOpCelda.Value);
            // LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{tercLetra}{numFila}", tabla3.PorcentajeFueraPlazo, Enums.TipoOpCelda.Value);
            // LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{cuarLetra}{numFila}", tabla3.ImporteFueraPlazo, Enums.TipoOpCelda.Value);
            hojaResumen.Cells[$"{tercLetra}{numFila}"].Formula = $"SUM({tercLetra}40:{tercLetra}42)";
            hojaResumen.Cells[$"{cuarLetra}{numFila}"].Formula = $"SUM({cuarLetra}40:{cuarLetra}42)";
            LibroExcelHelper.FormatoMergeCelda(hojaResumen, $"{primLetra}{numFila}:{segLetra}{numFila}");
            LibroExcelHelper.AplicarBordeGruesoARango(hojaResumen.Cells[$"{primLetra}{numFila}:{cuarLetra}{numFila}"]);
            LibroExcelHelper.FormatoNegrita(hojaResumen.Cells[$"{primLetra}{numFila}:{cuarLetra}{numFila}"]);
            hojaResumen.Cells[$"{primLetra}{numFila}:{segLetra}{numFila}"].Style.Fill.SetBackground(Color.LightBlue);
            hojaResumen.Cells[$"{cuarLetra}{numFila}"].Style.Numberformat.Format = "$ #,##0.00";
            hojaResumen.Cells[$"{tercLetra}{numFila}"].Style.Numberformat.Format = "0%";



            tabla3.CalcularTotalAMultar(importeBonif);
            LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{primLetra}{numFila + 1}", "Total a Multar", Enums.TipoOpCelda.Value);
            //  LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{tercLetra}{numFila + 1}", tabla3.TotalMulta, Enums.TipoOpCelda.Value);
            hojaResumen.Cells[$"{tercLetra}{numFila + 1}"].Formula = $"{cuarLetra}43-{cuarLetra}38";
            LibroExcelHelper.FormatoMergeCelda(hojaResumen, $"{primLetra}{numFila + 1}:{segLetra}{numFila + 1}");
            LibroExcelHelper.FormatoMergeCelda(hojaResumen, $"{tercLetra}{numFila + 1}:{cuarLetra}{numFila + 1}");
            LibroExcelHelper.AplicarBordeGruesoARango(hojaResumen.Cells[$"{primLetra}{numFila + 1}:{cuarLetra}{numFila + 1}"]);
            LibroExcelHelper.FormatoNegrita(hojaResumen.Cells[$"{primLetra}{numFila + 1}:{cuarLetra}{numFila + 1}"]);
            hojaResumen.Cells[$"{primLetra}{numFila + 1}:{segLetra}{numFila + 1}"].Style.Fill.SetBackground(Color.LightSteelBlue);
            hojaResumen.Cells[$"{tercLetra}{numFila + 1}:{cuarLetra}{numFila + 1}"].Style.Numberformat.Format = "$ #,##0.00";

            tabla3.DefinirEstadoFinal();
            //  LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{tercLetra}{numFila + 2}", tabla3.EstadoFinal, Enums.TipoOpCelda.Value);
            hojaResumen.Cells[$"{tercLetra}{numFila + 2}"].Formula = $"IF({tercLetra}44 < 0, \"Saldo a Favor del Periodo\",\"Multa del Periodo\")";
            LibroExcelHelper.ColorFondoLetra(hojaResumen, letra3, numFila + 2, tabla3.ColorEstado);
            LibroExcelHelper.FormatoMergeCelda(hojaResumen, $"{tercLetra}{numFila + 2}:{cuarLetra}{numFila + 2}");
            LibroExcelHelper.AplicarBordeGruesoARango(hojaResumen.Cells[$"{tercLetra}{numFila + 2}:{cuarLetra}{numFila + 2}"]);
            LibroExcelHelper.FormatoNegrita(hojaResumen.Cells[$"{tercLetra}{numFila + 2}:{cuarLetra}{numFila + 2}"]);


            hojaResumen.Workbook.Calculate();

            string estadoTarifa = hojaResumen.Cells[$"{tercLetra}{numFila + 2}"].GetValue<string>();

            if (estadoTarifa.ToLower().Contains("multa"))
            {
                hojaResumen.Cells[$"{tercLetra}{numFila + 2}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                hojaResumen.Cells[$"{tercLetra}{numFila + 2}"].Style.Fill.BackgroundColor.SetColor(Color.Red);
                hojaResumen.Cells[$"{tercLetra}{numFila + 2}"].Style.Font.Color.SetColor(Color.White);
            } else
            {
                hojaResumen.Cells[$"{tercLetra}{numFila + 2}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                hojaResumen.Cells[$"{tercLetra}{numFila + 2}"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(1, 204, 255, 204));
                hojaResumen.Cells[$"{tercLetra}{numFila + 2}"].Style.Font.Color.SetColor(Color.Black);
            }

           


            //  for( int i = 0; i < 5; i++)
            //{


            //}
            // LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{primLetra}{numFila + 6}", "FUERA DE PLAZO DEL PERIODO", Enums.TipoOpCelda.Value);
            // LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{primLetra}{numFila + 10}", "BONIFICACION DEL PERIODO", Enums.TipoOpCelda.Value);

            importesFinales["multa"] = tabla3.ImporteFueraPlazo;
            importesFinales["bonificacion"] = importeBonif;

            ValoresLibroPlazos valoresLibroPlazos = new ValoresLibroPlazos(importeBonif, tabla3.ImporteFueraPlazo, ftlsQij);



            //return importesFinales;
            return valoresLibroPlazos;

        }


        private void AplicarBordes(
        ExcelWorksheet worksheet,
        string startCell,
        string endCell,
        ExcelBorderStyle borderStyle = ExcelBorderStyle.Thick
    )
        {
            var rango = worksheet.Cells[$"{startCell}:{endCell}"];

            // Borde externo
            rango.Style.Border.Top.Style = borderStyle;
            rango.Style.Border.Bottom.Style = borderStyle;
            rango.Style.Border.Left.Style = borderStyle;
            rango.Style.Border.Right.Style = borderStyle;

            // Bordes internos
            /*rango.Style.Border.Horizontal.Style = borderStyle;
            rango.Style.Border.Vertical.Style = borderStyle;*/
        }

        private double CalcularImporteK(int cantidad, ExcelWorksheet hoja, string letra, int fila)
        {
            return (double)cantidad * int.Parse(hoja.Cells[$"{letra}{fila}"].Value.ToString()!);
        }

        private double CalcularTotalMultaDia(double porcentaje, double baremos, double importeK)
        {
            return (porcentaje * baremos) * importeK;
        }

       /* private void FilaDiaTabla3(ExcelWorksheet hoja, List<char> letras, int fila, double valor)
        {
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"{letras[0]}{fila}", valor, Enums.TipoOpCelda.Value);
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"{letras[1]}{fila}", valor, Enums.TipoOpCelda.Value);
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"{letras[2]}{fila}", valor, Enums.TipoOpCelda.Value);
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"{letras[3]}{fila}", valor, Enums.TipoOpCelda.Value);
        }*/


        private void CargarDatosColK(int ftl, ExcelWorksheet hojaResumen, int numFila, string letra)
        {
            if (ftl <= -3)
            {
               // hojaResumen.Cells[$"{letra}{numFila}"].Value = (ftl + 3) * -1;
                LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{letra}{numFila}", (ftl + 3) * -1,
                         Enums.TipoOpCelda.Value);

            }
            else if (ftl >= 2)
            {
               // hojaResumen.Cells[$"{letra}{numFila}"].Value = ftl - 1;
                LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{letra}{numFila}", ftl - 1,
                        Enums.TipoOpCelda.Value);
            }
            else
            {
                LibroExcelHelper.AsignarValorFormulaACelda(hojaResumen, $"{letra}{numFila}", 0, Enums.TipoOpCelda.Value);
            }
        }

        private void AplicarBordeParcialARango(int numFila, int numFilaInicial, List<int> ftl, ExcelWorksheet hojaResumen, char letraCelda)
        {
            string letra = letraCelda.ToString().ToUpper();
            string rango = $"{letra}{numFila}";

            if (numFila == numFilaInicial)
            {
                LibroExcelHelper.AplicarBordeParcialARango(hojaResumen, rango, Enums.TipoBordeParcial.SupDerIzq);
                /*hojaResumen.Cells[$"{letra}{numFila}"].Style.Border.Top.Style = ExcelBorderStyle.Thick;
                hojaResumen.Cells[$"{letra}{numFila}"].Style.Border.Left.Style = ExcelBorderStyle.Thick;
                hojaResumen.Cells[$"{letra}{numFila}"].Style.Border.Right.Style = ExcelBorderStyle.Thick;*/

            }
            else if (numFila == (ftl.Count + numFilaInicial - 1))
            {
                LibroExcelHelper.AplicarBordeParcialARango(hojaResumen, rango, Enums.TipoBordeParcial.InfDerIzq);
                /*hojaResumen.Cells[$"{letra}{numFila}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                hojaResumen.Cells[$"{letra}{numFila}"].Style.Border.Left.Style = ExcelBorderStyle.Thick;
                hojaResumen.Cells[$"{letra}{numFila}"].Style.Border.Right.Style = ExcelBorderStyle.Thick;*/

            }
            else
            {
                LibroExcelHelper.AplicarBordeParcialARango(hojaResumen, rango, Enums.TipoBordeParcial.CentroDerIzq);

              /*  hojaResumen.Cells[$"{letra}{numFila}"].Style.Border.Left.Style = ExcelBorderStyle.Thick;
                hojaResumen.Cells[$"{letra}{numFila}"].Style.Border.Right.Style = ExcelBorderStyle.Thick;*/
            }
        }

        private void ColorearRangoSegunNum(int num, ExcelWorksheet hojaResumen, int numFilaInicial, string letra1, string letra3)
        {
            if (num <= -6 || num >= 4)
            {
                LibroExcelHelper.FondoSolido(hojaResumen.Cells[$"{letra1}{numFilaInicial}:{letra3}{numFilaInicial}"], Color.FromArgb(1, 255, 102, 0));
            }
            else if (num == -5 || num == 3)
            {
                LibroExcelHelper.FondoSolido(hojaResumen.Cells[$"{letra1}{numFilaInicial}:{letra3}{numFilaInicial}"], Color.FromArgb(1, 255, 204, 153));
            }
            else if (num == -4 || num == 2)
            {
                LibroExcelHelper.FondoSolido(hojaResumen.Cells[$"{letra1}{numFilaInicial}:{letra3}{numFilaInicial}"], Color.FromArgb(1, 255, 255, 153));
            }
            else
            {
                LibroExcelHelper.FondoSolido(hojaResumen.Cells[$"{letra1}{numFilaInicial}:{letra3}{numFilaInicial}"], Color.FromArgb(1, 204, 255, 204));
            }
            /* List<int> ftl = new()
          ;

             for (int i = -14; i < 18; i++)
             {
                 ftl.Add(i);

             }

             foreach (int num in ftl)
             {
                if(num <= -6 || num >= 4)
                 {
                     LibroExcelHelper.FondoSolido();
                 }

             }*/



        }



        private int CalcularCantAtrasos(ExcelWorksheet hojaBase, int atraso, string tipo)
        {
            int cantFilas = hojaBase.Dimension.Rows;
          

            int colTipItin = LibroExcelHelper.ObtenerNumeroColumna(hojaBase, "tip_itin");
            int colAtraso = LibroExcelHelper.ObtenerNumeroColumna(hojaBase, "atraso");
            int colCantSum = LibroExcelHelper.ObtenerNumeroColumna(hojaBase, "cant_sum");

            int cantFinal = 0;

            if (colTipItin != -1 && colAtraso != -1 && colCantSum != -1)
            {
                for (int fila = 1; fila <= cantFilas; fila++)
                {
                    object cellValue = hojaBase.Cells[fila, colTipItin].Value;
                    if (cellValue != null)
                    {
                        if (cellValue.ToString()!.ToLower().Contains(tipo))
                        {
                            int atrasoBase = int.Parse(hojaBase.Cells[fila, colAtraso].Value.ToString()!);

                            if(atrasoBase == atraso)
                            {
                                cantFinal += int.Parse(hojaBase.Cells[fila, colCantSum].Value.ToString()!);
                            }
                           


                        }
                    }
                }
            }

            return cantFinal;
        }

       

        public void CrearTablaDinCertAtrasoTotal(ExcelWorksheet hoja, ExcelRange rango)
        {
            var pivotTable = hoja.PivotTables.Add(hoja.Cells["Z1"], rango, "TablaDinamicaCertAtrasoTotal");
            pivotTable.RowFields.Add(pivotTable.Fields["tip_itin"]);
            var atraso = pivotTable.RowFields.Add(pivotTable.Fields["atraso"]);
            pivotTable.DataFields.Add(pivotTable.Fields["cant_sum"]);
            pivotTable.DataFields[0].Function = DataFieldFunctions.Sum;

            atraso.Sort = eSortType.Ascending;
        }

        //  public void CrearTablaImportesFinales(ExcelWorksheet hoja, double importeT1, double importeT2, ref int numFila)
        public void CrearTablaImportesFinales(ExcelWorksheet hoja, ref int numFila)
        {
           // double total = importeT1 + importeT2;
            numFila += 10;
            /* Dictionary<string, double> fueraPlazo = new Dictionary<string, double>() {


                 { "t1", importeT1 } , {"t2", importeT2} , {"total", total}


             };

             Dictionary<string, double> bonificacion = new Dictionary<string, double>() {


                 { "t1", 5666 } , {"t2", 488} , {"total", 6512}


             };*/

            Dictionary<string, string> fueraPlazo = new Dictionary<string, string>() {


                { "t1", "D43+S43+N43+X43" } , {"t2", "I43"} , {"total", "SUM(G53:G54)"}


            };

            Dictionary<string, string> bonificacion = new Dictionary<string, string>() {


                { "t1", "D38+S38+N38+X38" } , {"t2", "I38"} , {"total", "-SUM(G57:G58)"}


            };

            TablaImportesFinalesLayout(hoja, numFila, "fuera de plazo del período", fueraPlazo, Color.FromArgb(1, 255, 255, 153));
            TablaImportesFinalesLayout(hoja, numFila + 4, "bonificación del período", bonificacion, Color.FromArgb(1, 204, 255, 204));


            // 👇 forzamos a calcular las fórmulas
            hoja.Workbook.Calculate();

            // Obtenemos los valores de las celdas donde quedó cada "total"
            double totalFueraPlazo = hoja.Cells[$"G{numFila + 2}"].GetValue<double>();
            double totalBonificacion = hoja.Cells[$"G{numFila + 6}"].GetValue<double>();

          //  return (totalFueraPlazo, totalBonificacion);


        }


        private void TablaImportesFinalesLayout(ExcelWorksheet hoja, int numFila, string desc, Dictionary<string, string> importes, Color color)
        {
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"A{numFila}", desc.ToUpper(), Enums.TipoOpCelda.Value);
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"F{numFila}", "T1", Enums.TipoOpCelda.Value);
            // LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"G{numFila}", importes["t1"], Enums.TipoOpCelda.Value);
             LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"G{numFila}", importes["t1"], Enums.TipoOpCelda.Formula);
            //hoja.Cells[$"G{numFila}"].Formula = "D43+S43+N43+X43";
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"F{numFila + 1}", "T2", Enums.TipoOpCelda.Value);
            //LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"G{numFila + 1}", importes["t2"], Enums.TipoOpCelda.Value);
             LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"G{numFila + 1}", importes["t2"], Enums.TipoOpCelda.Formula);
            //hoja.Cells[$"G{numFila + 1}"].Formula = "I43";
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"F{numFila + 2}", "Total", Enums.TipoOpCelda.Value);
            //LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"G{numFila + 2}", importes["total"], Enums.TipoOpCelda.Value);
            LibroExcelHelper.AsignarValorFormulaACelda(hoja, $"G{numFila + 2}", importes["total"], Enums.TipoOpCelda.Formula);
            //hoja.Cells[$"G{numFila + 2}"].Formula = $"SUM(G{numFila}:G{numFila + 2})";

            LibroExcelHelper.FormatoMergeCelda(hoja, $"A{numFila}:E{numFila + 2}");
            // LibroExcelHelper.CentrarContenidoCelda(hoja, $"A{numFila}:E{numFila + 2}");
            LibroExcelHelper.CentrarContenidoCelda(hoja, $"A{numFila}:G{numFila + 2}");
            LibroExcelHelper.FondoSolido(hoja.Cells[$"A{numFila}:E{numFila + 2}"], color);
            LibroExcelHelper.AplicarBordeGruesoARango(hoja.Cells[$"A{numFila}:G{numFila + 2}"]);
            LibroExcelHelper.FormatoNegrita(hoja.Cells[$"A{numFila}:G{numFila + 2}"]);

            LibroExcelHelper.FondoSolido(hoja.Cells[$"G{numFila}:G{numFila + 1}"], Color.FromArgb(1, 193, 192, 187));
            LibroExcelHelper.FondoSolido(hoja.Cells[$"G{numFila + 2}"], Color.FromArgb(1, 0, 255, 1));

            LibroExcelHelper.FormatoMoneda(hoja.Cells[$"G{numFila}:G{numFila + 2}"]);
        }


    }
}
