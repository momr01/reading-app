using MultasLectura.Enums;
using MultasLectura.Helpers;
using MultasLectura.LibroPlazos.Controllers;
using MultasLectura.LibroPlazos.Interfaces;
using MultasLectura.LibroPrincipal.Controllers;
using MultasLectura.LibroPrincipal.Interfaces;
using MultasLectura.Models;
using Newtonsoft.Json;
using Org.BouncyCastle.Asn1.Cmp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static MultasLectura.Views.GenerarTodo;

namespace MultasLectura.Views
{
    public partial class GenerarTodo : Form
    {
        private readonly HttpClient _httpClient = new HttpClient();

        private ILibroPrincipalController _principalController;
        private readonly ILibroPlazosController _plazosController;
        private BaremoModel _baremos = new();
        private readonly MetaModel _metas = new();
        public GenerarTodo()
        {
            InitializeComponent();

            Shown += async (s, e) => await CargarDatosBaremosAsync();
            Shown += async (s, e) => await CargarDatosMetasAsync();

            btnCargarBaremos.Click += async (s, e) => await CargarDatosBaremosAsync();
            btnCargarMetas.Click += async (s, e) => await CargarDatosMetasAsync();

            //ArchivoTextoHelper.VerificarExisteArchivoBaremos(_baremos);
            /* _baremos.T1 = _baremos.T3 = 732.48;
             _baremos.T2 = 6761.35;
             _baremos.AlturaT1 = _baremos.AlturaT3 = 8582.59;*/

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            //  _principalController = new LibroPrincipalController(_baremos, _metas, dateTimePicker1.Value.Date, dateTimePicker2.Value.Date);
            _plazosController = new LibroPlazosController(_baremos);



            // Configuración inicial de los DateTimePicker
            dateTimePicker1.Format = DateTimePickerFormat.Short;
            dateTimePicker2.Format = DateTimePickerFormat.Short;

            // Suscribir eventos
            dateTimePicker1.ValueChanged += DatePickers_ValueChanged;
            dateTimePicker2.ValueChanged += DatePickers_ValueChanged;

            btnGenerarLibros.Enabled = false;
            btnGenerarLibros.Visible = false;


            ViewHelper.DragDropTextBox(txtCertificacion, txtCertificacion_DragEnter!, txtCertificacion_DragDrop!);
            ViewHelper.DragDropTextBox(txtPlazos, txtPlazos_DragEnter!, txtPlazos_DragDrop!);
            ViewHelper.DragDropTextBox(txtCalidad, txtCalidad_DragEnter!, txtCalidad_DragDrop!);
            ViewHelper.DragDropTextBox(txtCalidadOperarios, txtCalidadOperarios_DragEnter!, txtCalidadOperarios_DragDrop!);
            ViewHelper.DragDropTextBox(txtReclamos, txtReclamos_DragEnter!, txtReclamos_DragDrop!);


        }

        private void btnGenerarLibros_Click(object sender, EventArgs e)
        {
            _principalController = new LibroPrincipalController(_baremos, _metas, dateTimePicker1.Value.Date, dateTimePicker2.Value.Date);

            string rutaCertificacion = txtCertificacion.Lines.FirstOrDefault()!;
            string rutaPlazos = txtPlazos.Lines.FirstOrDefault()!;
            string rutaCalidad = txtCalidad.Lines.FirstOrDefault()!;
            string rutaCalidadOperarios = txtCalidadOperarios.Lines.FirstOrDefault()!;
            string rutaReclamos = txtReclamos.Lines.FirstOrDefault()!;



            if (string.IsNullOrEmpty(rutaCertificacion)
                || string.IsNullOrEmpty(rutaPlazos)
               || string.IsNullOrEmpty(rutaCalidad)
                || string.IsNullOrEmpty(rutaCalidadOperarios)
                || string.IsNullOrEmpty(rutaReclamos)

                )
            {
                LibroExcelHelper.MostrarMensaje("Debe cargar todos los archivos solicitados.", true);
            }
            else
            {

                List<Dictionary<string, string>> rutas = new()
                    {
                       new() { [RutaArchivo.Certificacion.ToString()] = rutaCertificacion },
                        new() { [RutaArchivo.Plazos.ToString()] = rutaPlazos },
                         new() { [RutaArchivo.CalidadXOperario.ToString()] = rutaCalidadOperarios },
                          new() { [RutaArchivo.Reclamos.ToString()] = rutaReclamos },
                           new() { [RutaArchivo.CalidadDetalles.ToString()] = rutaCalidad }

                    };

                List<Dictionary<string, string>> rutasValidas = LibroExcelHelper.ProcesarPathArchivos(rutas, "Certificacion 28_02_2025 al 28_03_2025.xlsx");

                if (rutasValidas.Count == rutas.Count + 1)
                {
                    _principalController.GenerarLibros(
                        LibroExcelHelper.ObtenerValorPorClave(rutasValidas, RutaArchivo.Certificacion.ToString()),
                         LibroExcelHelper.ObtenerValorPorClave(rutasValidas, RutaArchivo.Plazos.ToString()),
                         LibroExcelHelper.ObtenerValorPorClave(rutasValidas, RutaArchivo.Guardar.ToString()),
                           LibroExcelHelper.ObtenerValorPorClave(rutasValidas, RutaArchivo.Plazos.ToString()),
                           LibroExcelHelper.ObtenerValorPorClave(rutasValidas, RutaArchivo.CalidadXOperario.ToString()),
                            LibroExcelHelper.ObtenerValorPorClave(rutasValidas, RutaArchivo.Reclamos.ToString()),
                             LibroExcelHelper.ObtenerValorPorClave(rutasValidas, RutaArchivo.CalidadDetalles.ToString())
                        );


                    /*  _plazosController.GenerarLibroPlazos(
                        //LibroExcelHelper.ObtenerValorPorClave(rutasValidas, RutaArchivo.PlazosDetalle.ToString()),
                         LibroExcelHelper.ObtenerValorPorClave(rutasValidas, RutaArchivo.Plazos.ToString()),
                         LibroExcelHelper.ObtenerValorPorClave(rutasValidas, RutaArchivo.Guardar.ToString())
                        );*/



                    MessageBox.Show("Archivos generados con éxito!", "ÉXITO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }


            }

        }

        private void btnCertificacion_Click(object sender, EventArgs e)
        {
            LibroExcelHelper.IniciarProcesoCarga(txtCertificacion);
        }

        private void btnPlazos_Click(object sender, EventArgs e)
        {
            LibroExcelHelper.IniciarProcesoCarga(txtPlazos);
        }

        private void btnCalidad_Click(object sender, EventArgs e)
        {
            LibroExcelHelper.IniciarProcesoCarga(txtCalidad);
        }

        private void btnCalidadOperarios_Click(object sender, EventArgs e)
        {
            LibroExcelHelper.IniciarProcesoCarga(txtCalidadOperarios);
        }

        private void btnReclamos_Click(object sender, EventArgs e)
        {
            LibroExcelHelper.IniciarProcesoCarga(txtReclamos);
        }





        private async Task CargarDatosBaremosAsync()
        {
            btnCargarBaremos.Enabled = false;
            lblEstadoBaremos.Text = "Cargando datos...";
            //dataGridView1.DataSource = null;

            try
            {
                var baremos = await ObtenerBaremosAsync();
                // dataGridView1.DataSource = tarifas;
                lblEstadoBaremos.Text = "Datos cargados correctamente.";
                lblEstadoBaremos.ForeColor = Color.DarkGreen;

                // Promedio
                //  var promedio = baremos.Average(t => t.Valor);
                // lblPromedio.Text = $"Promedio: {promedio:N2}";

                // Activas
                // var activas = baremos.Count(t => t.IsActive);
                //  lblActivas.Text = $"Activas: {activas}";

                // Última fecha
                // var ultimaFecha = tarifas.Max(t => t.Date);
                // lblUltimaFecha.Text = $"Última fecha: {ultimaFecha:dd/MM/yyyy}";

                baremosT1.Text = baremos.Where(t => t.Tarifa.ToLower() == "t1").First().Valor.ToString();
                baremosT2.Text = baremos.Where(t => t.Tarifa.ToLower() == "t2").First().Valor.ToString();
                baremosT3.Text = baremos.Where(t => t.Tarifa.ToLower() == "t3").First().Valor.ToString();
                baremosAlturaT1.Text = baremos.Where(t => t.Tarifa.ToLower() == "altura t1").First().Valor.ToString();
                baremosAlturaT3.Text = baremos.Where(t => t.Tarifa.ToLower() == "altura t3").First().Valor.ToString();

                _baremos.T1 = baremos.Where(t => t.Tarifa.ToLower() == "t1").First().Valor;
                _baremos.T2 = baremos.Where(t => t.Tarifa.ToLower() == "t2").First().Valor;
                _baremos.T3 = baremos.Where(t => t.Tarifa.ToLower() == "t3").First().Valor;
                _baremos.AlturaT1 = baremos.Where(t => t.Tarifa.ToLower() == "altura t1").First().Valor;
                _baremos.AlturaT3 = baremos.Where(t => t.Tarifa.ToLower() == "altura t3").First().Valor;
                _baremos.FechaCreacion = baremos.Where(t => t.Tarifa.ToLower() == "t1").First().DataEntry;
                _baremos.FechaEdicion = baremos.Where(t => t.Tarifa.ToLower() == "t1").First().DateEdit;


            }
            catch (Exception ex)
            {
                lblEstadoBaremos.Text = $"Error: {ex.Message}";
                lblEstadoBaremos.ForeColor = Color.Red;
                // MessageBox.Show(ex.Message);
            }
            finally
            {
                btnCargarBaremos.Enabled = true;
            }
        }


        private async Task CargarDatosMetasAsync()
        {
            btnCargarMetas.Enabled = false;
            lblMetasAPI.Text = "Cargando datos...";
            //dataGridView1.DataSource = null;

            try
            {
                var metas = await ObtenerMetasAsync();
                // dataGridView1.DataSource = tarifas;
                lblMetasAPI.Text = "Datos cargados correctamente.";
                lblMetasAPI.ForeColor = Color.DarkGreen;

                // Promedio
                //  var promedio = baremos.Average(t => t.Valor);
                // lblPromedio.Text = $"Promedio: {promedio:N2}";

                // Activas
                // var activas = baremos.Count(t => t.IsActive);
                //  lblActivas.Text = $"Activas: {activas}";

                // Última fecha
                // var ultimaFecha = tarifas.Max(t => t.Date);
                // lblUltimaFecha.Text = $"Última fecha: {ultimaFecha:dd/MM/yyyy}";

                meta1.Text = $"{metas.Where(t => t.Number == 1).First().Valor * 100}%";
                meta2.Text = $"{metas.Where(t => t.Number == 2).First().Valor * 100}%";

                /* baremosT1.Text = baremos.Where(t => t.Tarifa.ToLower() == "t1").First().Valor.ToString();
                 baremosT2.Text = baremos.Where(t => t.Tarifa.ToLower() == "t2").First().Valor.ToString();
                 baremosT3.Text = baremos.Where(t => t.Tarifa.ToLower() == "t3").First().Valor.ToString();
                 baremosAlturaT1.Text = baremos.Where(t => t.Tarifa.ToLower() == "altura t1").First().Valor.ToString();
                 baremosAlturaT3.Text = baremos.Where(t => t.Tarifa.ToLower() == "altura t3").First().Valor.ToString();*/

                _metas.Meta1 = metas.Where(t => t.Number == 1).First().Valor;
                _metas.Meta2 = metas.Where(t => t.Number == 2).First().Valor;

                /* _baremos.T1 = baremos.Where(t => t.Tarifa.ToLower() == "t1").First().Valor;
                 _baremos.T2 = baremos.Where(t => t.Tarifa.ToLower() == "t2").First().Valor;
                 _baremos.T3 = baremos.Where(t => t.Tarifa.ToLower() == "t3").First().Valor;
                 _baremos.AlturaT1 = baremos.Where(t => t.Tarifa.ToLower() == "altura t1").First().Valor;
                 _baremos.AlturaT3 = baremos.Where(t => t.Tarifa.ToLower() == "altura t3").First().Valor;*/


            }
            catch (Exception ex)
            {
                lblMetasAPI.Text = $"Error: {ex.Message}";
                lblMetasAPI.ForeColor = Color.Red;
                // MessageBox.Show(ex.Message);
            }
            finally
            {
                btnCargarMetas.Enabled = true;
            }
        }



        private async Task<List<Baremo>> ObtenerBaremosAsync()
        {
            // string url = "https://tudominio.com/api/tarifas"; // ← Cambiá esto por tu endpoint real
            string url = "http://localhost:3500/api/baremos/getAll";
            var response = await _httpClient.GetAsync(url);
            response.EnsureSuccessStatusCode();
            var json = await response.Content.ReadAsStringAsync();

            // ← Deserializar con Json.NET (Newtonsoft.Json)
            return JsonConvert.DeserializeObject<List<Baremo>>(json);
        }


        private async Task<List<Meta>> ObtenerMetasAsync()
        {
            // string url = "https://tudominio.com/api/tarifas"; // ← Cambiá esto por tu endpoint real
            string url = "http://localhost:3500/api/metas/getAll";
            var response = await _httpClient.GetAsync(url);
            response.EnsureSuccessStatusCode();
            var json = await response.Content.ReadAsStringAsync();

            // ← Deserializar con Json.NET (Newtonsoft.Json)
            return JsonConvert.DeserializeObject<List<Meta>>(json);
        }

        private void DatePickers_ValueChanged(object sender, EventArgs e)
        {
            /* DateTime fechaInicio = dateTimePicker1.Value.Date;
             DateTime fechaFin = dateTimePicker2.Value.Date;

             if (fechaFin > fechaInicio)
             {
                 lblFechasRes.Text = $"{fechaInicio.ToShortDateString()} - {fechaFin.ToShortDateString()}";
                 lblFechasRes.ForeColor = System.Drawing.Color.Green;






                 string rutaCertificacion = txtCertificacion.Lines.FirstOrDefault()!;
                 string rutaPlazos = txtPlazos.Lines.FirstOrDefault()!;
                 string rutaCalidad = txtCalidad.Lines.FirstOrDefault()!;
                 string rutaCalidadOperarios = txtCalidadOperarios.Lines.FirstOrDefault()!;
                 string rutaReclamos = txtReclamos.Lines.FirstOrDefault()!;



                 if (string.IsNullOrEmpty(rutaCertificacion)
                     || string.IsNullOrEmpty(rutaPlazos)
                    || string.IsNullOrEmpty(rutaCalidad)
                     || string.IsNullOrEmpty(rutaCalidadOperarios)
                     || string.IsNullOrEmpty(rutaReclamos)

                     )
                 {
                     btnGenerarLibros.Enabled = false;
                     btnGenerarLibros.Visible = false;
                 }
                 else
                 {
                     btnGenerarLibros.Enabled = true;
                     btnGenerarLibros.Visible = true;
                 }
             }
             else
             {
                 lblFechasRes.Text = string.Empty; // No mostrar nada si no es válido
                 btnGenerarLibros.Enabled = false;
                 btnGenerarLibros.Visible = false;
             }*/
            ValidarCampos();
        }


        // Asegurate de tener esta clase en el mismo archivo o separada
        public class Baremo
        {
            public string Name { get; set; }
            public string Description { get; set; }
            public string Tarifa { get; set; }
            public double Valor { get; set; }
            public bool IsActive { get; set; }
            public string DataEntry { get; set; }
            public string DateEdit { get; set; }
        }

        public class Meta
        {
            public string Name { get; set; }
            public string Description { get; set; }
            public double Number { get; set; }
            public double Valor { get; set; }
            public bool IsActive { get; set; }
            public string DataEntry { get; set; }
        }


        private void ValidarCampos()
        {
            DateTime fechaInicio = dateTimePicker1.Value.Date;
            DateTime fechaFin = dateTimePicker2.Value.Date;

            if (fechaFin > fechaInicio)
            {
                lblFechasRes.Text = $"{fechaInicio.ToShortDateString()} - {fechaFin.ToShortDateString()}";
                lblFechasRes.ForeColor = System.Drawing.Color.Green;






                string rutaCertificacion = txtCertificacion.Lines.FirstOrDefault()!;
                string rutaPlazos = txtPlazos.Lines.FirstOrDefault()!;
                string rutaCalidad = txtCalidad.Lines.FirstOrDefault()!;
                string rutaCalidadOperarios = txtCalidadOperarios.Lines.FirstOrDefault()!;
                string rutaReclamos = txtReclamos.Lines.FirstOrDefault()!;



                if (string.IsNullOrEmpty(rutaCertificacion)
                    || string.IsNullOrEmpty(rutaPlazos)
                   || string.IsNullOrEmpty(rutaCalidad)
                    || string.IsNullOrEmpty(rutaCalidadOperarios)
                    || string.IsNullOrEmpty(rutaReclamos)

                    )
                {
                    btnGenerarLibros.Enabled = false;
                    btnGenerarLibros.Visible = false;
                }
                else
                {
                    btnGenerarLibros.Enabled = true;
                    btnGenerarLibros.Visible = true;
                }
            }
            else
            {
               // lblFechasRes.Text = string.Empty; // No mostrar nada si no es válido
                lblFechasRes.Text = $"ERROR";
                lblFechasRes.ForeColor = Color.Red;
                btnGenerarLibros.Enabled = false;
                btnGenerarLibros.Visible = false;
            }
        }

        private void txtCertificacion_DragDrop(object sender, DragEventArgs e)
        {
            ViewHelper.AgregarRutaATextBox(e, txtCertificacion);
        }

        private void txtCertificacion_DragEnter(object sender, DragEventArgs e)
        {
            ViewHelper.EventoDragEnter(e);
        }

        private void txtPlazos_DragDrop(object sender, DragEventArgs e)
        {
            ViewHelper.AgregarRutaATextBox(e, txtPlazos);
        }

        private void txtPlazos_DragEnter(object sender, DragEventArgs e)
        {
            ViewHelper.EventoDragEnter(e);
        }

        private void txtCalidad_DragDrop(object sender, DragEventArgs e)
        {
            ViewHelper.AgregarRutaATextBox(e, txtCalidad);
        }

        private void txtCalidad_DragEnter(object sender, DragEventArgs e)
        {
            ViewHelper.EventoDragEnter(e);
        }

        private void txtCalidadOperarios_DragDrop(object sender, DragEventArgs e)
        {
            ViewHelper.AgregarRutaATextBox(e, txtCalidadOperarios);
        }

        private void txtCalidadOperarios_DragEnter(object sender, DragEventArgs e)
        {
            ViewHelper.EventoDragEnter(e);
        }

        private void txtReclamos_DragDrop(object sender, DragEventArgs e)
        {
            ViewHelper.AgregarRutaATextBox(e, txtReclamos);
        }

        private void txtReclamos_DragEnter(object sender, DragEventArgs e)
        {
            ViewHelper.EventoDragEnter(e);
        }

        private void txtCertificacion_TextChanged(object sender, EventArgs e)
        {
            ValidarCampos();
        }

        private void txtPlazos_TextChanged(object sender, EventArgs e)
        {
            ValidarCampos();
        }

        private void txtCalidad_TextChanged(object sender, EventArgs e)
        {
            ValidarCampos();
        }

        private void txtCalidadOperarios_TextChanged(object sender, EventArgs e)
        {
            ValidarCampos();
        }

        private void txtReclamos_TextChanged(object sender, EventArgs e)
        {
            ValidarCampos();
        }
    }
}
