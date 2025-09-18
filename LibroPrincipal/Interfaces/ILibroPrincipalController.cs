using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultasLectura.LibroPrincipal.Interfaces
{
    internal interface ILibroPrincipalController
    {
        void GenerarLibros(
         string rutaCertificacion,
         string rutaResumen,
        
         string rutaGuardar,
          string rutaPlazos,
            string rutaCalOperario,
            string rutaReclamos, string rutaCalidad

           //string rutaCalDetalles, string rutaCalXOper, string rutaReclDetalles, double importeCertificacion, string rutaGuardar
           );
    }
}
