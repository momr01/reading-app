using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MultasLectura.Helpers
{
    public class DBConexion
    {
        private static readonly HttpClient client = new HttpClient();

       /* private async Task ObtenerUsuariosAsync()
        {
            try
            {
                var response = await client.GetAsync("http://localhost:3000/api/usuarios");
                response.EnsureSuccessStatusCode();

                string responseBody = await response.Content.ReadAsStringAsync();
                // Deserializar a tu clase
                var usuarios = JsonConvert.DeserializeObject<List<Usuario>>(responseBody);

                // Usar los datos (por ejemplo, mostrarlos en un DataGridView)
                dataGridView1.DataSource = usuarios;
            }
            catch (HttpRequestException e)
            {
                MessageBox.Show($"Error: {e.Message}");
            }
        }*/

    }
}
