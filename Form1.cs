using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml;

namespace ReservaXmlToCsv
{
    public partial class Form1 : Form
    {
        DataTable tablaReservas = new DataTable();
        private string tituloExcel = "Listado de Reservas";

        DataRow filaSumatorio;

        public Form1()
        {
            InitializeComponent();
        }

        private void btnCargarXml_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Filter = "Archivos XML (*.xml)|*.xml",
                Title = "Seleccionar archivo XML"
            };

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                CargarXml(ofd.FileName);
            }
        }

        private void CargarXml(string ruta)
        {
            try
            {
                tablaReservas.Clear();
                tablaReservas.Columns.Clear();

                XmlDocument doc = new XmlDocument();
                doc.Load(ruta);

                if (doc.SelectNodes("//G_RESERVATION").Count > 0)
                {
                    ProcesarArrivals(doc);
                }
                else if (doc.SelectNodes("//G_DEPARTURE").Count > 0)
                {
                    ProcesarDepartures(doc);
                }
                else if (doc.SelectNodes("//G_ROOM").Count > 0)
                {
                    ProcesarGuestInHouse(doc);
                }
                else
                {
                    MessageBox.Show("Tipo de XML no reconocido.");
                    return;
                }

                dataGridView1.DataSource = tablaReservas;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al cargar XML: " + ex.Message);
            }
        }

        private void ProcesarArrivals(XmlDocument doc)
        {
            tituloExcel = "LLEGADAS | ARRIVALS";
            string[] columnas = new[]
            {
                "ROOM_NO", "FULL_NAME", "NAME", "SURNAME", "ARRIVAL_TIME", "DEPARTURE", "ADULTS", "CHILDREN", "VIP"
            };

            tablaReservas.Rows.Clear();
            tablaReservas.Columns.Clear();

            foreach (var col in columnas)
                tablaReservas.Columns.Add(col);

            var reservas = doc.SelectNodes("//G_RESERVATION");
            foreach (XmlNode reserva in reservas)
            {
                DataRow fila = tablaReservas.NewRow();

                string fullName = reserva.SelectSingleNode("FULL_NAME")?.InnerText ?? "";
                var (nombre, apellido) = ParsearNombreApellido(fullName);

                fila["FULL_NAME"] = fullName;
                fila["NAME"] = nombre;
                fila["SURNAME"] = apellido;

                foreach (var col in columnas)
                {
                    if (col == "FULL_NAME" || col == "NAME" || col == "SURNAME") continue;
                    fila[col] = reserva.SelectSingleNode(col)?.InnerText ?? "";
                }

                tablaReservas.Rows.Add(fila);
            }

            tablaReservas.Columns.Remove("FULL_NAME");

            var traducciones = new Dictionary<string, string>()
            {
                { "ROOM_NO", "Hab" },
                { "NAME", "NOMBRE" },
                { "SURNAME", "APELLIDO" },
                { "ARRIVAL_TIME", "LLEGADA" },
                { "DEPARTURE", "SALIDA" },
                { "ADULTS", "AD" },
                { "CHILDREN", "NI" },
                { "VIP", "VIP" }
            };

            foreach (var par in traducciones)
                if (tablaReservas.Columns.Contains(par.Key))
                    tablaReservas.Columns[par.Key].ColumnName = par.Value;

            // Sumatorios
            int totalReservas = tablaReservas.Rows.Cast<DataRow>()
                .Count(r => !r.IsNull("Hab") && !string.IsNullOrWhiteSpace(r["Hab"].ToString()));
            int sumaAD = SumarColumna("AD");
            int sumaNI = SumarColumna("NI");

            filaSumatorio = tablaReservas.NewRow();
            filaSumatorio["NOMBRE"] = "TOTALES";
            filaSumatorio["APELLIDO"] = $"Hab: {totalReservas}";
            filaSumatorio["AD"] = sumaAD;
            filaSumatorio["NI"] = sumaNI;

          // Antes:  tablaReservas.Rows.Add(filaSumatorio);
        }

        private void ProcesarDepartures(XmlDocument doc)
        {
            tituloExcel = "SALIDAS | DEPARTURES";
            string[] columnas = new[]
            {
                "ROOM", "GUEST_NAME", "NAME", "SURNAME", "CHAR_ARRIVAL", "DEPARTURE_TIME", "ADULTS", "CHILDREN", "VIP"
            };

            tablaReservas.Rows.Clear();
            tablaReservas.Columns.Clear();

            foreach (var col in columnas)
                tablaReservas.Columns.Add(col);

            var rooms = doc.SelectNodes("//G_ROOM");
            foreach (XmlNode room in rooms)
            {
                DataRow fila = tablaReservas.NewRow();

                string guestName = room.SelectSingleNode("GUEST_NAME")?.InnerText ?? "";
                var (nombre, apellido) = ParsearNombreApellido(guestName);

                fila["GUEST_NAME"] = guestName;
                fila["NAME"] = nombre;
                fila["SURNAME"] = apellido;

                foreach (var col in columnas)
                {
                    if (col == "GUEST_NAME" || col == "NAME" || col == "SURNAME") continue;
                    fila[col] = room.SelectSingleNode(col)?.InnerText ?? "";
                }

                tablaReservas.Rows.Add(fila);
            }

            tablaReservas.Columns.Remove("GUEST_NAME");

            var traducciones = new Dictionary<string, string>()
            {
                { "ROOM", "Hab" },
                { "NAME", "NOMBRE" },
                { "SURNAME", "APELLIDO" },
                { "CHAR_ARRIVAL", "LLEGADA" },
                { "DEPARTURE_TIME", "SALIDA" },
                { "ADULTS", "AD" },
                { "CHILDREN", "NI" },
                { "VIP", "VIP" }
            };

            foreach (var par in traducciones)
                if (tablaReservas.Columns.Contains(par.Key))
                    tablaReservas.Columns[par.Key].ColumnName = par.Value;

            // Sumatorios
            int totalReservas = tablaReservas.Rows.Cast<DataRow>()
                .Count(r => !r.IsNull("Hab") && !string.IsNullOrWhiteSpace(r["Hab"].ToString()));
            int sumaAD = SumarColumna("AD");
            int sumaNI = SumarColumna("NI");

            filaSumatorio = tablaReservas.NewRow();
            filaSumatorio["NOMBRE"] = "TOTALES";
            filaSumatorio["APELLIDO"] = $"Hab: {totalReservas}";
            filaSumatorio["AD"] = sumaAD;
            filaSumatorio["NI"] = sumaNI;

            // tablaReservas.Rows.Add(filaSumatorio);
        }

        private void ProcesarGuestInHouse(XmlDocument doc)
        {
            tituloExcel = "HUÉSPEDES EN CASA | GUEST IN HOUSE";
            string[] columnas = new[]
            {
                "ROOM", "FULL_NAME", "NAME", "SURNAME", "ARRIVAL", "DEPARTURE", "ADULTS", "CHILDREN"
            };

            tablaReservas.Rows.Clear();
            tablaReservas.Columns.Clear();

            foreach (var col in columnas)
                tablaReservas.Columns.Add(col);

            var rooms = doc.SelectNodes("//G_ROOM");
            foreach (XmlNode room in rooms)
            {
                DataRow fila = tablaReservas.NewRow();

                string fullName = room.SelectSingleNode("FULL_NAME")?.InnerText
                                ?? room.SelectSingleNode("GUEST_NAME")?.InnerText ?? "";
                var (nombre, apellido) = ParsearNombreApellido(fullName);

                fila["FULL_NAME"] = fullName;
                fila["NAME"] = nombre;
                fila["SURNAME"] = apellido;

                foreach (var col in columnas)
                {
                    if (col == "FULL_NAME" || col == "NAME" || col == "SURNAME") continue;
                    fila[col] = room.SelectSingleNode(col)?.InnerText ?? "";
                }

                tablaReservas.Rows.Add(fila);
            }

            tablaReservas.Columns.Remove("FULL_NAME");

            var traducciones = new Dictionary<string, string>()
            {
                { "ROOM", "Hab" },
                { "NAME", "NOMBRE" },
                { "SURNAME", "APELLIDO" },
                { "ARRIVAL", "LLEGADA" },
                { "DEPARTURE", "SALIDA" },
                { "ADULTS", "AD" },
                { "CHILDREN", "NI" }
            };

            foreach (var par in traducciones)
                if (tablaReservas.Columns.Contains(par.Key))
                    tablaReservas.Columns[par.Key].ColumnName = par.Value;


            // Sumatorios
            int totalReservas = tablaReservas.Rows.Cast<DataRow>()
                .Count(r => !r.IsNull("Hab") && !string.IsNullOrWhiteSpace(r["Hab"].ToString()));
            int sumaAD = SumarColumna("AD");
            int sumaNI = SumarColumna("NI");

            filaSumatorio = tablaReservas.NewRow();
            filaSumatorio["NOMBRE"] = "TOTALES";
            filaSumatorio["APELLIDO"] = $"Hab: {totalReservas}";
            filaSumatorio["AD"] = sumaAD;
            filaSumatorio["NI"] = sumaNI;

           // tablaReservas.Rows.Add(filaSumatorio);
        }

        private (string nombre, string apellido) ParsearNombreApellido(string fullName)
        {
            if (string.IsNullOrWhiteSpace(fullName)) return ("", "");

            var partes = fullName.Split(',');
            string apellido = partes.Length > 0 ? partes[0].Trim() : "";
            string nombre = partes.Length > 1 ? partes[1].Trim() : "";
            nombre = LimpiarPrefijos(nombre);

            return (nombre, apellido);
        }

        private string LimpiarPrefijos(string nombre)
        {
            if (string.IsNullOrWhiteSpace(nombre)) return "";

            var prefijos = new[]
            {
                "Don", "Doña", "Dr.", "Dra.", "Sr.", "Sra.", "Srta.",
                "Mr.", "Mrs.", "Ms.", "Miss", "Herr", "Frau",
                "Jr.", "Sr."
            };

            var palabras = nombre.Split(' ', StringSplitOptions.RemoveEmptyEntries).ToList();
            palabras = palabras.Where(p => !prefijos.Contains(p, StringComparer.OrdinalIgnoreCase)).ToList();

            return string.Join(" ", palabras);
        }

        private int SumarColumna(string colName)
        {
            return tablaReservas.Rows.Cast<DataRow>()
                .Where(r => tablaReservas.Columns.Contains(colName) && !r.IsNull(colName) && int.TryParse(r[colName].ToString(), out _))
                .Select(r => int.Parse(r[colName].ToString()))
                .Sum();
        }

private void btnExportarExcel_Click(object sender, EventArgs e)
{
    if (tablaReservas == null || tablaReservas.Rows.Count == 0)
    {
        MessageBox.Show("No hay datos para exportar.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        return;
    }

    if (filaSumatorio != null)
    {
        tablaReservas.Rows.Add(tablaReservas.NewRow());
        tablaReservas.Rows.Add(filaSumatorio);
    }

    SaveFileDialog sfd = new SaveFileDialog
    {
        Filter = "Archivos Excel (*.xlsx)|*.xlsx",
        Title = "Guardar como Excel"
    };

    if (sfd.ShowDialog() == DialogResult.OK)
    {
        try
        {
            using (var workbook = new ClosedXML.Excel.XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Datos");

                int columnasVisibles = dataGridView1.Columns.Cast<DataGridViewColumn>().Count(c => c.Visible);

                // 1. Columna extra a la izquierda para marcar
                worksheet.Column(1).Width = 2;

                // 2. Título en fila 1
                worksheet.Range(1, 2, 1, columnasVisibles + 1).Merge();
                worksheet.Cell(1, 2).Value = tituloExcel;
                worksheet.Cell(1, 2).Style.Font.Bold = true;
                worksheet.Cell(1, 2).Style.Font.FontSize = 16;
                worksheet.Cell(1, 2).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                worksheet.Cell(1, 2).Style.Alignment.Vertical = ClosedXML.Excel.XLAlignmentVerticalValues.Center;

                // 3. Cabeceras en fila 3
                int colIdx = 2;
                foreach (DataGridViewColumn col in dataGridView1.Columns)
                {
                    if (col.Visible)
                    {
                        worksheet.Cell(3, colIdx).Value = col.HeaderText;
                        worksheet.Cell(3, colIdx).Style.Font.Bold = true;
                        worksheet.Cell(3, colIdx).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.LightGray;
                        worksheet.Cell(3, colIdx).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
                        colIdx++;
                    }
                }

                // 4. Datos desde fila 4
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    colIdx = 2;
                    foreach (DataGridViewColumn col in dataGridView1.Columns)
                    {
                        if (col.Visible)
                        {
                            worksheet.Cell(i + 4, colIdx).Value = dataGridView1.Rows[i].Cells[col.Index].Value?.ToString() ?? "";
                            colIdx++;
                        }
                    }
                }

                // 5. Ajustar columnas
                for (int c = 2; c <= columnasVisibles + 1; c++)
                {
                    worksheet.Column(c).AdjustToContents();
                }

                // 6. Altura de filas para que la columna izquierda parezca cuadrada
                int filaUltima = dataGridView1.Rows.Count + 3;
                for (int r = 4; r <= filaUltima; r++)
                {
                    worksheet.Row(r).Height = 20;
                }

                // 7. Bordes incluyendo columna 1
                var rangoTabla = worksheet.Range(3, 1, filaUltima, columnasVisibles + 1);
                rangoTabla.Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;
                rangoTabla.Style.Border.InsideBorder = ClosedXML.Excel.XLBorderStyleValues.Thin;

                // 8. Ajuste para evitar solapamiento con el pie de página al imprimir
                worksheet.PageSetup.Margins.Bottom = 1.5; // más margen inferior en pulgadas (ajusta según necesidad)
                worksheet.PageSetup.FitToPages(1, 0);     // escalar para que quepa en 1 página de ancho

                // 9. Pie de página
                worksheet.PageSetup.Footer.Left.AddText(tituloExcel);
                worksheet.PageSetup.Footer.Center.AddText(ClosedXML.Excel.XLHFPredefinedText.PageNumber);
                worksheet.PageSetup.Footer.Center.AddText(" de ");
                worksheet.PageSetup.Footer.Center.AddText(ClosedXML.Excel.XLHFPredefinedText.NumberOfPages);
                worksheet.PageSetup.Footer.Right.AddText(DateTime.Now.ToString("dd/MM/yyyy"));

                // 10. Guardar
                workbook.SaveAs(sfd.FileName);
            }

            MessageBox.Show("Exportación a Excel completada correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show("Error al exportar a Excel: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}



    }
}
