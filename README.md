# ReservasXMLtoCSV
Aplicación de Windows Forms en C# diseñada para cargar, procesar y exportar archivos XML de reservas hoteleras desde Ópera Cloud a un archivo Excel (.xlsx) con formato determinado.
El programa detecta automáticamente el tipo de XML cargado y procesa los datos según su estructura:
1. Llegadas (Arrivals)
2. Salidas (Departures)
3. Huéspedes en casa (Guest In House)
   
Funcionalidades principales:
- Carga de archivos XML mediante selector de archivos.
- Detección automática del tipo de documento XML.
- Extracción y normalización de datos de huéspedes:
- Separación de nombre y apellido
- Limpieza de prefijos (Sr., Mr., Dr., etc.)
- Visualización de datos en DataGridView.
- Cálculo de sumatorios.
- Exportación a Excel (.xlsx) usando ClosedXML, incluyendo:
- Título dinámico según el tipo de listado
- Cabeceras formateadas
- Ajuste automático de columnas
- Bordes de tabla
- Pie de página con fecha y numeración
- Diseño optimizado para impresión

Tecnologías utilizadas:

C# (.NET, Windows Forms)
XML (XmlDocument)
DataTables
ClosedXML (exportación a Excel)

Uso:
1. Cargar un archivo XML de reservas. (Generalmente generado con Ópera Cloud)
2. Visualizar y validar los datos en pantalla.
3. Exportar el listado a Excel con un solo clic.
