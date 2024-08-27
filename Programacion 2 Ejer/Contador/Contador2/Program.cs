using System;
using Word = Microsoft.Office.Interop.Word;

namespace CrearBoletinNotas
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            // Definir el arreglo de estudiantes y notas
            string[][] estudiantes = new string[3][];
            estudiantes[0] = new string[] { "Juan", "85", "78", "92" }; // Nombre, Matemáticas, Lenguaje, Religión
            estudiantes[1] = new string[] { "María", "90", "88", "95" };
            estudiantes[2] = new string[] { "Pedro", "70", "75", "80" };

            // Especificar la ruta donde se guardará el archivo automáticamente
            string rutaArchivo = @"F:\Incos\2024\Programcion2\Programcion2\docs\BoletinNotas.docx"; // Cambia 'TuUsuario' por tu nombre de usuario

            // Crear una nueva aplicación de Word
            Word.Application wordApp = new Word.Application();
            Word.Document wordDoc = wordApp.Documents.Add();

            // Establecer el título del documento
            Word.Paragraph titulo = wordDoc.Content.Paragraphs.Add();
            titulo.Range.Text = "Boletín de Notas";
            titulo.Range.Font.Size = 24;
            titulo.Range.Font.Bold = 1;
            titulo.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            titulo.Range.InsertParagraphAfter();

            // Añadir una línea vacía
            Word.Paragraph lineaVacia = wordDoc.Content.Paragraphs.Add();
            lineaVacia.Range.InsertParagraphAfter();

            // Crear la tabla en el rango especificado
            Word.Range range = wordDoc.Content.Paragraphs.Add().Range;
            Word.Table tabla = wordDoc.Tables.Add(range, estudiantes.Length + 2, 4);
            tabla.Borders.Enable = 1;

            // Ajustar el ancho de las columnas
            tabla.Columns[1].Width = 100; // Columna NOMBRE
            tabla.Columns[2].Width = 100; // Columna MATEMATICAS
            tabla.Columns[3].Width = 100; // Columna LENGUAJE
            tabla.Columns[4].Width = 100; // Columna RELIGION

            // Combinar celdas para el encabezado
            tabla.Cell(1, 1).Merge(tabla.Cell(2, 1)); // Combina "NOMBRE"
            tabla.Cell(1, 2).Merge(tabla.Cell(1, 4)); // Combina "MATERIA/NOTA"

            // Rellenar el encabezado
            tabla.Cell(1, 1).Range.Text = "NOMBRE";
            tabla.Cell(1, 2).Range.Text = "MATERIA/NOTA";
            tabla.Cell(2, 2).Range.Text = "MATEMATICAS";
            tabla.Cell(2, 3).Range.Text = "LENGUAJE";
            tabla.Cell(2, 4).Range.Text = "RELIGION";

            // Aplicar formato a las celdas del encabezado
            tabla.Cell(1, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            tabla.Cell(1, 1).Range.Font.Bold = 1;

            tabla.Cell(1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            tabla.Cell(1, 2).Range.Font.Bold = 1;

            tabla.Cell(2, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            tabla.Cell(2, 2).Range.Font.Bold = 1;

            tabla.Cell(2, 3).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            tabla.Cell(2, 3).Range.Font.Bold = 1;

            tabla.Cell(2, 4).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            tabla.Cell(2, 4).Range.Font.Bold = 1;

            // Rellenar la tabla con los datos de los estudiantes
            for (int i = 0; i < estudiantes.Length; i++)
            {
                tabla.Cell(i + 3, 1).Range.Text = estudiantes[i][0]; // Nombre
                tabla.Cell(i + 3, 2).Range.Text = estudiantes[i][1]; // Matemáticas
                tabla.Cell(i + 3, 3).Range.Text = estudiantes[i][2]; // Lenguaje
                tabla.Cell(i + 3, 4).Range.Text = estudiantes[i][3]; // Religión
            }

            // Guardar el documento automáticamente en la ruta especificada
            wordDoc.SaveAs2(rutaArchivo);
            wordDoc.Close();
            wordApp.Quit();

            Console.WriteLine("El boletín de notas se ha creado y guardado exitosamente en " + rutaArchivo);
            Console.WriteLine("Presiona cualquier tecla para salir...");
            Console.ReadKey();
        }
    }
}
