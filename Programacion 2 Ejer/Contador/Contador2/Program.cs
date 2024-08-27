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

            // Agregar los datos de los estudiantes
            foreach (var estudiante in estudiantes)
            {
                Word.Paragraph parrafo = wordDoc.Content.Paragraphs.Add();
                parrafo.Range.Text = "Estudiante: " + estudiante[0];
                parrafo.Range.Font.Size = 14;
                parrafo.Range.Font.Bold = 1;
                parrafo.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                parrafo.Range.InsertParagraphAfter();

                parrafo = wordDoc.Content.Paragraphs.Add();
                parrafo.Range.Text = "  Matemáticas: " + estudiante[1];
                parrafo.Range.InsertParagraphAfter();

                parrafo = wordDoc.Content.Paragraphs.Add();
                parrafo.Range.Text = "  Lenguaje: " + estudiante[2];
                parrafo.Range.InsertParagraphAfter();

                parrafo = wordDoc.Content.Paragraphs.Add();
                parrafo.Range.Text = "  Religión: " + estudiante[3];
                parrafo.Range.InsertParagraphAfter();

                // Añadir una línea vacía después de cada estudiante
                parrafo = wordDoc.Content.Paragraphs.Add();
                parrafo.Range.InsertParagraphAfter();
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
