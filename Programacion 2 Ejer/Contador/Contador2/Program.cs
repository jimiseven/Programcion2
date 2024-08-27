using System;
using System.Diagnostics;
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
            estudiantes[0] = new string[] { "Juan", "85", "78", "92" }; // Nombre, Matemáticas, LenguajeDesdeGion, Religión
            estudiantes[1] = new string[] { "María", "90", "51", "95" };
            estudiantes[2] = new string[] { "Pedro", "70", "75", "80" };

            // Especificar la ruta donde se guardará el archivo automáticamente
            string rutaArchivo = @"F:\Incos\2024\Programcion2\Programcion2\docs\BoletinNotas.docx"; // Cambia 'TuUsuario' por tu nombre de usuario

            // Crear una nueva aplicación de Word
            Word.Application wordApp = new Word.Application();
            Word.Document wordDoc = wordApp.Documents.Add();

            // Añadir un encabezado
            foreach (Word.Section section in wordDoc.Sections)
            {
                Word.HeaderFooter header = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                header.Range.Text = "Nombre de la Institución - Boletín de Notas";
                header.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                header.Range.Font.Size = 14;
                header.Range.Font.Bold = 1;
            }

            // Añadir un pie de página
            foreach (Word.Section section in wordDoc.Sections)
            {
                Word.HeaderFooter footer = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                footer.Range.Text = "Fecha: " + DateTime.Now.ToString("dd/MM/yyyy") + " - Página ";
                footer.PageNumbers.Add();
                footer.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            }

            // Establecer el título del documento
            Word.Paragraph titulo = wordDoc.Content.Paragraphs.Add();
            titulo.Range.Text = "Boletín de Notas";
            titulo.Range.Font.Size = 14;
            titulo.Range.Font.Bold = 1;
            titulo.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            titulo.Range.InsertParagraphAfter();

            // Añadir una línea vacía
            Word.Paragraph lineaVacia = wordDoc.Content.Paragraphs.Add();
            lineaVacia.Range.InsertParagraphAfter();

            // Crear un cuadro para cada estudiante
            foreach (var estudiante in estudiantes)
            {
                // Crear la tabla en el rango especificado
                Word.Range range = wordDoc.Content.Paragraphs.Add().Range;
                Word.Table tabla = wordDoc.Tables.Add(range, 3, 2); // 3 filas, 2 columnas
                tabla.Borders.Enable = 1;

                // Aplicar estilos de bordes a cada borde individual
                tabla.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                tabla.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                tabla.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                tabla.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                tabla.Borders[Word.WdBorderType.wdBorderHorizontal].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                tabla.Borders[Word.WdBorderType.wdBorderVertical].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                // Opcional: Puedes cambiar el grosor de los bordes si quieres que sean más gruesos
                tabla.Borders[Word.WdBorderType.wdBorderLeft].LineWidth = Word.WdLineWidth.wdLineWidth050pt;
                tabla.Borders[Word.WdBorderType.wdBorderRight].LineWidth = Word.WdLineWidth.wdLineWidth050pt;
                tabla.Borders[Word.WdBorderType.wdBorderTop].LineWidth = Word.WdLineWidth.wdLineWidth050pt;
                tabla.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth050pt;
                tabla.Borders[Word.WdBorderType.wdBorderHorizontal].LineWidth = Word.WdLineWidth.wdLineWidth025pt;
                tabla.Borders[Word.WdBorderType.wdBorderVertical].LineWidth = Word.WdLineWidth.wdLineWidth025pt;

                // Rellenar la tabla con el nombre y las notas
                tabla.Cell(1, 1).Range.Text = "NOMBRE";
                tabla.Cell(1, 2).Range.Text = estudiante[0]; // Nombre del estudiante

                tabla.Cell(2, 1).Range.Text = "MATEMÁTICAS";
                tabla.Cell(2, 2).Range.Text = estudiante[1]; // Nota de Matemáticas

                tabla.Cell(3, 1).Range.Text = "LENGUAJE";
                tabla.Cell(3, 2).Range.Text = estudiante[2]; // Nota de Lenguaje

                tabla.Rows.Add();
                tabla.Cell(4, 1).Range.Text = "RELIGIÓN";
                tabla.Cell(4, 2).Range.Text = estudiante[3]; // Nota de Religión

                // Aplicar colores a las celdas del encabezado
                tabla.Cell(1, 1).Shading.BackgroundPatternColor = Word.WdColor.wdColorGray20;
                tabla.Cell(1, 1).Range.Font.Color = Word.WdColor.wdColorWhite;
                tabla.Cell(1, 2).Shading.BackgroundPatternColor = Word.WdColor.wdColorGray20;
                tabla.Cell(1, 2).Range.Font.Color = Word.WdColor.wdColorWhite;

                // Aplicar color a las filas de materia
                for (int i = 2; i <= tabla.Rows.Count; i++)
                {
                    tabla.Cell(i, 1).Shading.BackgroundPatternColor = Word.WdColor.wdColorLightBlue;
                }

                // Resaltar notas bajas en rojo
                int umbralNota = 60;
                for (int i = 2; i <= tabla.Rows.Count; i++)
                {
                    string notaTexto = tabla.Cell(i, 2).Range.Text.Trim(); // Eliminar espacios y caracteres no deseados
                    notaTexto = notaTexto.Replace("\r", "").Replace("\a", ""); // Eliminar caracteres de fin de párrafo o de celda
                    int nota;  // Declarar la variable 'nota' antes del uso en TryParse
                    if (int.TryParse(notaTexto, out nota) && nota < umbralNota)
                    {
                        tabla.Cell(i, 2).Range.Font.Color = Word.WdColor.wdColorRed;
                    }
                }

                // Aplicar formato a las celdas de la tabla
                for (int i = 1; i <= tabla.Rows.Count; i++)
                {
                    tabla.Cell(i, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    tabla.Cell(i, 1).Range.Font.Bold = 1;
                    tabla.Cell(i, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }

                // Añadir una fila para observaciones
                tabla.Rows.Add();
                tabla.Cell(tabla.Rows.Count, 1).Range.Text = "Observaciones:";
                tabla.Cell(tabla.Rows.Count, 1).Merge(tabla.Cell(tabla.Rows.Count, 2)); // Combinar celdas para comentario
                tabla.Cell(tabla.Rows.Count, 1).Range.Italic = 1;
                tabla.Cell(tabla.Rows.Count, 1).Range.Font.Size = 12;

                // Añadir una línea vacía después de cada tabla
                Word.Paragraph espacio = wordDoc.Content.Paragraphs.Add();
                espacio.Range.InsertParagraphAfter();
            }

            // Añadir espacio para la firma del profesor
            Word.Paragraph firma = wordDoc.Content.Paragraphs.Add();
            firma.Range.Text = "\n\nFirma del Profesor/Tutor: ____________________";
            firma.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

            // Guardar el documento automáticamente en la ruta especificada
            wordDoc.SaveAs2(rutaArchivo);

            // Cerrar el documento y la aplicación de Word
            wordDoc.Close();
            wordApp.Quit();

            // Abrir el documento de Word automáticamente
            Process.Start(rutaArchivo);

            Console.WriteLine("El boletín de notas se ha creado y guardado exitosamente en " + rutaArchivo);
            Console.WriteLine("El documento se está abriendo...");
            Console.WriteLine("Presiona cualquier tecla para salir...");
            Console.ReadKey();
        }
    }
}
