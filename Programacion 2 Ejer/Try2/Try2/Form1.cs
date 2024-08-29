using System;
using System.Diagnostics;
using System.IO;
using OfficeOpenXml;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace Try2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Configurar el filtro del diálogo para que solo muestre archivos de Excel
            openFileDialog1.Filter = "Archivos de Excel (*.xlsx)|*.xlsx|Todos los archivos (*.*)|*.*";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;

            // Mostrar el cuadro de diálogo y verificar si el usuario seleccionó un archivo
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string rutaArchivoExcel = openFileDialog1.FileName;

                string curso = textBoxCurso.Text;

                CrearBoletin(rutaArchivoExcel, curso);
            }
            else
            {
                MessageBox.Show("No se seleccionó ningún archivo.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void CrearBoletin(string rutaArchivoExcel, string curso)
        {
            saveFileDialog1.Filter = "Documento de Word (.docx)|.docx";
            saveFileDialog1.Title = "Guardar Boletín de Notas";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string rutaArchivoWord = saveFileDialog1.FileName;

                // Inicializar EPPlus y leer el archivo Excel
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var package = new ExcelPackage(new FileInfo(rutaArchivoExcel));
                var worksheet = package.Workbook.Worksheets[0]; // Leer la primera hoja de cálculo

                int filas = worksheet.Dimension.End.Row;

                // Crear un arreglo para almacenar los datos
                string[][] estudiantes = new string[filas - 1][];

                for (int i = 2; i <= filas; i++) 
                {
                    estudiantes[i - 2] = new string[]
                    {
                        worksheet.Cells[i, 1].Text, // Nombre
                        worksheet.Cells[i, 2].Text, // Matemáticas
                        worksheet.Cells[i, 3].Text, // Lenguaje
                        worksheet.Cells[i, 4].Text, // Religión
                        worksheet.Cells[i, 5].Text  // Observaciones
                    };
                }

                // Crear una nueva aplicación de Word
                Word.Application wordApp = new Word.Application();
                Word.Document wordDoc = wordApp.Documents.Add();

                
                foreach (Word.Section section in wordDoc.Sections)
                {
                    Word.HeaderFooter header = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                    header.Range.Text = ""; 
                }

                
                foreach (Word.Section section in wordDoc.Sections)
                {
                    Word.HeaderFooter footer = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                    footer.Range.Text = "Fecha: " + DateTime.Now.ToString("dd/MM/yyyy") + " - Página ";
                    footer.PageNumbers.Add();
                    footer.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                }

                
                foreach (var estudiante in estudiantes)
                {
                    
                    Word.Range range = wordDoc.Content.Paragraphs.Add().Range;
                    Word.Table tabla = wordDoc.Tables.Add(range, 6, 2); // 6 filas, 2 columnas 
                    tabla.Borders.Enable = 1;

                    // Estilos bordes
                    tabla.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    tabla.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    tabla.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    tabla.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                    tabla.Borders[Word.WdBorderType.wdBorderHorizontal].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    tabla.Borders[Word.WdBorderType.wdBorderVertical].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                    tabla.Cell(1, 1).Range.Text = "NOMBRE DE LA INSTITUCIÓN";
                    tabla.Cell(1, 2).Range.Text = "U.E. SIMÓN BOLÍVAR";

                    tabla.Cell(2, 1).Range.Text = "CURSO";
                    tabla.Cell(2, 2).Range.Text = curso;

                    tabla.Cell(3, 1).Range.Text = "NOMBRE";
                    tabla.Cell(3, 2).Range.Text = estudiante[0]; 

                    tabla.Cell(4, 1).Range.Text = "MATEMÁTICAS";
                    tabla.Cell(4, 2).Range.Text = estudiante[1]; 

                    tabla.Cell(5, 1).Range.Text = "LENGUAJE";
                    tabla.Cell(5, 2).Range.Text = estudiante[2]; 

                    tabla.Cell(6, 1).Range.Text = "RELIGIÓN";
                    tabla.Cell(6, 2).Range.Text = estudiante[3]; 

                    // Encabezado
                    for (int i = 1; i <= 3; i++)
                    {
                        tabla.Cell(i, 1).Shading.BackgroundPatternColor = Word.WdColor.wdColorGray20;
                        tabla.Cell(i, 1).Range.Font.Color = Word.WdColor.wdColorBlack;
                        tabla.Cell(i, 2).Shading.BackgroundPatternColor = Word.WdColor.wdColorGray20;
                        tabla.Cell(i, 2).Range.Font.Color = Word.WdColor.wdColorBlack;
                    }

                    // Color materias
                    for (int i = 4; i <= tabla.Rows.Count; i++)
                    {
                        tabla.Cell(i, 1).Shading.BackgroundPatternColor = Word.WdColor.wdColorLightBlue;
                    }

                    // Resaltar notas bajas en rojo
                    int umbralNota = 60;
                    for (int i = 4; i <= 6; i++)
                    {
                        string notaTexto = tabla.Cell(i, 2).Range.Text.Trim(); // Eliminar espacios y caracteres no deseados
                        notaTexto = notaTexto.Replace("\r", "").Replace("\a", ""); // Eliminar caracteres de fin de párrafo o de celda
                        int nota;
                        if (int.TryParse(notaTexto, out nota) && nota < umbralNota)
                        {
                            tabla.Cell(i, 2).Range.Font.Color = Word.WdColor.wdColorRed;
                        }
                    }

                    
                    tabla.Rows.Add();
                    tabla.Cell(tabla.Rows.Count, 1).Range.Text = "Observaciones:";
                    tabla.Cell(tabla.Rows.Count, 2).Range.Text = estudiante[4]; 
                    tabla.Cell(tabla.Rows.Count, 1).Range.Italic = 1;
                    tabla.Cell(tabla.Rows.Count, 1).Range.Font.Size = 12;

                    
                    tabla.Rows.Add();
                    tabla.Cell(tabla.Rows.Count, 1).Range.Text = "_________________________\nFIRMA DEL PROFESOR";
                    tabla.Cell(tabla.Rows.Count, 2).Range.Text = "_________________________\nFIRMA DEL PADRE DE FAMILIA";

                    
                    tabla.Rows[tabla.Rows.Count].Height = 30;  // Ajusta la altura según sea necesario

                    tabla.Cell(tabla.Rows.Count, 1).Range.Font.Bold = 1;
                    tabla.Cell(tabla.Rows.Count, 2).Range.Font.Bold = 1;
                    tabla.Cell(tabla.Rows.Count, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    tabla.Cell(tabla.Rows.Count, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    // Aplicar formato a las celdas de la tabla
                    for (int i = 1; i <= tabla.Rows.Count; i++)
                    {
                        tabla.Cell(i, 1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        tabla.Cell(i, 1).Range.Font.Bold = 1;
                        tabla.Cell(i, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    }

                    // Añadir una línea vacía después de cada tabla
                    Word.Paragraph espacio = wordDoc.Content.Paragraphs.Add();
                    espacio.Range.InsertParagraphAfter();
                }

                
                wordDoc.SaveAs2(rutaArchivoWord);

                
                wordDoc.Close();
                wordApp.Quit();

                
                Process.Start(rutaArchivoWord);

                MessageBox.Show("El boletín de notas se ha creado y guardado exitosamente en " + rutaArchivoWord, "Boletín Creado", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("No se seleccionó ninguna ubicación para guardar el archivo.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {
            //error
        }
    }
}
