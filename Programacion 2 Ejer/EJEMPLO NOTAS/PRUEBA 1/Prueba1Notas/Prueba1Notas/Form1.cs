using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace Prueba1NotasD
{
    public partial class zlForm1 : Form
    {
        private string excelFilePath;

        public Form1()
        {
            InitializeComponent();
        }

        private void ButonSeleccionarExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Archivos Excel|*.xls;*.xlsx;*.xlsm";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                excelFilePath = openFileDialog.FileName;
                MessageBox.Show("Archivo Excel seleccionado: " + excelFilePath);
            }
        }

        private void ButtonGenerarWord_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(excelFilePath))
            {
                MessageBox.Show("Por favor, seleccione un archivo de Excel primero.");
                return;
            }

            // Cargar el archivo de Excel
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            Excel._Worksheet worksheet = workbook.Sheets[1];
            Excel.Range range = worksheet.UsedRange;

            // Crear un nuevo documento de Word
            Word.Application wordApp = new Word.Application();
            Word.Document doc = wordApp.Documents.Add();

            // Variables para datos
            int numRows = range.Rows.Count;
            int numCols = range.Columns.Count;

            for (int i = 2; i <= numRows; i++) // Empieza desde la segunda fila si la primera es encabezado
            {
                // Crear nueva página para cada estudiante
                if (i > 2) doc.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);

                // Insertar cabecera con información del estudiante
                Word.Paragraph para = doc.Content.Paragraphs.Add();
                para.Range.Text = $"U.E SIMON BOLIVAR\nESTUDIANTE: {range.Cells[i, 1].Value2.ToString()}\nCURSO: {range.Cells[i, 2].Value2.ToString()} {range.Cells[i, 3].Value2.ToString()}\nGESTION: {DateTime.Now.Year}";
                para.Range.InsertParagraphAfter();

                // Crear tabla para materias y notas
                Word.Table table = doc.Tables.Add(doc.Bookmarks["\\endofdoc"].Range, numCols - 1, 4);
                table.Borders.Enable = 1;

                // Encabezado de tabla
                table.Cell(1, 1).Range.Text = "ASIGNATURA";
                table.Cell(1, 2).Range.Text = "1TRI";
                table.Cell(1, 3).Range.Text = "2TRI";
                table.Cell(1, 4).Range.Text = "3TRI";

                for (int j = 4; j <= numCols; j++)
                {
                    string subject = range.Cells[1, j].Value2.ToString();
                    table.Cell(j - 3, 1).Range.Text = subject;

                    for (int k = 2; k <= 4; k++)
                    {
                        double grade = range.Cells[i, j].Value2 != null ? range.Cells[i, j].Value2 : 0;
                        Word.Cell cell = table.Cell(j - 3, k);

                        if (k == 2) cell.Range.Text = grade.ToString(); // 1TRI
                        if (k == 3) cell.Range.Text = grade.ToString(); // 2TRI
                        if (k == 4) cell.Range.Text = grade.ToString(); // 3TRI

                        if (grade < 51)
                        {
                            cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorRed;
                        }
                    }
                }
            }

            // Mostrar el documento Word
            wordApp.Visible = true;

            // Liberar recursos
            workbook.Close(false);
            excelApp.Quit();
            wordApp.Quit();
        }
    }
}