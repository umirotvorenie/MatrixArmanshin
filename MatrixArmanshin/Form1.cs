using System;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Excel = Microsoft.Office.Interop.Excel;
using iText.Kernel.Pdf;
using iText.Layout.Element;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using Document = iText.Layout.Document;
using Paragraph = iText.Layout.Element.Paragraph;
using Table = iText.Layout.Element.Table;
using iText.Layout.Properties;
using iText.Kernel.Font;
using System.Globalization;

namespace MatrixArmanshin
{
    public partial class Form1 : Form
    {
        private int matrixSize;
        private double[,] matrixA;

        public Form1()
        {
            InitializeComponent();

            MatrixSize.Items.AddRange(new object[] { 2, 3, 4, 5 });
            MatrixSize.SelectedIndex = 0;
        }

        private void MatrixSize_SelectedIndexChanged(object sender, EventArgs e)
        {
            matrixSize = (int)MatrixSize.SelectedItem;

            double[,] matrixA = new double[matrixSize, matrixSize];
            FillMatrix(matrixA, MatrixA);
            DisplayMatrix(matrixA, MatrixA);
        }

        private void FillMatrix(double[,] matrix, DataGridView dataGridView)
        {
            Array.Clear(matrix, 0, matrix.Length);
            for (int i = 0; i < dataGridView.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView.Columns.Count; j++)
                {
                    if (i < matrix.GetLength(0) && j < matrix.GetLength(1))
                    {
                        object cellValue = dataGridView.Rows[i].Cells[j].Value;

                        if (cellValue != null)
                        {
                            // Выводим значения введенные в ячейки для отладки

                            if (double.TryParse(cellValue.ToString(), out double parsedValue))
                            {
                                matrix[i, j] = parsedValue;
                            }
                            else
                            {
                                // Обработка ошибки, например, установка значения по умолчанию
                                matrix[i, j] = 0;
                                // Выводим сообщение об ошибке
                                MessageBox.Show($"Введите числа, другие символы вводить запрещено");
                                resultLabel.Text = null;
                            }
                        }
                    }
                }
            }
        }

        private void DisplayMatrix(double[,] matrix, DataGridView dataGridView)
        {
            int rows = matrix.GetLength(0);
            int cols = matrix.GetLength(1);
            dataGridView.Rows.Clear();
            dataGridView.Columns.Clear();
            for (int j = 0; j < cols; j++)
            {
                dataGridView.Columns.Add("", "");
            }
            for (int i = 0; i < rows; i++)
            {
                dataGridView.Rows.Add();
                for (int j = 0; j < cols; j++)
                {
                    dataGridView.Rows[i].Cells[j].Value = matrix[i, j];
                }
            }
        }

        private void ResultButton_Click(object sender, EventArgs e)
        {
            try
            {
                matrixSize = (int)MatrixSize.SelectedItem;

                double[,] matrixA = new double[matrixSize, matrixSize];
                FillMatrix(matrixA, MatrixA);
                DisplayMatrix(matrixA, MatrixA);
                if (!IsValidMatrix(matrixA))
                {
                    MessageBox.Show("Большое число! Введите число не более 8 знаков");
                    Clear(MatrixA);
                    resultLabel.Text = null;
                    return;
                }

                double determinant = CalculateDeterminant(matrixA);
                if (determinant == 0)
                {
                    resultLabel.Text = ($"Определитель матрицы: 0");
                }
                else
                {
                    resultLabel.Text = ($"Определитель матрицы: {determinant.ToString("N0")}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private bool IsValidMatrix(double[,] matrix)
        {
            if (matrix == null)
            {
                return false;
            }
            foreach (var value in matrix)
            {
                string stringValue = value.ToString("G");
                if (stringValue.Length > 8)
                {
                    return false;
                }
            }
            return true;
        }

        private double CalculateDeterminant(double[,] matrix)
        {
            int size = matrix.GetLength(0);

            // Создание копии матрицы
            double[,] matrixCopy = (double[,])matrix.Clone();

            // Приведение матрицы к треугольному виду
            double det = 1;
            for (int col = 0; col < size - 1; col++)
            {
                for (int row = col + 1; row < size; row++)
                {
                    double factor = matrixCopy[row, col] / matrixCopy[col, col];
                    for (int i = col; i < size; i++)
                    {
                        matrixCopy[row, i] -= factor * matrixCopy[col, i];
                    }
                }
            }

            // Умножение элементов главной диагонали
            for (int i = 0; i < size; i++)
            {
                det *= matrixCopy[i, i];
            }

            return det;
        }

        private static void Clear(DataGridView dataGridView)
        {
            int rows = dataGridView.RowCount;
            int columns = dataGridView.ColumnCount;
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    dataGridView.Rows[i].Cells[j].Value = 0;
                }
            }
        }

        private void WordButton_Click(object sender, EventArgs e)
        {
            try
            {
                int matrixSize = (int)MatrixSize.SelectedItem;
                double[,] matrixA = new double[matrixSize, matrixSize];
                FillMatrix(matrixA, MatrixA);

                using (WordprocessingDocument wordDocument = WordprocessingDocument.Create("MatrixDocument.docx", WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                    mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
                    Body body = mainPart.Document.AppendChild(new Body());

                    // Матрица A
                    ExportMatrixToWord(matrixA, wordDocument, "Матрица A", matrixSize);

                    // Определитель матрицы
                    double determinant = CalculateDeterminant(matrixA);
                    string determinantResult = $"Определитель матрицы: {determinant.ToString("N0")}";
                    DocumentFormat.OpenXml.Wordprocessing.Paragraph determinantParagraph = body.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
                    Run determinantRun = determinantParagraph.AppendChild(new Run());
                    determinantRun.AppendChild(new Text(determinantResult));
                }

                MessageBox.Show("Документ Word успешно создан и сохранен.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка при создании документа Word: {ex.Message}", "Ошибка");
            }
        }

        private string ExportMatrixToWord(double[,] matrix, WordprocessingDocument wordDocument, string matrixName, int matrixSize)
        {
            StringBuilder matrixString = new StringBuilder();

            if (wordDocument != null && matrix != null)
            {
                Body body = wordDocument.MainDocumentPart.Document.Body;

                // Заголовок матрицы
                DocumentFormat.OpenXml.Wordprocessing.Paragraph matrixTitleParagraph = body.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
                Run matrixTitleRun = matrixTitleParagraph.AppendChild(new Run());
                matrixTitleRun.AppendChild(new Text(matrixName));

                // Создаем таблицу для матрицы
                DocumentFormat.OpenXml.Wordprocessing.Table table = body.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Table());

                TableProperties tableProperties = new TableProperties(
                        new TableBorders(
                        new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 },
                        new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 },
                        new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 },
                        new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 }
                    )
                );

                table.AppendChild(tableProperties);

                // Создаем строки и столбцы таблицы
                for (int i = 0; i < matrixSize; i++)
                {
                    TableRow row = table.AppendChild(new TableRow());

                    for (int j = 0; j < matrixSize; j++)
                    {
                        TableCell cell = row.AppendChild(new TableCell());

                        TableCellProperties cellProperties = new TableCellProperties(
                            new TableCellBorders(
                                new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 },
                                new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 },
                                new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 },
                                new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 5 }
                            )
                        );

                        cell.AppendChild(cellProperties);

                        DocumentFormat.OpenXml.Wordprocessing.Paragraph cellParagraph = cell.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
                        Run cellRun = cellParagraph.AppendChild(new Run());
                        RunProperties cellRunProperties = cellRun.AppendChild(new RunProperties());
                        cellRunProperties.AppendChild(new RunFonts() { Ascii = "Calibri" });
                        cellRunProperties.AppendChild(new FontSize() { Val = "28" });

                        if (matrix[i, j] != null)
                        {
                            cellRun.AppendChild(new Text($"{matrix[i, j]}"));
                        }
                        else
                        {
                            cellRun.AppendChild(new Text(""));
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Объект wordDocument или матрица не были корректно инициализированы.", "Ошибка");
            }

            return matrixString.ToString();
        }

        private void ExportMatrixToExcel(double[,] matrix, Excel.Worksheet worksheet, string matrixName, int startRow)
        {
            int rows = matrix.GetLength(0);
            int cols = matrix.GetLength(1);

            // Заголовок матрицы
            worksheet.Cells[startRow, 1] = matrixName;

            // Заполнение матрицы
            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    worksheet.Cells[startRow + i, j + 2].Value = matrix[i, j];
                    worksheet.Cells[startRow + i, j + 2].Font.Name = "Calibri";
                    worksheet.Cells[startRow + i, j + 2].Font.Size = 14;
                }
            }

            // Оформление таблицы
            Excel.Range matrixRange = worksheet.Range[worksheet.Cells[startRow, 2], worksheet.Cells[startRow + rows - 1, cols + 1]];
            matrixRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            Marshal.ReleaseComObject(matrixRange);
        }

        private void ExcelButton_Click(object sender, EventArgs e)
        {
            try
            {
                int matrixSize = (int)MatrixSize.SelectedItem;
                double[,] matrixA = new double[matrixSize, matrixSize];
                FillMatrix(matrixA, MatrixA);
                MatrixA.Refresh();
                DisplayMatrix(matrixA, MatrixA);

                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = true;
                Excel.Workbook workbook = excelApp.Workbooks.Add();
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];
                worksheet.Name = "Результат определителя матрицы";

                // Матрица A
                ExportMatrixToExcel(matrixA, worksheet, "Матрица A", 1);

                // Определитель матрицы
                double determinant = CalculateDeterminant(matrixA);
                int startRowSolution = 1 + matrixA.GetLength(0) + 1;

                // Заголовок "Определитель"
                worksheet.Cells[startRowSolution, 1].Value2 = "Определитель";

                // Значение определителя
                worksheet.Cells[startRowSolution, 2].Value2 = determinant.ToString("N0");

                // Дополнительная информация
                // ...

                // Автоподгон ширины столбцов
                int startColumn = 1;
                int endColumn = matrixA.GetLength(1) * 3;
                Excel.Range range = worksheet.Range[worksheet.Cells[1 + matrixA.GetLength(0) + 1, startColumn], worksheet.Cells[1 + matrixA.GetLength(0) + 1, endColumn]];
                range.EntireColumn.AutoFit();
                Marshal.ReleaseComObject(range);

                worksheet.Columns.AutoFit();
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при экспорте в Excel: " + ex.Message);
            }
        }

        private void ExportMatrixToPDF(Document document, double[,] matrix)
        {
            var table = new Table(UnitValue.CreatePercentArray(matrix.GetLength(1))).UseAllAvailableWidth();

            // Используйте шрифт Helvetica
            var font = PdfFontFactory.CreateFont("Helvetica");

            for (int i = 0; i < matrix.GetLength(0); i++)
            {
                for (int j = 0; j < matrix.GetLength(1); j++)
                {
                    var cellText = matrix[i, j].ToString("0.######", CultureInfo.InvariantCulture);
                    var div = new iText.Layout.Element.Div().Add(new Paragraph(cellText));
                    var cell = new Cell().Add(div);
                    cell.SetBorderBottom(new iText.Layout.Borders.SolidBorder(1));
                    cell.SetBorderTop(new iText.Layout.Borders.SolidBorder(1));
                    cell.SetBorderLeft(new iText.Layout.Borders.SolidBorder(1));
                    cell.SetBorderRight(new iText.Layout.Borders.SolidBorder(1));

                    table.AddCell(cell);
                }
            }

            document.Add(table);
        }

        private void PDFButton_Click(object sender, EventArgs e)
        {
            try
            {
                int matrixSize = (int)MatrixSize.SelectedItem;
                double[,] matrixA = new double[matrixSize, matrixSize];
                FillMatrix(matrixA, MatrixA);
                MatrixA.Refresh();
                DisplayMatrix(matrixA, MatrixA);
                string pdfFilePath = "output.pdf";

                // Рассчитываем определитель матрицы A
                double determinant = CalculateDeterminant(matrixA);
                using (var pdfWriter = new PdfWriter(pdfFilePath))
                using (var pdfDocument = new PdfDocument(pdfWriter))
                {
                    var document = new Document(pdfDocument);
                    document.Add(new Paragraph("Matrix A:"));
                    ExportMatrixToPDF(document, matrixA);
                    document.Add(new Paragraph("Determinant:"));
                    document.Add(new Paragraph(determinant.ToString("N0")));
                }

                MessageBox.Show("Экспорт в PDF успешно завершен.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при экспорте в PDF: " + ex.Message);
            }
        }

        private void ClearButton_Click(object sender, EventArgs e)
        {
            Clear(MatrixA);
            resultLabel.Text = "";
        }

        private void ExitButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}