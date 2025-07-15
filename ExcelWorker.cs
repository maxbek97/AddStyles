using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.ComponentModel;
using System.IO;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Style;
using System.Drawing;



namespace AddStyles
{
    class ExcelWorker
    {
        private string ColumnLetterToAdd { get; set; }
        private string exl_name { get; init; }
        private string dst_folder { get; set; }
        private ExcelPackage src_exl_file { get; init; }
        private ExcelPackage redacted_exl_file { get; set; }
        public ExcelWorker(string name, string dst)
        {
            exl_name = name;
            dst_folder = dst;
            src_exl_file = new ExcelPackage(new FileInfo(exl_name));
            redacted_exl_file = new ExcelPackage();
            ColumnLetterToAdd = "L";
        }

        /// <summary>
        /// Преобразует буквенное обозначение столбца Excel (например, "A", "AB") в числовой индекс (1, 28).
        /// </summary>
        /// <param name="columnLetter">Название столбца (без номера строки), например "A", "Z", "AA".</param>
        /// <returns>Номер столбца, начиная с 1.</returns>
        private int GetColumnNumber(string columnLetter)
        {
            if (string.IsNullOrWhiteSpace(columnLetter))
                throw new ArgumentException("Column letter cannot be null or empty.");

            columnLetter = columnLetter.Trim().ToUpperInvariant();

            int result = 0;

            foreach (char c in columnLetter)
            {
                if (c < 'A' || c > 'Z')
                    throw new ArgumentException($"Invalid character '{c}' in column letter.");

                result = result * 26 + (c - 'A' + 1);
            }

            return result;
        }

        public void getNewCollectionSheets()
        {
            var worksheet_list = src_exl_file.Workbook.Worksheets;
            var redacted_workseets_list = redacted_exl_file.Workbook.Worksheets;
            int counter = 1;
            foreach (var worksheet in worksheet_list)
            {
                if (worksheet.Dimension != null)
                {
                    redacted_workseets_list.Add($"Лист {counter}", worksheet);
                    counter++;
                }
            }
        }

        public void getSheet()
        {
            foreach (var sheet in redacted_exl_file.Workbook.Worksheets)
            {
                RedactSheet(sheet);
            }
        }
        public ExcelWorksheet RedactSheet(ExcelWorksheet worksheet)
        {
            var right_corner_table_adress = Get_sizes_table(worksheet);
            Console.WriteLine(right_corner_table_adress.column_address + " " + right_corner_table_adress.str_adress);
            AddColumns(worksheet, right_corner_table_adress.column_address, right_corner_table_adress.str_adress);
            ApplyConditionalFormatting(worksheet);
            return worksheet;
        }

        public void AddColumns(ExcelWorksheet worksheet, string right_corner_column_address, int right_corner_srt_adress)
        {
            
            AddColumn_header(worksheet,
                            ColumnLetterToAdd,
                            right_corner_srt_adress,
                            "Осталось дней до");
            FillColumn(worksheet,
                ColumnLetterToAdd,
                right_corner_srt_adress,
                "MID({source}{row},1,10)-TODAY()",
                "G");
            AddColumn_header(worksheet,
                ExcelCellAddress.GetColumnLetter(GetColumnNumber(ColumnLetterToAdd) + 1),
                right_corner_srt_adress,
                "Сигнализация");
            FillColumn(worksheet,
                ExcelCellAddress.GetColumnLetter(GetColumnNumber(ColumnLetterToAdd) + 1),
                right_corner_srt_adress,
                "IF(${source}{row}>10,\"Всё ок\",IF(${source}{row}>0,\"Надо рассмотреть\",\"Просрочено\"))",
                ColumnLetterToAdd);
        }

        //templates
        //"MID({source}{row},1,10)-TODAY()"
        //"IF(${source}{row}>10,\"Всё ок\",IF(${source}{row}>0,\"Надо рассмотреть\",\"Просрочено\"))"
        public void FillColumn(ExcelWorksheet worksheet, string column_adress, int begin_str_adress, string excel_function_template, string source_column_letter)
        {
            for (int row = begin_str_adress + 3; row <= worksheet.Dimension.Rows; row++)
            {
                string formula = excel_function_template.Replace("{row}", row.ToString());

                formula = formula
                    .Replace("{source}", source_column_letter);
                
                worksheet.Cells[row, GetColumnNumber(column_adress)].Formula = formula;
            }
        }

        public void ApplyConditionalFormatting(ExcelWorksheet worksheet)
        {
            var range = worksheet.Cells[$"A1:{ExcelCellAddress.GetColumnLetter(GetColumnNumber(ColumnLetterToAdd) + 1)}{worksheet.Dimension.Rows}"];

            var ruleOk = range.ConditionalFormatting.AddEqual();
            ruleOk.Formula = "\"Всё ок\"";
            ruleOk.Style.Fill.PatternType = ExcelFillStyle.Solid;
            ruleOk.Style.Fill.BackgroundColor.Color = Color.LightGreen;

            var ruleWarn = range.ConditionalFormatting.AddEqual();
            ruleWarn.Formula = "\"Надо рассмотреть\"";
            ruleWarn.Style.Fill.PatternType = ExcelFillStyle.Solid;
            ruleWarn.Style.Fill.BackgroundColor.Color = Color.Khaki;

            var ruleLate = range.ConditionalFormatting.AddEqual();
            ruleLate.Formula = "\"Просрочено\"";
            ruleLate.Style.Fill.PatternType = ExcelFillStyle.Solid;
            ruleLate.Style.Fill.BackgroundColor.Color = Color.LightCoral;
        }


        private void AddColumn_header(ExcelWorksheet worksheet, string column_address, int str_adress, string name_column)
        {
            var range = worksheet.Cells[$"{column_address}{str_adress}:{column_address}{str_adress + 2}"];
            range.Merge = true;

            // Вставляем текст
            range.Value = name_column;

            // Применяем стили
            range.Style.Font.Bold = true;
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            range.Style.Font.Size = 11;
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;
            range.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            range.Style.WrapText = true;
            range.Style.Font.Name = "Times New Roman";
        }

        private int Find_index_str(ExcelWorksheet worksheet)
        {
            int rowCount = worksheet.Dimension?.Rows ?? 0;
            for (int i = 1; i < rowCount; i++)
            {

                var value = worksheet.Cells[$"A{i}"].Text;
                if (value.IndexOf("№") != -1)
                {
                    return i;
                }
            }
            return 1;
        }

        private int Find_index_column(ExcelWorksheet worksheet)
        {
            int columnCount = worksheet.Dimension?.Columns ?? 0;
            return columnCount;
        }

        public (string column_address, int str_adress) Get_sizes_table(ExcelWorksheet worksheet)
        {
            var rightcorner_str = Find_index_str(worksheet);
            var rightcorner_column = Find_index_column(worksheet);
            return (ExcelCellAddress.GetColumnLetter(rightcorner_column), rightcorner_str);
        }

        public void main_process()
        {
            getNewCollectionSheets();
            getSheet();
            try {
                var outputPath = Path.Combine(dst_folder, Path.GetFileNameWithoutExtension(exl_name) + "_redacted.xlsx");
                redacted_exl_file.SaveAs(new FileInfo(outputPath));
            }
            catch (Exception e)
            {
                Console.WriteLine("Не удалось перезаписать, закройте программу");
            }

        }
    }
}
