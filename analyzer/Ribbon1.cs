using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;
using analyzer.Core;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

namespace analyzer
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private async void buttonImportCsv_Click(object sender, RibbonControlEventArgs e)
        {
            using (var dialog = new OpenFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv",
                Multiselect = true,
                Title = "楽天カード CSV を選択"
            })
            {
                if (dialog.ShowDialog() != DialogResult.OK || dialog.FileNames.Length == 0)
                {
                    return;
                }

                await ImportAsync(dialog.FileNames);
            }
        }

        private void buttonUpdateType_Click(object sender, RibbonControlEventArgs e)
        {
            var addIn = Globals.ThisAddIn;
            if (addIn?.Application == null)
            {
                MessageBox.Show("Excel アプリケーションが見つかりません。", "RelaxAnalyzer", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (addIn.Application.ActiveWorkbook == null)
            {
                MessageBox.Show("アクティブなブックがありません。", "RelaxAnalyzer", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var activeSheet = addIn.Application.ActiveSheet as Excel.Worksheet;
            if (activeSheet == null)
            {
                MessageBox.Show("アクティブなシートがありません。", "RelaxAnalyzer", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var warnings = new List<string>();
            var config = addIn.Configuration ?? RelaxAnalyzerConfig.Load();
            var typeMappings = TypeMappingProvider.LoadMappings(config, addIn.Application.ActiveWorkbook, warnings);
            if (typeMappings.Count == 0)
            {
                MessageBox.Show("type マッピングデータが見つかりません。\ntype シートまたは type.csv を確認してください。", "RelaxAnalyzer", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var resolver = new TypeResolver(typeMappings);
            try
            {
                UpdateTypeColumn(activeSheet, resolver, warnings);
            }
            catch (Exception ex)
            {
                MessageBox.Show("消費種類更新中にエラーが発生しました。\n" + ex.Message, "RelaxAnalyzer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            NotifyCompletion(warnings);
        }

        private void buttonAmazonOrderSummary_Click(object sender, RibbonControlEventArgs e)
        {
            using (var dialog = new OpenFileDialog
            {
                Filter = "Amazon OrderHistory CSV (Retail.OrderHistory*.csv)|Retail.OrderHistory*.csv|CSV files (*.csv)|*.csv",
                Multiselect = false,
                Title = "Amazon: Retail.OrderHistory*.csv を選択"
            })
            {
                if (dialog.ShowDialog() != DialogResult.OK || string.IsNullOrWhiteSpace(dialog.FileName))
                {
                    return;
                }

                string outputPath;
                using (var save = new SaveFileDialog
                {
                    Filter = "CSV files (*.csv)|*.csv",
                    FileName = "amazon_order_summary.csv",
                    Title = "出力先 amazon_order_summary.csv"
                })
                {
                    if (save.ShowDialog() != DialogResult.OK || string.IsNullOrWhiteSpace(save.FileName))
                    {
                        return;
                    }

                    outputPath = save.FileName;
                }

                try
                {
                    var service = new AmazonOrderSummaryService();
                    var result = service.Generate(dialog.FileName, outputPath);

                    var addIn = Globals.ThisAddIn;
                    if (addIn?.Application != null && addIn.Application.ActiveWorkbook != null)
                    {
                        PasteCsvToAmazonSheet(addIn.Application.ActiveWorkbook, outputPath);
                    }

                    MessageBox.Show(string.Join(Environment.NewLine, result.Logs), "Amazon CSV サマリ作成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Amazon CSV サマリ作成中にエラーが発生しました。\n" + ex.Message, "RelaxAnalyzer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private static void PasteCsvToAmazonSheet(Excel.Workbook workbook, string csvPath)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (string.IsNullOrWhiteSpace(csvPath)) throw new ArgumentException(nameof(csvPath));

            var app = workbook.Application;
            var prevCalc = app.Calculation;
            var prevScreenUpdating = app.ScreenUpdating;
            var prevEnableEvents = app.EnableEvents;

            try
            {
                app.ScreenUpdating = false;
                app.EnableEvents = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;

                Excel.Worksheet sheet = null;
                foreach (Excel.Worksheet ws in workbook.Worksheets)
                {
                    if (string.Equals(ws.Name, "amazon", StringComparison.OrdinalIgnoreCase))
                    {
                        sheet = ws;
                        break;
                    }
                }

                if (sheet == null)
                {
                    sheet = (Excel.Worksheet)workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
                    sheet.Name = "amazon";
                }

                // Clear sheet then paste CSV
                sheet.Cells.ClearContents();

                var data = AmazonOrderSummaryService.ReadOutputCsvForPaste(csvPath);
                if (data == null || data.GetLength(0) ==0 || data.GetLength(1) ==0)
                {
                    return;
                }

                var rows = data.GetLength(0);
                var cols = data.GetLength(1);
                var start = (Excel.Range)sheet.Cells[1,1];
                var end = (Excel.Range)sheet.Cells[rows, cols];
                var range = sheet.Range[start, end];
                range.Value2 = data;

                // Apply simple formatting
                sheet.Columns.AutoFit();
                sheet.Range[sheet.Cells[1,1], sheet.Cells[1, cols]].Font.Bold = true;
                sheet.Activate();
            }
            finally
            {
                app.Calculation = prevCalc;
                app.EnableEvents = prevEnableEvents;
                app.ScreenUpdating = prevScreenUpdating;
            }
        }

        private void UpdateTypeColumn(Excel.Worksheet worksheet, TypeResolver resolver, IList<string> warnings)
        {
            var usedRange = worksheet.UsedRange;
            if (usedRange == null)
            {
                warnings.Add("アクティブシートにデータがありません。");
                return;
            }

            var lastRow = usedRange.Rows.Count;
            const int startRow = 4;
            if (lastRow < startRow)
            {
                warnings.Add("アクティブシートのデータが不足しています。");
                return;
            }

            var totalRows = lastRow - startRow + 1;
            var bRange = (Excel.Range)worksheet.Range[worksheet.Cells[startRow, 2], worksheet.Cells[lastRow, 2]];
            var bValuesObj = bRange.Value2;
            if (bValuesObj == null)
            {
                warnings.Add("B列データが取得できませんでした。");
                return;
            }

            var bValues = bValuesObj as object[,];
            if (bValues == null)
            {
                warnings.Add("B列データの形式が不正です。");
                return;
            }

            var kRange = (Excel.Range)worksheet.Range[worksheet.Cells[startRow, 11], worksheet.Cells[lastRow, 11]];
            var kValuesObj = kRange.Value2;
            var kValues = kValuesObj as object[,];
            if (kValues == null)
            {
                kValues = new object[totalRows, 1];
            }

            var updatedCount = 0;

            var app = worksheet.Application;
            var prevCalc = app.Calculation;
            var prevScreenUpdating = app.ScreenUpdating;
            var prevEnableEvents = app.EnableEvents;
            try
            {
                if (totalRows > 50)
                {
                    app.ScreenUpdating = false;
                    app.EnableEvents = false;
                    app.Calculation = Excel.XlCalculation.xlCalculationManual;
                }

                for (var i = 1; i <= totalRows; i++)
                {
                    var raw = bValues[i, 1];
                    var storeName = raw == null ? string.Empty : Convert.ToString(raw).Trim();
                    if (string.IsNullOrEmpty(storeName))
                    {
                        continue;
                    }

                    var resolvedType = resolver.Resolve(storeName);
                    if (!string.IsNullOrEmpty(resolvedType))
                    {
                        kValues[i, 1] = resolvedType;
                        updatedCount++;
                    }
                }

                kRange.Value2 = kValues;
            }
            finally
            {
                if (totalRows > 50)
                {
                    app.Calculation = prevCalc;
                    app.EnableEvents = prevEnableEvents;
                    app.ScreenUpdating = prevScreenUpdating;
                }
            }

            if (updatedCount == 0)
            {
                warnings.Add("更新可能な消費種類が見つかりませんでした。");
            }
        }

        private async Task ImportAsync(string[] fileNames)
        {
            var addIn = Globals.ThisAddIn;
            if (addIn?.Application == null)
            {
                MessageBox.Show("Excel アプリケーションが見つかりません。", "RelaxAnalyzer", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (addIn.Application.ActiveWorkbook == null)
            {
                MessageBox.Show("アクティブなブックがありません。先にブックを開いてください。", "RelaxAnalyzer", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var warnings = new List<string>();
            var config = addIn.Configuration ?? RelaxAnalyzerConfig.Load();
            var typeMappings = TypeMappingProvider.LoadMappings(config, addIn.Application.ActiveWorkbook, warnings);
            var resolver = new TypeResolver(typeMappings);
            var service = new CsvImportService(resolver, warnings);

            IReadOnlyList<SheetImportBatch> batches;
            try
            {
                batches = await Task.Run(() => service.LoadFiles(fileNames));
            }
            catch (Exception ex)
            {
                MessageBox.Show("CSV 読込中にエラーが発生しました。\n" + ex.Message, "RelaxAnalyzer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                var writer = new SheetWriter(addIn.Application, addIn.SheetStates, warnings);
                writer.ApplyBatches(batches);
            }
            catch (Exception ex)
            {
                MessageBox.Show("シート書き込み中にエラーが発生しました。\n" + ex.Message, "RelaxAnalyzer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            NotifyCompletion(warnings);
        }

        private static void NotifyCompletion(IList<string> warnings)
        {
            if (warnings == null || warnings.Count == 0)
            {
                MessageBox.Show("処理が完了しました。", "RelaxAnalyzer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var message = "処理が完了しました。\n\n" + string.Join(Environment.NewLine, warnings);
            MessageBox.Show(message, "RelaxAnalyzer", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void buttonAmazonCheck_Click(object sender, RibbonControlEventArgs e)
        {
            var addIn = Globals.ThisAddIn;
            if (addIn?.Application == null)
            {
                MessageBox.Show("Excel アプリケーションが見つかりません。", "RelaxAnalyzer", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (addIn.Application.ActiveWorkbook == null)
            {
                MessageBox.Show("アクティブなブックがありません。", "RelaxAnalyzer", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var activeSheet = addIn.Application.ActiveSheet as Excel.Worksheet;
            if (activeSheet == null)
            {
                MessageBox.Show("アクティブなシートがありません。", "RelaxAnalyzer", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var sheetName = activeSheet.Name;

            // アクティブシートが1〜12の月シートかどうか確認
            var isMonthSheet = int.TryParse(sheetName, out var month) && month >= 1 && month <= 12;

            if (!isMonthSheet)
            {
                // アクティブシートが月シートでない場合、全シート処理を確認
                var result = MessageBox.Show(
                    "アクティブシートは月シート（1〜12）ではありません。\n全ての月シートに対してAmazon照合を実行しますか？",
                    "RelaxAnalyzer - Amazon Check",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result != DialogResult.Yes)
                {
                    return;
                }
            }

            var warnings = new List<string>();
            var service = new AmazonCheckService(warnings);

            try
            {
                if (isMonthSheet)
                {
                    // アクティブシートのみ処理
                    service.CheckAmazonForSheet(addIn.Application.ActiveWorkbook, sheetName);
                }
                else
                {
                    // 全シート処理
                    service.CheckAmazonForAllSheets(addIn.Application.ActiveWorkbook);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Amazon照合中にエラーが発生しました。\n" + ex.Message, "RelaxAnalyzer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            NotifyCompletion(warnings);
        }
    }
}
