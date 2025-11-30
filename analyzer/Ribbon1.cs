using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;
using analyzer.Core;
using Microsoft.Office.Tools.Ribbon;

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

            if (addIn.Application.ActiveSheet == null)
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
                UpdateTypeColumn(addIn.Application.ActiveSheet, resolver, warnings);
            }
            catch (Exception ex)
            {
                MessageBox.Show("消費種類更新中にエラーが発生しました。\n" + ex.Message, "RelaxAnalyzer", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            NotifyCompletion(warnings);
        }

        private void UpdateTypeColumn(dynamic worksheet, TypeResolver resolver, IList<string> warnings)
        {
            var usedRange = worksheet.UsedRange;
            if (usedRange == null)
            {
                warnings.Add("アクティブシートにデータがありません。");
                return;
            }

            var lastRow = usedRange.Rows.Count;
            if (lastRow < 4)
            {
                warnings.Add("アクティブシートのデータが不足しています。");
                return;
            }

            var updatedCount = 0;
            for (var row = 4; row <= lastRow; row++)
            {
                var storeNameCell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[row, 2];
                var typeCell = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[row, 11];
                var storeName = Convert.ToString(storeNameCell.Value2)?.Trim();

                if (string.IsNullOrEmpty(storeName))
                {
                    continue;
                }

                var resolvedType = resolver.Resolve(storeName);
                if (!string.IsNullOrEmpty(resolvedType))
                {
                    typeCell.Value2 = resolvedType;
                    updatedCount++;
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
    }
}
