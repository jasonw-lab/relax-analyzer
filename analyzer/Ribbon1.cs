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
                MessageBox.Show("CSV 取込が完了しました。", "RelaxAnalyzer", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var message = "CSV 取込が完了しました。\n\n" + string.Join(Environment.NewLine, warnings);
            MessageBox.Show(message, "RelaxAnalyzer", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}
