using System;
using System.Collections.Generic;
using analyzer.Core;

namespace analyzer
{
    public partial class ThisAddIn
    {
        internal RelaxAnalyzerConfig Configuration { get; private set; }

        internal IDictionary<string, SheetState> SheetStates { get; } = new Dictionary<string, SheetState>(StringComparer.OrdinalIgnoreCase);

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Configuration = RelaxAnalyzerConfig.Load();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
