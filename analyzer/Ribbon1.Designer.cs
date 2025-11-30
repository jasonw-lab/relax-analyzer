namespace analyzer
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region コンポーネント デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.tabRelaxAnalyzer = this.Factory.CreateRibbonTab();
            this.groupAnalyze = this.Factory.CreateRibbonGroup();
            this.buttonImportCsv = this.Factory.CreateRibbonButton();
            this.buttonUpdateType = this.Factory.CreateRibbonButton();
            this.tabRelaxAnalyzer.SuspendLayout();
            this.groupAnalyze.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabRelaxAnalyzer
            // 
            this.tabRelaxAnalyzer.Groups.Add(this.groupAnalyze);
            this.tabRelaxAnalyzer.Label = "RelaxAnalyzer";
            this.tabRelaxAnalyzer.Name = "tabRelaxAnalyzer";
            // 
            // groupAnalyze
            // 
            this.groupAnalyze.Items.Add(this.buttonImportCsv);
            this.groupAnalyze.Items.Add(this.buttonUpdateType);
            this.groupAnalyze.Label = "Analyze";
            this.groupAnalyze.Name = "groupAnalyze";
            // 
            // buttonImportCsv
            // 
            this.buttonImportCsv.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonImportCsv.Label = "CSV取込";
            this.buttonImportCsv.Name = "buttonImportCsv";
            this.buttonImportCsv.OfficeImageId = "ImportTextFile";
            this.buttonImportCsv.ScreenTip = "CSV取込";
            this.buttonImportCsv.ShowImage = true;
            this.buttonImportCsv.SuperTip = "カード明細 CSV を取り込んで月別シートへ集約します。";
            this.buttonImportCsv.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonImportCsv_Click);
            // 
            // buttonUpdateType
            // 
            this.buttonUpdateType.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonUpdateType.Label = "消費種類";
            this.buttonUpdateType.Name = "buttonUpdateType";
            this.buttonUpdateType.OfficeImageId = "TagMarkComplete";
            this.buttonUpdateType.ScreenTip = "消費種類更新";
            this.buttonUpdateType.ShowImage = true;
            this.buttonUpdateType.SuperTip = "アクティブシートのK列を type シートのキーワードで更新します。";
            this.buttonUpdateType.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonUpdateType_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabRelaxAnalyzer);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tabRelaxAnalyzer.ResumeLayout(false);
            this.tabRelaxAnalyzer.PerformLayout();
            this.groupAnalyze.ResumeLayout(false);
            this.groupAnalyze.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabRelaxAnalyzer;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupAnalyze;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonImportCsv;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonUpdateType;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
