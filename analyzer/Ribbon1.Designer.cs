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
            this.buttonUpdateTypeAllSheets = this.Factory.CreateRibbonButton();
            this.buttonAmazonOrderSummary = this.Factory.CreateRibbonButton();
            this.buttonAmazonCheck = this.Factory.CreateRibbonButton();
            this.buttonUpdateYearSheets = this.Factory.CreateRibbonButton();
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
            this.groupAnalyze.Items.Add(this.buttonUpdateTypeAllSheets);
            this.groupAnalyze.Items.Add(this.buttonAmazonOrderSummary);
            this.groupAnalyze.Items.Add(this.buttonAmazonCheck);
            this.groupAnalyze.Items.Add(this.buttonUpdateYearSheets);
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
            // buttonUpdateTypeAllSheets
            // 
            this.buttonUpdateTypeAllSheets.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonUpdateTypeAllSheets.Label = "all消費種類";
            this.buttonUpdateTypeAllSheets.Name = "buttonUpdateTypeAllSheets";
            this.buttonUpdateTypeAllSheets.OfficeImageId = "RecordsMenu";
            this.buttonUpdateTypeAllSheets.ScreenTip = "全シート消費種類更新";
            this.buttonUpdateTypeAllSheets.ShowImage = true;
            this.buttonUpdateTypeAllSheets.SuperTip = "全ての月シート（1〜12）のK列を type シートのキーワードで更新します。";
            this.buttonUpdateTypeAllSheets.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonUpdateTypeAllSheets_Click);
            // 
            // buttonAmazonOrderSummary
            // 
            this.buttonAmazonOrderSummary.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonAmazonOrderSummary.Label = "Amazon CSV";
            this.buttonAmazonOrderSummary.Name = "buttonAmazonOrderSummary";
            this.buttonAmazonOrderSummary.OfficeImageId = "ExportTextFile";
            this.buttonAmazonOrderSummary.ScreenTip = "Amazon 注文CSVサマリ作成";
            this.buttonAmazonOrderSummary.ShowImage = true;
            this.buttonAmazonOrderSummary.SuperTip = "Retail.OrderHistory*.csvから注文サマリCSV (amazon_order_summary.csv) を作成します。";
            this.buttonAmazonOrderSummary.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAmazonOrderSummary_Click);
            // 
            // buttonAmazonCheck
            // 
            this.buttonAmazonCheck.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonAmazonCheck.Label = "Amazon Check";
            this.buttonAmazonCheck.Name = "buttonAmazonCheck";
            this.buttonAmazonCheck.OfficeImageId = "FindDialog";
            this.buttonAmazonCheck.ScreenTip = "Amazon 照合";
            this.buttonAmazonCheck.ShowImage = true;
            this.buttonAmazonCheck.SuperTip = "カード明細のAmazon利用にamazonシートから商品名を記入します。";
            this.buttonAmazonCheck.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAmazonCheck_Click);
            // 
            // buttonUpdateYearSheets
            // 
            this.buttonUpdateYearSheets.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonUpdateYearSheets.Label = "年間消費";
            this.buttonUpdateYearSheets.Name = "buttonUpdateYearSheets";
            this.buttonUpdateYearSheets.OfficeImageId = "ChartColumnClustered";
            this.buttonUpdateYearSheets.ScreenTip = "年間消費更新";
            this.buttonUpdateYearSheets.ShowImage = true;
            this.buttonUpdateYearSheets.SuperTip = "全ての月シート（1〜12）を集計して「年間」シートのマトリクスを更新します。";
            this.buttonUpdateYearSheets.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonUpdateYearSheets_Click);
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonUpdateTypeAllSheets;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAmazonOrderSummary;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAmazonCheck;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonUpdateYearSheets;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
