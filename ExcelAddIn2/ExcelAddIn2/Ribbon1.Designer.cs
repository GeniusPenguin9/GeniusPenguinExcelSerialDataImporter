namespace ExcelAddIn2
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.DB_Tab = this.Factory.CreateRibbonTab();
            this.label_bottom = this.Factory.CreateRibbonGroup();
            this.DB_Open = this.Factory.CreateRibbonButton();
            this.DB_Tab.SuspendLayout();
            this.label_bottom.SuspendLayout();
            this.SuspendLayout();
            // 
            // DB_Tab
            // 
            this.DB_Tab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.DB_Tab.Groups.Add(this.label_bottom);
            this.DB_Tab.Label = "DB_Tab";
            this.DB_Tab.Name = "DB_Tab";
            // 
            // label_bottom
            // 
            this.label_bottom.Items.Add(this.DB_Open);
            this.label_bottom.Label = "数据分析工具";
            this.label_bottom.Name = "label_bottom";
            // 
            // DB_Open
            // 
            this.DB_Open.Label = "DB_Open";
            this.DB_Open.Name = "DB_Open";
            this.DB_Open.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DB_Open_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.DB_Tab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.DB_Tab.ResumeLayout(false);
            this.DB_Tab.PerformLayout();
            this.label_bottom.ResumeLayout(false);
            this.label_bottom.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab DB_Tab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup label_bottom;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DB_Open;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
