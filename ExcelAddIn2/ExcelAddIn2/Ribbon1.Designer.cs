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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl7 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl8 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl9 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl10 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl11 = this.Factory.CreateRibbonDropDownItem();
            this.DB_Tab = this.Factory.CreateRibbonTab();
            this.label_bottom = this.Factory.CreateRibbonGroup();
            this.DB_Open = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.XAxis = this.Factory.CreateRibbonDropDown();
            this.YAxis = this.Factory.CreateRibbonDropDown();
            this.DataCreate = this.Factory.CreateRibbonButton();
            this.DB_Tab.SuspendLayout();
            this.label_bottom.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // DB_Tab
            // 
            this.DB_Tab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.DB_Tab.Groups.Add(this.label_bottom);
            this.DB_Tab.Groups.Add(this.group1);
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
            this.DB_Open.Label = "打开数据文件";
            this.DB_Open.Name = "DB_Open";
            this.DB_Open.ShowImage = true;
            this.DB_Open.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DB_Open_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.XAxis);
            this.group1.Items.Add(this.YAxis);
            this.group1.Items.Add(this.DataCreate);
            this.group1.Label = "图表生成工具";
            this.group1.Name = "group1";
            // 
            // XAxis
            // 
            ribbonDropDownItemImpl2.Label = "每分钟";
            ribbonDropDownItemImpl2.Tag = "2";
            ribbonDropDownItemImpl3.Label = "每小时";
            ribbonDropDownItemImpl3.Tag = "120";
            ribbonDropDownItemImpl4.Label = "每天";
            ribbonDropDownItemImpl4.Tag = "2880";
            this.XAxis.Items.Add(ribbonDropDownItemImpl1);
            this.XAxis.Items.Add(ribbonDropDownItemImpl2);
            this.XAxis.Items.Add(ribbonDropDownItemImpl3);
            this.XAxis.Items.Add(ribbonDropDownItemImpl4);
            this.XAxis.Label = "横坐标";
            this.XAxis.Name = "XAxis";
            // 
            // YAxis
            // 
            ribbonDropDownItemImpl6.Label = "流量L/min";
            ribbonDropDownItemImpl6.Tag = "C";
            ribbonDropDownItemImpl7.Label = "转速rpm";
            ribbonDropDownItemImpl7.Tag = "D";
            ribbonDropDownItemImpl8.Label = "电流A";
            ribbonDropDownItemImpl8.Tag = "E";
            ribbonDropDownItemImpl9.Label = "入口压力mmHg";
            ribbonDropDownItemImpl9.Tag = "F";
            ribbonDropDownItemImpl10.Label = "出口压力mmHg";
            ribbonDropDownItemImpl10.Tag = "G";
            ribbonDropDownItemImpl11.Label = "系统电压V";
            ribbonDropDownItemImpl11.Tag = "H";
            this.YAxis.Items.Add(ribbonDropDownItemImpl5);
            this.YAxis.Items.Add(ribbonDropDownItemImpl6);
            this.YAxis.Items.Add(ribbonDropDownItemImpl7);
            this.YAxis.Items.Add(ribbonDropDownItemImpl8);
            this.YAxis.Items.Add(ribbonDropDownItemImpl9);
            this.YAxis.Items.Add(ribbonDropDownItemImpl10);
            this.YAxis.Items.Add(ribbonDropDownItemImpl11);
            this.YAxis.Label = "纵坐标";
            this.YAxis.Name = "YAxis";
            // 
            // DataCreate
            // 
            this.DataCreate.Label = "生成数据";
            this.DataCreate.Name = "DataCreate";
            this.DataCreate.ShowImage = true;
            this.DataCreate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DataCreate_Click);
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
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab DB_Tab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup label_bottom;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DB_Open;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton DataCreate;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown XAxis;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown YAxis;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
