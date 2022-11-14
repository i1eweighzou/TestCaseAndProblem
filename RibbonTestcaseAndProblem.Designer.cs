namespace TestCaseAndProblem
{
    partial class RibbonTestcaseAndProblem : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonTestcaseAndProblem()
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RibbonTestcaseAndProblem));
            this.tabTestCaseAndProblem = this.Factory.CreateRibbonTab();
            this.groupProject = this.Factory.CreateRibbonGroup();
            this.editBoxProjectName = this.Factory.CreateRibbonEditBox();
            this.groupTestcase = this.Factory.CreateRibbonGroup();
            this.buttonGenTestCaseWord = this.Factory.CreateRibbonButton();
            this.groupProblem = this.Factory.CreateRibbonGroup();
            this.buttonGenProblemWord = this.Factory.CreateRibbonButton();
            this.groupSetting = this.Factory.CreateRibbonGroup();
            this.checkBoxSetting = this.Factory.CreateRibbonCheckBox();
            this.checkBoxToggleActionsPane = this.Factory.CreateRibbonCheckBox();
            this.groupLog = this.Factory.CreateRibbonGroup();
            this.editBoxLog = this.Factory.CreateRibbonEditBox();
            this.backgroundWorkerTestRecord = new System.ComponentModel.BackgroundWorker();
            this.notifyIconReport = new System.Windows.Forms.NotifyIcon(this.components);
            this.tabTestCaseAndProblem.SuspendLayout();
            this.groupProject.SuspendLayout();
            this.groupTestcase.SuspendLayout();
            this.groupProblem.SuspendLayout();
            this.groupSetting.SuspendLayout();
            this.groupLog.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabTestCaseAndProblem
            // 
            this.tabTestCaseAndProblem.Groups.Add(this.groupProject);
            this.tabTestCaseAndProblem.Groups.Add(this.groupTestcase);
            this.tabTestCaseAndProblem.Groups.Add(this.groupProblem);
            this.tabTestCaseAndProblem.Groups.Add(this.groupSetting);
            this.tabTestCaseAndProblem.Groups.Add(this.groupLog);
            this.tabTestCaseAndProblem.Label = "离线测试记录";
            this.tabTestCaseAndProblem.Name = "tabTestCaseAndProblem";
            this.tabTestCaseAndProblem.Position = this.Factory.RibbonPosition.BeforeOfficeId("TabHome");
            // 
            // groupProject
            // 
            this.groupProject.Items.Add(this.editBoxProjectName);
            this.groupProject.Label = "项目名称";
            this.groupProject.Name = "groupProject";
            // 
            // editBoxProjectName
            // 
            this.editBoxProjectName.Label = "项目名称";
            this.editBoxProjectName.Name = "editBoxProjectName";
            this.editBoxProjectName.Text = null;
            // 
            // groupTestcase
            // 
            this.groupTestcase.Items.Add(this.buttonGenTestCaseWord);
            this.groupTestcase.Label = "测试用例";
            this.groupTestcase.Name = "groupTestcase";
            // 
            // buttonGenTestCaseWord
            // 
            this.buttonGenTestCaseWord.Label = "生成测试记录";
            this.buttonGenTestCaseWord.Name = "buttonGenTestCaseWord";
            this.buttonGenTestCaseWord.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGenTestCaseWord_Click);
            // 
            // groupProblem
            // 
            this.groupProblem.Items.Add(this.buttonGenProblemWord);
            this.groupProblem.Label = "问题报告";
            this.groupProblem.Name = "groupProblem";
            // 
            // buttonGenProblemWord
            // 
            this.buttonGenProblemWord.Label = "生成问题报告";
            this.buttonGenProblemWord.Name = "buttonGenProblemWord";
            this.buttonGenProblemWord.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGenProblemWord_Click);
            // 
            // groupSetting
            // 
            this.groupSetting.Items.Add(this.checkBoxSetting);
            this.groupSetting.Items.Add(this.checkBoxToggleActionsPane);
            this.groupSetting.Label = "设置";
            this.groupSetting.Name = "groupSetting";
            // 
            // checkBoxSetting
            // 
            this.checkBoxSetting.Label = "详细设置";
            this.checkBoxSetting.Name = "checkBoxSetting";
            this.checkBoxSetting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBoxSetting_Click);
            // 
            // checkBoxToggleActionsPane
            // 
            this.checkBoxToggleActionsPane.Label = "显示文档窗格";
            this.checkBoxToggleActionsPane.Name = "checkBoxToggleActionsPane";
            this.checkBoxToggleActionsPane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.checkBoxToggleActionsPane_Click);
            // 
            // groupLog
            // 
            this.groupLog.Items.Add(this.editBoxLog);
            this.groupLog.Label = "日志";
            this.groupLog.Name = "groupLog";
            // 
            // editBoxLog
            // 
            this.editBoxLog.Label = "生成进度";
            this.editBoxLog.MaxLength = 200;
            this.editBoxLog.Name = "editBoxLog";
            this.editBoxLog.Text = null;
            // 
            // backgroundWorkerTestRecord
            // 
            this.backgroundWorkerTestRecord.WorkerReportsProgress = true;
            this.backgroundWorkerTestRecord.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorkerTestRecord_DoWork);
            this.backgroundWorkerTestRecord.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorkerTestRecord_ProgressChanged);
            this.backgroundWorkerTestRecord.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorkerTestRecord_RunWorkerCompleted);
            // 
            // notifyIconReport
            // 
            this.notifyIconReport.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIconReport.Icon")));
            this.notifyIconReport.Text = "测试记录";
            this.notifyIconReport.Visible = true;
            // 
            // RibbonTestcaseAndProblem
            // 
            this.Name = "RibbonTestcaseAndProblem";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabTestCaseAndProblem);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonTestcaseAndProblem_Load);
            this.tabTestCaseAndProblem.ResumeLayout(false);
            this.tabTestCaseAndProblem.PerformLayout();
            this.groupProject.ResumeLayout(false);
            this.groupProject.PerformLayout();
            this.groupTestcase.ResumeLayout(false);
            this.groupTestcase.PerformLayout();
            this.groupProblem.ResumeLayout(false);
            this.groupProblem.PerformLayout();
            this.groupSetting.ResumeLayout(false);
            this.groupSetting.PerformLayout();
            this.groupLog.ResumeLayout(false);
            this.groupLog.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabTestCaseAndProblem;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupTestcase;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupProblem;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupSetting;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupProject;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxProjectName;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonGenTestCaseWord;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonGenProblemWord;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxSetting;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBoxToggleActionsPane;
        private System.ComponentModel.BackgroundWorker backgroundWorkerTestRecord;
        private System.Windows.Forms.NotifyIcon notifyIconReport;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupLog;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBoxLog;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonTestcaseAndProblem RibbonTestcaseAndProblem
        {
            get { return this.GetRibbon<RibbonTestcaseAndProblem>(); }
        }
    }
}
