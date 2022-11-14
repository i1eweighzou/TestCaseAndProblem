namespace TestCaseAndProblem
{
    [System.ComponentModel.ToolboxItemAttribute(false)]
    partial class ActionsPaneControlTestCase
    {
        /// <summary> 
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ActionsPaneControlTestCase));
            this.textBoxDetail = new System.Windows.Forms.TextBox();
            this.labelPrompt = new System.Windows.Forms.Label();
            this.pictureBoxLog = new System.Windows.Forms.PictureBox();
            this.labelTool = new System.Windows.Forms.Label();
            this.labelTel = new System.Windows.Forms.Label();
            this.flowLayoutPanelRoot = new System.Windows.Forms.FlowLayoutPanel();
            this.flowLayoutPanelDes = new System.Windows.Forms.FlowLayoutPanel();
            this.flowLayoutPanelDesText = new System.Windows.Forms.FlowLayoutPanel();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxLog)).BeginInit();
            this.flowLayoutPanelRoot.SuspendLayout();
            this.flowLayoutPanelDes.SuspendLayout();
            this.flowLayoutPanelDesText.SuspendLayout();
            this.SuspendLayout();
            // 
            // textBoxDetail
            // 
            this.textBoxDetail.AcceptsReturn = true;
            this.textBoxDetail.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBoxDetail.ForeColor = System.Drawing.Color.Crimson;
            this.textBoxDetail.Location = new System.Drawing.Point(3, 22);
            this.textBoxDetail.Multiline = true;
            this.textBoxDetail.Name = "textBoxDetail";
            this.textBoxDetail.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBoxDetail.Size = new System.Drawing.Size(324, 626);
            this.textBoxDetail.TabIndex = 0;
            this.textBoxDetail.Enter += new System.EventHandler(this.textBoxDetail_Enter);
            this.textBoxDetail.Leave += new System.EventHandler(this.textBoxDetail_Leave);
            // 
            // labelPrompt
            // 
            this.labelPrompt.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelPrompt.ForeColor = System.Drawing.Color.DeepSkyBlue;
            this.labelPrompt.Location = new System.Drawing.Point(3, 0);
            this.labelPrompt.Name = "labelPrompt";
            this.labelPrompt.Size = new System.Drawing.Size(109, 19);
            this.labelPrompt.TabIndex = 1;
            this.labelPrompt.Text = "详细编辑项";
            // 
            // pictureBoxLog
            // 
            this.pictureBoxLog.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.pictureBoxLog.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxLog.Image")));
            this.pictureBoxLog.Location = new System.Drawing.Point(3, 3);
            this.pictureBoxLog.Name = "pictureBoxLog";
            this.pictureBoxLog.Size = new System.Drawing.Size(122, 50);
            this.pictureBoxLog.TabIndex = 2;
            this.pictureBoxLog.TabStop = false;
            // 
            // labelTool
            // 
            this.labelTool.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.labelTool.AutoSize = true;
            this.labelTool.Font = new System.Drawing.Font("华文行楷", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelTool.ForeColor = System.Drawing.Color.Red;
            this.labelTool.Location = new System.Drawing.Point(3, 12);
            this.labelTool.Name = "labelTool";
            this.labelTool.Size = new System.Drawing.Size(110, 17);
            this.labelTool.TabIndex = 3;
            this.labelTool.Text = "测评记录工具\r\n";
            this.labelTool.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            // 
            // labelTel
            // 
            this.labelTel.AutoSize = true;
            this.labelTel.Dock = System.Windows.Forms.DockStyle.Top;
            this.labelTel.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelTel.ForeColor = System.Drawing.Color.Red;
            this.labelTel.Location = new System.Drawing.Point(3, 0);
            this.labelTel.Name = "labelTel";
            this.labelTel.Size = new System.Drawing.Size(82, 12);
            this.labelTel.TabIndex = 4;
            this.labelTel.Text = "Tel:2492602";
            // 
            // flowLayoutPanelRoot
            // 
            this.flowLayoutPanelRoot.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.flowLayoutPanelRoot.Controls.Add(this.labelPrompt);
            this.flowLayoutPanelRoot.Controls.Add(this.textBoxDetail);
            this.flowLayoutPanelRoot.Controls.Add(this.flowLayoutPanelDes);
            this.flowLayoutPanelRoot.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.flowLayoutPanelRoot.Location = new System.Drawing.Point(3, 3);
            this.flowLayoutPanelRoot.Name = "flowLayoutPanelRoot";
            this.flowLayoutPanelRoot.Size = new System.Drawing.Size(337, 719);
            this.flowLayoutPanelRoot.TabIndex = 5;
            // 
            // flowLayoutPanelDes
            // 
            this.flowLayoutPanelDes.Controls.Add(this.pictureBoxLog);
            this.flowLayoutPanelDes.Controls.Add(this.flowLayoutPanelDesText);
            this.flowLayoutPanelDes.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flowLayoutPanelDes.Location = new System.Drawing.Point(3, 654);
            this.flowLayoutPanelDes.Name = "flowLayoutPanelDes";
            this.flowLayoutPanelDes.Size = new System.Drawing.Size(324, 56);
            this.flowLayoutPanelDes.TabIndex = 5;
            // 
            // flowLayoutPanelDesText
            // 
            this.flowLayoutPanelDesText.Controls.Add(this.labelTel);
            this.flowLayoutPanelDesText.Controls.Add(this.labelTool);
            this.flowLayoutPanelDesText.Dock = System.Windows.Forms.DockStyle.Right;
            this.flowLayoutPanelDesText.Location = new System.Drawing.Point(131, 3);
            this.flowLayoutPanelDesText.Name = "flowLayoutPanelDesText";
            this.flowLayoutPanelDesText.Size = new System.Drawing.Size(144, 50);
            this.flowLayoutPanelDesText.TabIndex = 3;
            // 
            // ActionsPaneControlTestCase
            // 
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Controls.Add(this.flowLayoutPanelRoot);
            this.Name = "ActionsPaneControlTestCase";
            this.Size = new System.Drawing.Size(343, 725);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxLog)).EndInit();
            this.flowLayoutPanelRoot.ResumeLayout(false);
            this.flowLayoutPanelRoot.PerformLayout();
            this.flowLayoutPanelDes.ResumeLayout(false);
            this.flowLayoutPanelDesText.ResumeLayout(false);
            this.flowLayoutPanelDesText.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxDetail;
        private System.Windows.Forms.Label labelPrompt;
        private System.Windows.Forms.PictureBox pictureBoxLog;
        private System.Windows.Forms.Label labelTool;
        private System.Windows.Forms.Label labelTel;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanelRoot;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanelDes;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanelDesText;
    }
}
