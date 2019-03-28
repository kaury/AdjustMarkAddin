namespace AdjustMarkAddin
{
    partial class RibbonAdjust : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonAdjust()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.AdjustAddin = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btn_MarkPanel = this.Factory.CreateRibbonButton();
            this.AdjustAddin.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // AdjustAddin
            // 
            this.AdjustAddin.Groups.Add(this.group1);
            this.AdjustAddin.Label = "CL-Adjust";
            this.AdjustAddin.Name = "AdjustAddin";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btn_MarkPanel);
            this.group1.Label = "标签";
            this.group1.Name = "group1";
            // 
            // btn_MarkPanel
            // 
            this.btn_MarkPanel.Label = "标签面板";
            this.btn_MarkPanel.Name = "btn_MarkPanel";
            this.btn_MarkPanel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_MarkPanel_Click);
            // 
            // RibbonAdjust
            // 
            this.Name = "RibbonAdjust";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.AdjustAddin);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonAdjust_Load);
            this.AdjustAddin.ResumeLayout(false);
            this.AdjustAddin.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab AdjustAddin;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_MarkPanel;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonAdjust RibbonAdjust {
            get { return this.GetRibbon<RibbonAdjust>(); }
        }
    }
}
