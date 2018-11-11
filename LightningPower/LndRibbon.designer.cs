namespace LightningPower
{
    partial class LndRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public LndRibbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(LndRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.setupWB = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.editBox2 = this.Factory.CreateRibbonEditBox();
            this.tab1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "LND";
            this.tab1.Name = "tab1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.setupWB);
            this.group2.Label = "Workbook";
            this.group2.Name = "group2";
            // 
            // setupWB
            // 
            this.setupWB.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.setupWB.Image = ((System.Drawing.Image)(resources.GetObject("setupWB.Image")));
            this.setupWB.Label = "Setup Workbook";
            this.setupWB.Name = "setupWB";
            this.setupWB.ShowImage = true;
            this.setupWB.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SetupWB_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.button1);
            this.group3.Items.Add(this.editBox2);
            this.group3.Label = "Bitcoin";
            this.group3.Name = "group3";
            // 
            // button1
            // 
            this.button1.Description = "Generate New Address";
            this.button1.Label = "Generate New Address";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button1_Click);
            // 
            // editBox2
            // 
            this.editBox2.Label = "New Address:";
            this.editBox2.Name = "editBox2";
            this.editBox2.SizeString = "paddingtb1q0tt3rdscteaftam3rktfg37at27qtdctlree7g";
            this.editBox2.Text = null;
            this.editBox2.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.EditBox2_TextChanged);
            // 
            // LndRibbon
            // 
            this.Name = "LndRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton setupWB;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox2;
    }

    partial class ThisRibbonCollection
    {
        internal LndRibbon Ribbon1
        {
            get { return this.GetRibbon<LndRibbon>(); }
        }
    }
}
