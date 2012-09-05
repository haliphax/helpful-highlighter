namespace HelpfulHighlighter
{
	partial class HelpfulHighlighterRibbon
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

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
			this.tab1 = new Microsoft.Office.Tools.Ribbon.RibbonTab();
			this.grpHelpfulHighlighter = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
			this.chkEnabled = new Microsoft.Office.Tools.Ribbon.RibbonCheckBox();
			this.chkPreserve = new Microsoft.Office.Tools.Ribbon.RibbonCheckBox();
			this.btnColor = new Microsoft.Office.Tools.Ribbon.RibbonButton();
			this.clrHighlight = new System.Windows.Forms.ColorDialog();
			this.tab1.SuspendLayout();
			this.grpHelpfulHighlighter.SuspendLayout();
			this.SuspendLayout();
			// 
			// tab1
			// 
			this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
			this.tab1.Groups.Add(this.grpHelpfulHighlighter);
			this.tab1.Label = "TabAddIns";
			this.tab1.Name = "tab1";
			// 
			// grpHelpfulHighlighter
			// 
			this.grpHelpfulHighlighter.Items.Add(this.chkEnabled);
			this.grpHelpfulHighlighter.Items.Add(this.chkPreserve);
			this.grpHelpfulHighlighter.Items.Add(this.btnColor);
			this.grpHelpfulHighlighter.Label = "Helpful Highlighter";
			this.grpHelpfulHighlighter.Name = "grpHelpfulHighlighter";
			// 
			// chkEnabled
			// 
			this.chkEnabled.Checked = true;
			this.chkEnabled.Label = "Enabled";
			this.chkEnabled.Name = "chkEnabled";
			this.chkEnabled.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.chkEnabled_Click);
			// 
			// chkPreserve
			// 
			this.chkPreserve.Checked = true;
			this.chkPreserve.Label = "Preserve colors";
			this.chkPreserve.Name = "chkPreserve";
			this.chkPreserve.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.chkPreserve_Click);
			// 
			// btnColor
			// 
			this.btnColor.Label = "Choose Color";
			this.btnColor.Name = "btnColor";
			this.btnColor.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.btnColor_Click);
			// 
			// clrHighlight
			// 
			this.clrHighlight.Color = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
			// 
			// HelpfulHighlighterRibbon
			// 
			this.Name = "HelpfulHighlighterRibbon";
			this.RibbonType = "Microsoft.Excel.Workbook";
			this.Tabs.Add(this.tab1);
			this.Load += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonUIEventArgs>(this.HelpfulHighlighterRibbon_Load);
			this.tab1.ResumeLayout(false);
			this.tab1.PerformLayout();
			this.grpHelpfulHighlighter.ResumeLayout(false);
			this.grpHelpfulHighlighter.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
		public Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkEnabled;
		public Microsoft.Office.Tools.Ribbon.RibbonGroup grpHelpfulHighlighter;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton btnColor;
		private System.Windows.Forms.ColorDialog clrHighlight;
		internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkPreserve;
	}

	partial class ThisRibbonCollection : Microsoft.Office.Tools.Ribbon.RibbonReadOnlyCollection
	{
		internal HelpfulHighlighterRibbon HelpfulHighlighterRibbon
		{
			get { return this.GetRibbon<HelpfulHighlighterRibbon>(); }
		}
	}
}
