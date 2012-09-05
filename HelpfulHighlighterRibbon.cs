/***
 * Project:		Helpful Highlighter Excel Add-In
 * Component:	Options ribbon menu group
 * Author:		Todd Boyd
 * Date:		2010/10/26
 * Description:
 * 
 * The options for Helpful Highlighter are found in the Helpful Highlighter group on the Add-Ins ribbon tab.
 * Helpful Highlighter saves its options so that they will persist between Excel sessions.
 * 
 *		Enable:
 *			If this unchecked, the Add-In will not highlight/cache anything. When it is unchecked, any current
 *			highlighting will be removed.
 *		Preserve color:
 *			If this is unchecked, the background color of the highlighted cells will not be preserved, and
 *			will be reset to XlColor.XlNothing when a new area is selected. Even if Preserve Color is unchecked,
 *			the color buffer will be restored on the next selection change if it is not already empty.
 *		Choose color:
 *			Opens a color dialog to allow the user to choose the color that will be used when highlighting
 *			cells.
 ***/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace HelpfulHighlighter
{
	public partial class HelpfulHighlighterRibbon : OfficeRibbon
	{
		public HelpfulHighlighter.ThisAddIn addin = null;
		
		public HelpfulHighlighterRibbon()
		{
			InitializeComponent();
		}

		// get currently-selected highlighting color		
		public int GetColor()
		{
			return System.Drawing.ColorTranslator.ToOle(clrHighlight.Color);
		}

		// startup process
		private void HelpfulHighlighterRibbon_Load(object sender, RibbonUIEventArgs e)
		{
			// load settings and affect ribbon
			clrHighlight.Color = System.Drawing.ColorTranslator.FromOle((int)HelpfulHighlighter.Properties.Settings.Default.color);
			this.chkPreserve.Checked = (bool)HelpfulHighlighter.Properties.Settings.Default.preserve;
			this.chkEnabled.Checked = (bool)HelpfulHighlighter.Properties.Settings.Default.enabled;
		}
		
		private void chkEnabled_Click(object sender, RibbonControlEventArgs e)
		{
			// set value and clear highlighting
			if(! this.chkEnabled.Checked)
			{
				HelpfulHighlighter.Properties.Settings.Default.enabled = false;
				this.addin.CleanUp();
			}
			// set value
			else
				HelpfulHighlighter.Properties.Settings.Default.enabled = true;

			HelpfulHighlighter.Properties.Settings.Default.Save();
		}

		private void btnColor_Click(object sender, RibbonControlEventArgs e)
		{
			// use color dialog to determine selection
			System.Windows.Forms.DialogResult result = clrHighlight.ShowDialog();
			if(result != System.Windows.Forms.DialogResult.OK)
				return;
			// set value
			HelpfulHighlighter.Properties.Settings.Default.color = System.Drawing.ColorTranslator.ToOle(clrHighlight.Color);
			HelpfulHighlighter.Properties.Settings.Default.Save();
		}

		private void chkPreserve_Click(object sender, RibbonControlEventArgs e)
		{
			// clear highlighting if preserve disabled
			if (! this.chkPreserve.Checked)
				this.addin.CleanUp();
			// set value
			HelpfulHighlighter.Properties.Settings.Default.preserve = this.chkPreserve.Checked;			
			HelpfulHighlighter.Properties.Settings.Default.Save();
		}
	}
}
