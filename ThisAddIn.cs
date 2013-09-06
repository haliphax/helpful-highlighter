/***
 * Project:		Helpful Highlighter Excel Add-In
 * Component:	Core Add-In class
 * Author:		Todd Boyd
 * Date:		2010/10/26
 * Description:
 *
 * This Excel Add-In highlights the current row and column (but not the currently selected area) to assist
 * users with cognitive and/or visual impairments in maintaining focus through the use of visual cues. The
 * background color of the cells is actually changed on-the-fly in order to facilitate this functionality.
 *
 * The highlighting will be removed prior to printing or saving the worksheet in order to prevent the Add-In
 * interfering with the spreadsheets it is highlighting for the user.
 ***/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Excel.Extensions;
using Microsoft.Office.Interop.Excel;
using System.Collections;

namespace HelpfulHighlighter
{
	public partial class ThisAddIn
	{
		private Range hlRange = null; // row and column ranges for highlighting
		private HelpfulHighlighterRibbon ribbon = null; // ribbon object (options)
		private List<object[]> old = null;
		private Microsoft.Office.Interop.Excel.Worksheet oldSheet = null;

		// startup process
		private void ThisAddIn_Startup(object sender, System.EventArgs e)
		{
			// intialize
			{
				// cache for color and colorindex properties
				this.old = new List<object[]>();
			}

			// bind events
			{
				// selection changed
				this.Application.SheetSelectionChange += new AppEvents_SheetSelectionChangeEventHandler(this.Sheet_Selection_Changed);
				// before save
				this.Application.WorkbookBeforeSave += new AppEvents_WorkbookBeforeSaveEventHandler(this.Workbook_Before_Save);
				// before print
				this.Application.WorkbookBeforePrint += new AppEvents_WorkbookBeforePrintEventHandler(this.Workbook_Before_Print);
			}
		}

		// shutdown process
		private void ThisAddIn_Shutdown(object sender, System.EventArgs e) { /* nothing to do */ }

		// remove highlighting
		public void CleanUp(params object[] p)
		{
			// pull parameter
			bool closeBuffer = true;
			if(p.Count() > 0)
				closeBuffer = (bool)p[0];
			// nothing to do?
			if(this.hlRange == null)
				return;
			// buffer on
			this.Application.ScreenUpdating = false;
			XlCalculation calc = this.Application.Calculation;
			this.Application.Calculation = XlCalculation.xlCalculationManual;

			// only restore colors if preserve is set and there are colors to restore in the first place
			if(this.old.Count > 0)
			{
				// cycle through "remembered" colors
				IEnumerator enu = this.old.GetEnumerator();

				while(enu.MoveNext())
				{
					object[] curr = (object[])enu.Current;
					Interior intr;
					// get old sheet's interior
					if (this.oldSheet != null)
						intr = ((Range)oldSheet.Cells[curr[0], curr[1]]).Interior;
					// get current sheet's interior
					else
						intr = ((Range)this.Application.Cells[curr[0], curr[1]]).Interior;
					// restore colors
					intr.Color = curr[2];
					intr.ColorIndex = curr[3];
				}

				this.old.Clear();
			}
			// set colors to "nothing" -- preserve is disabled
			else if(this.hlRange != null)
				this.hlRange.Interior.ColorIndex = XlColorIndex.xlColorIndexNone;

			// buffer off
			if(closeBuffer)
				this.Application.ScreenUpdating = true;
			this.Application.Calculation = calc;
		}

		// highlight selection's entire row/column (excluding selected area)
		private void HighLight()
		{
			// nothing to do?
			if(this.hlRange == null)
				return;
			// buffer on
			this.Application.ScreenUpdating = false;
			XlCalculation calc = this.Application.Calculation;
			this.Application.Calculation = XlCalculation.xlCalculationManual;
			Range sel = (Range)this.Application.Selection;

			// preserve colors?
			if(HelpfulHighlighter.Properties.Settings.Default.preserve)
			{
				// iterate through highlight range and save color/colorindex to cache
				IEnumerator enu = this.hlRange.GetEnumerator();

				while(enu.MoveNext())
				{
					Range curr = (Range)enu.Current;
					this.old.Add(new object[] { curr.Row, curr.Column, curr.Interior.Color, curr.Interior.ColorIndex });
				}
			}

			// highlight row and column ranges
			this.hlRange.Interior.Color = this.ribbon.GetColor();
			// buffer off
			this.Application.ScreenUpdating = true;
			this.Application.Calculation = calc;
			// remember last sheet highlighted
			this.oldSheet = (Microsoft.Office.Interop.Excel.Worksheet)this.Application.ActiveSheet;
		}

		// selection changed; re-highlight
		private void Sheet_Selection_Changed(object sh, Range target)
		{
			// do nothing if addin disabled
			if (! HelpfulHighlighter.Properties.Settings.Default.enabled)
				return;
			// clean up old highlighting, but keep buffer on
			if(this.hlRange != null)
				this.CleanUp(new object[] { false });
			Range sel = (Range)this.Application.Selection;

			// if we have an entire row, and entire column, or just a ton of cells selected, don't highlight
			if(sel.Rows.Cells.Count > this.Application.ActiveWindow.VisibleRange.Rows.Cells.Count
				|| sel.Columns.Cells.Count > this.Application.ActiveWindow.VisibleRange.Columns.Cells.Count)
			{
				this.Application.ScreenUpdating = true;
				return;
			}

			// get cell ranges for areas to highlight
			Range colsAhead = this.Application.get_Range(this.Application.Cells[target.Row, sel.Column + sel.Columns.Count], this.Application.Cells[target.Row, target.Application.ActiveWindow.VisibleRange.Column + target.Application.ActiveWindow.VisibleRange.Columns.Count]);
			// substitute first cell to the right of selection if on column 1
			Range colsBehind = (
				sel.Column <= 1
					? (Range)this.Application.Cells[sel.Row, sel.Column + sel.Columns.Count]
					: this.Application.get_Range(this.Application.Cells[target.Row, 1], this.Application.Cells[target.Row, sel.Column - 1]));
			Range rowsBelow = this.Application.get_Range(this.Application.Cells[sel.Row + sel.Rows.Count, target.Column], this.Application.Cells[target.Application.ActiveWindow.VisibleRange.Row + target.Application.ActiveWindow.VisibleRange.Rows.Count, target.Column]);
			// substitute first cell above selection if on row 1
			Range rowsAbove = (
				sel.Row <= 1
					? (Range)this.Application.Cells[sel.Row + sel.Rows.Count, sel.Column]
					: this.Application.get_Range(this.Application.Cells[1, target.Column], this.Application.Cells[sel.Row - 1, target.Column]));

			this.hlRange = this.Application.Union(rowsAbove, rowsBelow, colsBehind, colsAhead,
				missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
			// highlight row/column
			this.HighLight();
		}

		// clear highlighting before save
		private void Workbook_Before_Save(Microsoft.Office.Interop.Excel.Workbook wb, bool SaveAsUI, ref bool Cancel)
		{
			this.CleanUp();
		}

		// clear highlighting before print
		private void Workbook_Before_Print(Microsoft.Office.Interop.Excel.Workbook wb, ref bool Cancel)
		{
			this.CleanUp();
		}

		// load ribbon
		protected override Microsoft.Office.Tools.Ribbon.OfficeRibbon[] CreateRibbonObjects()
		{
			if (this.ribbon == null)
			{
				this.ribbon = new HelpfulHighlighterRibbon();
				this.ribbon.addin = this;
			}

			return new Microsoft.Office.Tools.Ribbon.OfficeRibbon[] { this.ribbon };
		}

		#region VSTO generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(ThisAddIn_Startup);
			this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
		}

		#endregion
	}
}
