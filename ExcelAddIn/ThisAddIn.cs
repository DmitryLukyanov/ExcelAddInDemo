using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Collections.Generic;

namespace ExcelAddIn
{
    public partial class ThisAddIn
    {
        private const string MockedText = "123.456";
        private int _excelColumnLimit;

        private Dictionary<string, ControlSite> _workSheetsControlSite = new Dictionary<string, ControlSite>();
        // TODO: merge these caches
        private Dictionary<string, SuggestionBox> _workSheetsUserControl = new Dictionary<string, SuggestionBox>();

        #region Application level initialization
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookActivate += Application_WorkBookActivate;
            this.Application.WorkbookDeactivate += Application_WorkBookDeactivate;

            // TODO: should be an easier way to do it
            /*
            o [BONUS] When a user presses “ctrl-alt-3” at a specific location, remove the rendering
            o [BONUS] When a user presses “ctrl-alt-1” at a specific location, insert the suggested
               number, then remove the rendering:

               At this point I found only a way to work with OS hooks directly which also works not ideal.
               For example, Alt click is not identified correctly.

               I assume it should be a better approach here, it requires more investigating
             */
            NativeHooks.SetHook(
                (isKeyDown) => 
                {
                    if (isKeyDown(Keys.ControlKey) && 
                        // isKeyDown(Keys.Alt) && // todo: restore
                        isKeyDown(Keys.D3))
                    {
                        ShowControl(selection: null, sheetName: GetActiveWorkSheet().Name);
                    }

                    if (isKeyDown(Keys.ControlKey) &&
                        // isKeyDown(Keys.Alt) && // todo: restore
                        isKeyDown(Keys.D1))
                    {
                        // TODO: should be possible to get an userControl from ControlSite
                        if (_workSheetsUserControl.TryGetValue(GetActiveWorkSheet().Name, out var control))
                        {
                            var currentCell = GetActiveCell();
                            currentCell.Value = control.Text;
                        }
                    }
                });
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            NativeHooks.ReleaseHook();
            this.Application.WorkbookActivate -= Application_WorkBookActivate;
            this.Application.WorkbookDeactivate -= Application_WorkBookDeactivate;

            foreach (var toDispose in _workSheetsControlSite)
            {
                toDispose.Value.Dispose();
            }

            foreach (var toDispose in _workSheetsUserControl)
            {
                toDispose.Value.Dispose();
            }
        }
        #endregion

        #region Workbook level initialization
        private void Application_WorkBookActivate(object activeWorkbook)
        {
            _excelColumnLimit = this.Application.Columns.Count; // 16384 (XFD)

            var workBook = activeWorkbook as Excel.Workbook;
            workBook.SheetActivate += WorkBook_SheetActivate;
            workBook.SheetDeactivate += WorkBook_SheetDeactivate;

            var firstWorkSheet = workBook.Sheets[1] as Excel.Worksheet;
            WorkBook_SheetActivate(firstWorkSheet);
        }

        private void Application_WorkBookDeactivate(object activeWorkbook)
        {
            var workBook = activeWorkbook as Excel.Workbook;
            workBook.SheetActivate -= WorkBook_SheetActivate;
            workBook.SheetDeactivate -= WorkBook_SheetDeactivate;

            var firstWorkSheet = workBook.Sheets[1];
            WorkBook_SheetDeactivate(firstWorkSheet);
        }
        #endregion

        #region Sheet level initialization
        private void WorkBook_SheetActivate(object activeSheet)
        {
            var sheet = activeSheet as Excel.Worksheet;
            sheet.SelectionChange += ThisAddIn_SelectionChange;
            ShowControl(GetActiveCell(), sheet.Name);
        }

        private void WorkBook_SheetDeactivate(object activeSheet)
        {
            var sheet = activeSheet as Excel.Worksheet;
            sheet.SelectionChange -= ThisAddIn_SelectionChange;
        }

        private void ThisAddIn_SelectionChange(Excel.Range selection)
        {
            ShowControl(selection, GetActiveWorkSheet().Name);
        }
        #endregion


        #region Helpers
        private Excel.Worksheet GetActiveWorkSheet() => Application.ActiveSheet;
        private Excel.Range GetActiveCell() => Application.ActiveCell;

        private void ShowControl(Range selection, string sheetName)
        {
            if (_workSheetsControlSite.TryGetValue(sheetName, out var workSheetControlSite))
            {
                if (selection != null && selection.Count == 1 && selection.Column < _excelColumnLimit)
                {
                    workSheetControlSite.Visible = true;
                }
                else
                {
                    workSheetControlSite.Visible = false;
                    return;
                }
            }

            if (selection == null)
            {
                return;
            }

            var expectedRange = selection.Next;

            const string ControlName = "SuggestionBoxControl";
            if (workSheetControlSite?.Visible ?? false)
            {
                // TODO: This requirement is not 100% done in case if resize happens: Same height and width as the selected cell
                workSheetControlSite.Top = expectedRange.Top;
                workSheetControlSite.Left = expectedRange.Left;
                workSheetControlSite.Height = selection.Height;
                workSheetControlSite.Width = selection.Width;
            }
            else
            {
                var control = new SuggestionBox();
                control.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Bottom;
                control.Height = (int)selection.Height;
                control.Width = (int)selection.Width;
                var selectionFont = selection.Font;
                control.Font = new System.Drawing.Font(selectionFont.Name, (float)selectionFont.Size, FontStyle.Regular, GraphicsUnit.Point);
                control.Text = MockedText;
                Microsoft.Office.Tools.Excel.Worksheet worksheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);
                
                //TODO: should be enough to have only one cache
                _workSheetsControlSite[sheetName] = worksheet.Controls.AddControl(control, expectedRange, ControlName);
                _workSheetsUserControl[sheetName] = control;
            }
        }
        #endregion

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
