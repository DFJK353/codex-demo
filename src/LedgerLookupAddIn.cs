using System;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using WinForms = System.Windows.Forms;

namespace LedgerLookup
{
    /// <summary>
    /// VSTO add-in helper that links the "科目余额表" sheet and the "序时账" sheet.
    /// Double-clicking an account row in the balance sheet filters and selects
    /// the matching ledger entries automatically.
    /// </summary>
    public partial class ThisAddIn
    {
        private const string BalanceSheetName = "科目余额表";
        private const string JournalSheetName = "序时账";

        private Excel.Worksheet _balanceSheet;
        private Excel.Worksheet _journalSheet;

        private Excel.AppEvents_SheetBeforeDoubleClickEventHandler _beforeDoubleClickHandler;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _beforeDoubleClickHandler = new Excel.AppEvents_SheetBeforeDoubleClickEventHandler(Application_SheetBeforeDoubleClick);
            this.Application.SheetBeforeDoubleClick += _beforeDoubleClickHandler;
            this.Application.WorkbookOpen += Application_WorkbookOpen;
            this.Application.WorkbookActivate += Application_WorkbookActivate;

            WireActiveWorkbook(this.Application.ActiveWorkbook);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            this.Application.WorkbookOpen -= Application_WorkbookOpen;
            this.Application.WorkbookActivate -= Application_WorkbookActivate;
            if (_beforeDoubleClickHandler != null)
            {
                this.Application.SheetBeforeDoubleClick -= _beforeDoubleClickHandler;
            }
        }

        private void Application_WorkbookOpen(Excel.Workbook workbook)
        {
            WireActiveWorkbook(workbook);
        }

        private void Application_WorkbookActivate(Excel.Workbook workbook)
        {
            WireActiveWorkbook(workbook);
        }

        private void WireActiveWorkbook(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                _balanceSheet = null;
                _journalSheet = null;
                return;
            }

            _balanceSheet = workbook.Worksheets
                .Cast<Excel.Worksheet>()
                .FirstOrDefault(ws => string.Equals(ws.Name, BalanceSheetName, StringComparison.OrdinalIgnoreCase));

            _journalSheet = workbook.Worksheets
                .Cast<Excel.Worksheet>()
                .FirstOrDefault(ws => string.Equals(ws.Name, JournalSheetName, StringComparison.OrdinalIgnoreCase));
        }

        private void Application_SheetBeforeDoubleClick(object sh, Excel.Range target, ref bool cancel)
        {
            if (_balanceSheet == null || _journalSheet == null)
            {
                return;
            }

            var sheet = sh as Excel.Worksheet;
            if (sheet == null || !string.Equals(sheet.Name, _balanceSheet.Name, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            if (target == null || target.Row < 2)
            {
                return; // ignore header and invalid rows
            }

            string accountCode = Convert.ToString(_balanceSheet.Cells[target.Row, 1].Value2)?.Trim();
            if (string.IsNullOrEmpty(accountCode) || accountCode.Contains("合计"))
            {
                return;
            }

            cancel = true; // cancel the default double-click behavior
            ShowLedgerEntries(accountCode);
        }

        private void ShowLedgerEntries(string accountCode)
        {
            if (_journalSheet == null)
            {
                return;
            }

            Excel.Range usedRange = _journalSheet.UsedRange;
            if (usedRange == null)
            {
                WinForms.MessageBox.Show($"未找到科目 {accountCode} 的序时账记录。", "查询结果",
                    WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                return;
            }

            // Clear previous filters if any
            try
            {
                if (_journalSheet.AutoFilter != null)
                {
                    _journalSheet.AutoFilter.ShowAllData();
                }
            }
            catch (Exception)
            {
                // Ignore the exception thrown when there is no filter applied yet.
            }

            int accountField = FindAccountField(usedRange);
            accountField = Math.Max(1, Math.Min(accountField, usedRange.Columns.Count));
            usedRange.AutoFilter(Field: accountField, Criteria1: accountCode);

            Excel.Range accountColumn = usedRange.Columns[accountField];
            Excel.Range found = accountColumn.Find(
                What: accountCode,
                LookIn: Excel.XlFindLookIn.xlValues,
                LookAt: Excel.XlLookAt.xlWhole,
                SearchOrder: Excel.XlSearchOrder.xlByRows,
                SearchDirection: Excel.XlSearchDirection.xlNext,
                MatchCase: false);

            _journalSheet.Activate();

            if (found != null)
            {
                found.Worksheet.Application.Goto(found, true);
                HighlightRow(found.EntireRow, accountCode);
            }
            else
            {
                usedRange.AutoFilter(Field: accountField); // remove the filter when nothing was found
                WinForms.MessageBox.Show($"未找到科目 {accountCode} 的序时账记录。", "查询结果",
                    WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
            }
        }

        private int FindAccountField(Excel.Range usedRange)
        {
            Excel.Range headerRow = usedRange.Rows[1];
            foreach (Excel.Range cell in headerRow.Cells)
            {
                string header = Convert.ToString(cell.Value2)?.Trim();
                if (string.Equals(header, "科目编码", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(header, "科目号", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(header, "科目", StringComparison.OrdinalIgnoreCase))
                {
                    return cell.Column - usedRange.Column + 1; // field index relative to the filtered range
                }
            }

            // Default to the second column (科目编码通常在B列)
            return Math.Max(1, 2 - usedRange.Column + 1);
        }

        private void HighlightRow(Excel.Range row, string accountCode)
        {
            Excel.Application app = row.Application;
            Excel.Range highlightRange = row;

            highlightRange.Select();
            app.ActiveWindow.ScrollRow = Math.Max(1, highlightRange.Row - 3);

            app.StatusBar = false;
            app.StatusBar = $"已定位到科目 {accountCode} 的序时账记录";
        }
    }
}
