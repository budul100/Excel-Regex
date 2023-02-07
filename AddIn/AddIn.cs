using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace ExcelRegex
{
    internal class AddIn
        : IExcelAddIn
    {
        #region Public Methods

        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }

        public void AutoOpen()
        {
            IntelliSenseServer.Install();
        }

        #endregion Public Methods
    }
}