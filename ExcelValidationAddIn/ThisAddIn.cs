using System;
using ExcelValidationAddIn.Ribbon;
using ExcelValidationAddIn.UI;
using Microsoft.Office.Tools;
using Office = Microsoft.Office.Core;

namespace ExcelValidationAddIn
{
    public partial class ThisAddIn
    {
        private ValidationPaneControl _paneControl;
        private ValidationRibbon _ribbon;

        /// <summary>Exposed so the ribbon's toggle button can show/hide the pane.</summary>
        public CustomTaskPane ValidationTaskPane { get; private set; }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _paneControl = new ValidationPaneControl(Application);
            ValidationTaskPane = CustomTaskPanes.Add(_paneControl, "اعتبارسنجی اسناد حسابداری");
            ValidationTaskPane.Width = 340;
            ValidationTaskPane.Visible = true;
            ValidationTaskPane.VisibleChanged += (s, args) => _ribbon?.InvalidatePane();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Excel often kills the process before Shutdown reliably fires;
            // this mainly helps clean re-loads while developing (F5 in VS).
            if (ValidationTaskPane != null)
            {
                CustomTaskPanes.Remove(ValidationTaskPane);
            }
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new ValidationRibbon();
            return _ribbon;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify the contents
        /// of this method with the code editor. Normally lives in the
        /// auto-generated ThisAddIn.Designer.cs alongside it.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}
