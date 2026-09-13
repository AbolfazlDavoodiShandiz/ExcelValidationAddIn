using System;
using System.IO;
using System.Reflection;
using Office = Microsoft.Office.Core;

namespace ExcelValidationAddIn.Ribbon
{
    [System.Runtime.InteropServices.ComVisible(true)]
    public class ValidationRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonId)
        {
            using (var stream = Assembly.GetExecutingAssembly()
                .GetManifestResourceStream("ExcelValidationAddIn.Ribbon.Ribbon.xml"))
            {
                if (stream == null)
                {
                    throw new InvalidOperationException(
                        "Ribbon.xml را به عنوان Embedded Resource تنظیم کنید (Build Action = Embedded Resource).");
                }
                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        public void OnRibbonLoad(Office.IRibbonUI ribbonUi)
        {
            _ribbon = ribbonUi;
        }

        public void OnShowPaneToggled(Office.IRibbonControl control, bool pressed)
        {
            if (Globals.ThisAddIn.ValidationTaskPane != null)
            {
                Globals.ThisAddIn.ValidationTaskPane.Visible = pressed;
            }
        }

        public bool GetPanePressed(Office.IRibbonControl control)
        {
            return Globals.ThisAddIn.ValidationTaskPane != null && Globals.ThisAddIn.ValidationTaskPane.Visible;
        }

        public void InvalidatePane()
        {
            _ribbon?.InvalidateControl("ShowPaneButton");
        }
    }
}
