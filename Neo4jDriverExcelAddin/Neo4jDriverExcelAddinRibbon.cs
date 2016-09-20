using Office = Microsoft.Office.Core;

namespace Neo4jDriverExcelAddin
{
    using System;
    using System.IO;
    using System.Reflection;
    using System.Runtime.InteropServices;

    [ComVisible(true)]
    public class Neo4jDriverExcelAddinRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Neo4jDriverExcelAddin.Neo4jDriverExcelAddinRibbon.xml");
        }

        #endregion

        internal event EventHandler ShowHide;

        public void OnShowHideButton(Office.IRibbonControl control)
        {
            ShowHide?.Invoke(this, null);
        }

        #region Ribbon Callbacks

        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            var resourceNames = asm.GetManifestResourceNames();
            for (var i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (var resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}