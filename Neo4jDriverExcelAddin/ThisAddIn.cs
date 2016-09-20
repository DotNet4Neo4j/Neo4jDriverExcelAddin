using Office = Microsoft.Office.Core;

namespace Neo4jDriverExcelAddin
{
    using System;
    using Microsoft.Office.Interop.Excel;
    using Microsoft.Office.Tools;
    using Neo4j.Driver.V1;

    public partial class ThisAddIn
    {
        private CustomTaskPane _customTaskPane;
        private IDriver _driver;
        private Neo4jDriverExcelAddinRibbon _ribbon;

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new Neo4jDriverExcelAddinRibbon();
            _ribbon.ShowHide += RibbonShowHide;
            return _ribbon;
        }

        private void RibbonShowHide(object sender, EventArgs e)
        {
            if (_customTaskPane == null)
                InitializePane();

            if (_customTaskPane != null)
                _customTaskPane.Visible = !_customTaskPane.Visible;
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _driver = GraphDatabase.Driver(new Uri("bolt://localhost")); //TODO: Hard coded Neo4j Instance URL
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            RemoveOrphanedTaskPanes();
            _driver?.Dispose();
        }

        private void RemoveOrphanedTaskPanes()
        {
            try
            {
                for (var i = CustomTaskPanes.Count; i > 0; i--)
                {
                    var ctp = CustomTaskPanes[i - 1];
                    if (ctp.Window == null)
                    {
                        CustomTaskPanes.Remove(ctp);
                        var control = ctp.Control as ExecuteQuery;
                        control?.Dispose();
                    }
                }
            }
            catch (ObjectDisposedException)
            {
            }
        }

        internal ExecuteQuery InitializePane()
        {
            try
            {
                var gotPane = GetPane();
                if (gotPane != null)
                {
                    _customTaskPane = gotPane;

                    return _customTaskPane.Control as ExecuteQuery;
                }

                var executeQueryControl = new ExecuteQuery();
                executeQueryControl.ExecuteCypher += ExecuteCypher;

                _customTaskPane = CustomTaskPanes.Add(executeQueryControl, "Execute Query");
                
                _customTaskPane.Visible = true;
                return executeQueryControl;
            }
            catch
            {
                return null;
            }
        }

        private void ExecuteCypher(object sender, ExecuteCypherQueryArgs e)
        {
            var worksheet = ((Worksheet) Application.ActiveSheet);
           
            using (var session = _driver.Session())
            {
                var result = session.Run(e.Cypher);
                int row = 1;
                foreach (var record in result)
                {
                    var range = worksheet.Range[$"A{row++}"]; //TODO: Hard coded range
                    range.Value2 = record["UserId"].As<string>(); //TODO: Hard coded 'UserId' here.
                }
            }
        }

        /// <summary></summary>
        /// <remarks>
        ///     Based on:
        ///     http://svn.alfresco.com/repos/alfresco-open-mirror/alfresco/COMMUNITYTAGS/V4.0d/root/projects/extensions/AlfrescoOffice2007/AlfrescoWord2007/ThisAddIn.cs
        /// </remarks>
        /// <returns></returns>
        private CustomTaskPane GetPane()
        {
            try
            {
                if (CustomTaskPanes.Count > 0)
                {
                    foreach (var ctp in CustomTaskPanes)
                    {
                        try
                        {
                            if (ctp.Window == Application.ActiveWindow)
                            {
                                return ctp;
                            }
                        }
                        catch
                        {
                            // Likely due to no active window
                            if (ctp.Window == null)
                            {
                                // This is the one
                                return ctp;
                            }
                        }
                    }
                }
            }
            catch
            {
                return null;
            }
            return null;
        }

        #region VSTO generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}