using Office = Microsoft.Office.Core;

namespace Neo4jDriverExcelAddin
{
    using System;
    using System.Globalization;
    using System.Linq;
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
            bool forceVisible = false;
            if (_customTaskPane == null)
            {
                InitializePane();
                forceVisible = true;
            }

            if (_customTaskPane != null)
                _customTaskPane.Visible = forceVisible || !_customTaskPane.Visible;
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

        internal ExecuteQuery CurrentControl => _customTaskPane.Control as ExecuteQuery;

        /// <summary>
        /// Gets the appropriate Excel column name given a number index.
        /// </summary>
        /// <remarks>Initial source: http://stackoverflow.com/questions/4583191/incrementation-of-char </remarks>
        private static string GetColNameFromIndex(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = string.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }


        private void ExecuteCypher(object sender, ExecuteCypherQueryArgs e)
        {
            try
            {
                var worksheet = ((Worksheet) Application.ActiveSheet);

                using (var session = _driver.Session())
                {
                    var result = session.Run(e.Cypher);
                    bool isFirstRow = true;
                    int row = 2;
                    foreach (var record in result)
                    {
                        for (int i = 0; i < record.Keys.Count; i++)
                        {
                            var colName = GetColNameFromIndex(i + 1);
                            var key = record.Keys[i];
                            if (isFirstRow)
                                worksheet.Range[$"{colName}1"].Value2 = key;
                            worksheet.Range[$"{colName}{row}"].Value2 = record.Values[key].As<string>();
                        }
                        row++;
                        isFirstRow = false;
                    }
                }
            }
            catch (Neo4jException ex)
            {
                CurrentControl.SetMessage(ex.Message);
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