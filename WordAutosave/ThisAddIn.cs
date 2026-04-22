using System;
using Microsoft.Office.Tools;

namespace WordAutosave
{
    public partial class ThisAddIn
    {
        public TaskPaneControl taskPaneControl;
        public CustomTaskPane myCustomTaskPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            taskPaneControl = new TaskPaneControl();
            myCustomTaskPane = this.CustomTaskPanes.Add(taskPaneControl, "Word Auto-Save");
            myCustomTaskPane.Visible = true;
            myCustomTaskPane.Width = 300;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
