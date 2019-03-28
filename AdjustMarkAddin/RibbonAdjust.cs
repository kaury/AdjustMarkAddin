using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Ribbon;

namespace AdjustMarkAddin
{
    public partial class RibbonAdjust
    {
        Dictionary<string, CustomTaskPane> TaskPanels = new Dictionary<string, CustomTaskPane>();

        private void RibbonAdjust_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void btn_MarkPanel_Click(object sender, RibbonControlEventArgs e)
        {
            if (TaskPanels.TryGetValue(Globals.ThisAddIn.Application.ActiveWindow.Hwnd.ToString(), out CustomTaskPane pane))
            {
                pane.Visible = true;
            }
            else
            {
                UCMarksTaskPanel taskPane = new UCMarksTaskPanel();
                Microsoft.Office.Tools.CustomTaskPane customPane = Globals.ThisAddIn.CustomTaskPanes.Add(taskPane, "CL-Adjust", Globals.ThisAddIn.Application.ActiveWindow); /// 这一步很重要将决定是否显示到当前窗口，第三个参数的意思就是依附到那个窗口                
                customPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                customPane.Width = 290;
                TaskPanels.Add(Globals.ThisAddIn.Application.ActiveWindow.Hwnd.ToString(), customPane);
                customPane.Visible = true;
            }
        }
    }
}
