using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Reports.models
{
   public static class PanelExtensions
    {
        public static void OpenPanel2(this Panel panel)
        {
            panel.Dock = DockStyle.Fill;
            panel.Visible = true;
        }
        public static void ClosePanel2(this Panel panel)
        {
            panel.Dock = DockStyle.None;
            panel.Visible = false;
        }
        public static void OpenAddPanel2(this Form form, Panel target, List<Panel> addPanels)
        {
            addPanels.Except(new List<Panel> { target }).ToList().ForEach(x =>
            {
                x.ClosePanel2();
            });

            target.OpenPanel2();
        }
    }
}
