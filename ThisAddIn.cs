using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Windows.Forms.Integration;
using System.Windows.Forms;
using System.Windows.Controls;

namespace WordCards_WPF
{
    public partial class ThisAddIn
    {
        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

          

        public int CheckifPanelOn = 0;
        private UserControl1 Usercontrol1;
        public StackPanel stackpanelCards;

        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        public Office.CustomXMLPart myXML;

        #region MY Methods

        internal void InitializeCards()
        {
            UserControlWPF controlWPF = new UserControlWPF();
            ElementHost _eh = new ElementHost { Child = controlWPF };
            Usercontrol1 = new UserControl1();
            Usercontrol1.Controls.Add(_eh);
            _eh.Dock = DockStyle.Fill;
            myCustomTaskPane = this.CustomTaskPanes.Add(Usercontrol1, "WordCards");
            myCustomTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            myCustomTaskPane.Visible = true;
            myCustomTaskPane.Width = 300;

            stackpanelCards = controlWPF.StackPanel; 
            //myCustomTaskPane.Control.SizeChanged += new EventHandler(CustomTasKPane_SizeChanged);
            //myCustomTaskPane.VisibleChanged += new EventHandler(CustomTaskPane_VisibleChanged);
        }

        #endregion




        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }



        #endregion
    
    }
}
