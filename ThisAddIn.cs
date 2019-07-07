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
        public UserControlWPF userControlWPF;

        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        public Office.CustomXMLPart myXML;

        

        internal void InitializeCards()
        {



            UserControlWPF controlWPF = new UserControlWPF();
            userControlWPF = controlWPF;
            ElementHost _eh = new ElementHost { Child = controlWPF };
            Usercontrol1 = new UserControl1();
            Usercontrol1.Controls.Add(_eh);
            _eh.Dock = DockStyle.Fill;
            myCustomTaskPane = this.CustomTaskPanes.Add(Usercontrol1, "WordCards");
            myCustomTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            myCustomTaskPane.Visible = true;
            myCustomTaskPane.Width = 300;

            //stackpanelCards = controlWPF.StackPanel;

            Check_CustomXML();
            controlWPF.LoadXmltoListCardControls(myXML);
           // controlWPF.AddListCardControltoUI();
            //myCustomTaskPane.Control.SizeChanged += new EventHandler(CustomTasKPane_SizeChanged);
            //myCustomTaskPane.VisibleChanged += new EventHandler(CustomTaskPane_VisibleChanged);
        }
        #region VARIABLES
        public System.Windows.Media.Color CopiedColor;
        #endregion


        #region MY Methods
        public Office.CustomXMLPart GetMyXML()
        {
            System.Collections.IEnumerator ienumerator = this.Application.ActiveDocument.CustomXMLParts.GetEnumerator();
            Office.CustomXMLPart xmlPart;

            while (ienumerator.MoveNext())
            {

                xmlPart = (Office.CustomXMLPart)ienumerator.Current;
                if (xmlPart.XML.Contains("Data for Cards Add-In"))// this is the file that is coming from this application
                {

                    return xmlPart;
                }

            }
            return null;
        }
        public void Create_CustomXML()// this adds a custom XML file with the right parameters
        {

            var xDoc = new XDocument(
                        new XDeclaration("1.0", "utf-8", "no"),
                        new XComment("Data for Cards Add-In"),
                            new XElement("Root",
                            new XElement("cardList", "")
                                        )
                                    );

            myXML = this.Application.ActiveDocument.CustomXMLParts.Add(xDoc.ToString(), missing);
            //CardTotal = 0;
        }
        public void Check_CustomXML(int firstoccurrence = 0)// this check if there is a custom XML file and if not add one 
        {

            myXML = GetMyXML();
            if (myXML != null)
            {
               MessageBox.Show("This document had already a Custom XML file.");
            }
            else
            {
                Create_CustomXML();
                MessageBox.Show("Created a new Custom XML file.");
            }

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
