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
using System.Drawing;

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

        #region VARIABLES 

        public int CheckifPanelOn = 0;
        private UserControl1 Usercontrol1;
        public StackPanel stackpanelCards;
        public List<CardObj> ListofCards = new List<CardObj>(); // the list of cards as CardObj
        public int CardTotal = 0;


        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        public Office.CustomXMLPart myXML;

        #endregion
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


            bool xmlpresent = Check_CustomXML();
            if (xmlpresent)
            {
                bool loadList = UpdateListCardfromXML();
                if (loadList)
                {
                    controlWPF.UpdatePanelCardsfromXML(ListofCards);
                }
            }
             
        }

        #region XML MANAGEMENT
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
        public void AddXMLNode(Office.CustomXMLPart TreeviewXMLPart, string id, string wordcount, string pgcount, string text, string bookmark, Color CardColor)// this adds a node for the card in the custom XML file
        {
            string color = CardColor.R.ToString() + "," + CardColor.G.ToString() + "," + CardColor.B.ToString();
            Office.CustomXMLNode node = TreeviewXMLPart.SelectSingleNode("//cardList[1]");//[@xmlns=CARDS]
            Office.CustomXMLNode childNode;
            if (node == null) { MessageBox.Show("Node = null"); }
            try
            {

                node.AppendChildNode("node", "", Office.MsoCustomXMLNodeType.msoCustomXMLNodeElement, text);
                childNode = node.LastChild;
                childNode.AppendChildNode("id", "", Office.MsoCustomXMLNodeType.msoCustomXMLNodeAttribute, id);
                childNode.AppendChildNode("words", "", Office.MsoCustomXMLNodeType.msoCustomXMLNodeAttribute, wordcount);
                childNode.AppendChildNode("Pages", "", Office.MsoCustomXMLNodeType.msoCustomXMLNodeAttribute, pgcount);
                childNode.AppendChildNode("bookmark", "", Office.MsoCustomXMLNodeType.msoCustomXMLNodeAttribute, bookmark);
                childNode.AppendChildNode("color", "", Office.MsoCustomXMLNodeType.msoCustomXMLNodeAttribute, color);

            }
            catch
            {
                MessageBox.Show("problem: " + TreeviewXMLPart.DocumentElement.XML + "\n---------\n" + TreeviewXMLPart.XML);
            }

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
            CardTotal = 0;
        }
        
        public bool Check_CustomXML()// this check if there is a custom XML file and if not add one 
        {

            myXML = GetMyXML();
            if (myXML != null)
            {
                MessageBox.Show("This document had already a Custom XML file.");
                return true;
            }
            else
            {
                Create_CustomXML();
                MessageBox.Show("Created a new Custom XML file.");
                return false;
            }

        }

        public bool UpdateListCardfromXML()
        {
            if (myXML != null)
            {
                Office.CustomXMLNodes XMLnodes = myXML.SelectNodes("//node");
                foreach (Office.CustomXMLNode nodElem in XMLnodes)
                {
                    CardObj cardobj = new CardObj();
                    string color = "";

                    cardobj.Text = nodElem.Text;


                    foreach (Office.CustomXMLNode attr in nodElem.Attributes) // grab the attributes for the node tag
                    {

                        if (attr.XML.Contains("id")) { cardobj.Id = int.Parse(attr.NodeValue); }
                        if (attr.XML.Contains("bookmark"))
                        {
                            if (attr.NodeValue == "")
                            {
                                attr.NodeValue = "NONE";
                            }
                            cardobj.Bookmark = attr.NodeValue;
                        }
                        if (attr.XML.Contains("color")) { color = attr.NodeValue; }
                    }


                    
                    string[] i = color.Split(',');
                    
                    try
                    {
                        cardobj.Color = Color.FromArgb(int.Parse(i[0]), int.Parse(i[1]), int.Parse(i[2]));

                    }
                    catch
                    {
                        cardobj.Color = Color.WhiteSmoke;
                    }

                    ListofCards.Add(cardobj);
                    

                }
                return true;
            }
            else
            {
                MessageBox.Show("error in UpdateListCardfromXML");
                return false;
            }
        }
        
        /*
        public void Update_XML(int mode) // select the Xml file and update it accordingly to the TreeView. 0=add; 1=delete; and 2=move (perhaps)
        {
            Office.CustomXMLParts xmlFiles = this.Application.ActiveDocument.CustomXMLParts.SelectByNamespace("");

            foreach (Office.CustomXMLPart xmlFile in xmlFiles)// this selects the file without a namespace
            {
                if (xmlFile.XML.Contains("--Data for Cards Add-In--")) // this is just a precaution if there are other custom xml files without namespace in the document
                {
                    if (mode == 0)// in this case it only adds a node
                    {
                        //create a new file
                        Office.CustomXMLNodes basenodes = xmlFile.SelectNodes("//node");


                        foreach (Office.CustomXMLNode bnode in basenodes)
                        {
                            bnode.Delete(); // delete all nodes in the xml
                        }
                        CardTotal = 1;
                        Control.ControlCollection CardPanels = myUserControl1.flowPanel_Nodes.Controls;
                        foreach (Control Cardpanel in CardPanels)
                        {
                            Control.ControlCollection cardObjs = Cardpanel.Controls;
                            string text = "";
                            string words = "";
                            string pages = "";
                            string id = "";
                            string bookmark = nobookmark;
                            Color CardColor = Cardpanel.BackColor;
                            foreach (Control Obj in cardObjs)
                            {

                                {

                                    string caseswitch = Obj.Name;
                                    switch (caseswitch)
                                    {
                                        case "Text":
                                            text = Obj.Text;
                                            break;
                                        case "words":
                                            words = Obj.Text;
                                            break;
                                        case "Pages":
                                            pages = Obj.Text;
                                            break;
                                        case "ID":
                                            id = CardTotal.ToString();
                                            CardTotal++;

                                            break;
                                        case "Bookmark":
                                            bookmark = Obj.Text;
                                            break;

                                    }
                                }
                            }


                            AddXMLNode(xmlFile, id, words, pages, text, bookmark, CardColor);

                        }


                    }
                    else if (mode == 1)// in this case deletes the file, crea
                    {
                        xmlFile.LoadXML("<!--Data for Cards Add-In--><Root>\n< cardList ></ cardList >\n</ Root > ");// this should delete all the nodes
                    }
                    else if (mode == 2)
                    {

                    }
                    else
                    {
                        MessageBox.Show("Error in the input of Update_XML");
                    }

                }


            }

        }
        */
        #endregion







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
