using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Office = Microsoft.Office.Core;

namespace WordCards_WPF
{
    /// <summary>
    /// Interaction logic for UserControlWPF.xaml
    /// </summary>
    public partial class UserControlWPF : UserControl  
    {
        


        public UserControlWPF()
        {
            InitializeComponent();
            this.ListViewxaml.ItemsSource = ListCardControls;
        }

        #region VARIABLES

        ObservableCollection<CardControl> ListCardControls = new ObservableCollection<CardControl>();
        System.Windows.Media.Color CopiedColor = new Color();
        #endregion
        #region MY METHODS


        #region XML-ListCardControls CONNECTION
        public void LoadXmltoListCardControls(Office.CustomXMLPart xmlPart)
        {

            ListCardControls.Clear();

            Office.CustomXMLNodes XMLnodes = xmlPart.SelectNodes("//node");
            string colorstring = "";
            foreach (Office.CustomXMLNode nodElem in XMLnodes)
            {
                CardControl card = new CardControl();
                card.Textfield = nodElem.Text;


                foreach (Office.CustomXMLNode attr in nodElem.Attributes) // grab the attributes for the node tag
                {

                    if (attr.XML.Contains("id")) { card.IDfield = attr.NodeValue; }
                    if (attr.XML.Contains("bookmark"))
                    {
                        if (attr.NodeValue == "")
                        {
                            attr.NodeValue = "NONE";
                        }
                        card.Bookmarkfield = attr.NodeValue;
                    }
                    if (attr.XML.Contains("color")) { colorstring = attr.NodeValue; }
                }
                string wordcount = "0";
                string pgcount = "0";

                string[] i = colorstring.Split(',');
                System.Windows.Media.Color CardColor = System.Windows.Media.Color.FromRgb(250, 250, 160);
                try
                {
                    card.Colorfield = Color.FromRgb(byte.Parse(i[0]), byte.Parse(i[1]), byte.Parse(i[2]));

                }
                catch
                {
                    card.Colorfield = CardColor;
                }

                ListCardControls.Add(card);
                card.IDfield = FindIndexCard(card);

            }

        }

        public void UpdateXMLFile()
        {
            Office.CustomXMLNodes basenodes = Globals.ThisAddIn.myXML.SelectNodes("//node");


            foreach (Office.CustomXMLNode bnode in basenodes)
            {
                bnode.Delete(); // delete all nodes in the xml
            }
            
            foreach (CardControl card in ListCardControls)
            {
                

                AddXMLNode(Globals.ThisAddIn.myXML, card.IDfield, "0", "0", card.Textfield, card.Bookmarkfield, card.Colorfield);

            }
        }
        public void AddXMLNode(Office.CustomXMLPart TreeviewXMLPart, string id, string wordcount, string pgcount, string text, string bookmark, System.Windows.Media.Color CardColor)// this adds a node for the card in the custom XML file
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

        #endregion
        public string FindIndexCard(CardControl card)
        {
            int idx = 1;
            foreach (CardControl carditem in ListCardControls)
            {
                if (card == carditem)
                {
                    break;
                }
                idx++;
            }
            return idx.ToString();
        }


            
      


        public void ChangeCardColor()
        {
            
            ColorWindow Colordialog = new ColorWindow();
            if(Colordialog.ShowDialog() == true)
            {
                if (ListViewxaml.SelectedItems.Count > 0)
                {
                    foreach(CardControl item in ListViewxaml.SelectedItems)
                    {
                    item.Colorfield = Colordialog.anweredColor;
                    }
                }
                
            }
        }
        public void PasteCardColor(Color newcolor)
        {
            if (ListViewxaml.SelectedItems.Count > 0)
            {
                foreach (CardControl item in ListViewxaml.SelectedItems)
                {
                    item.Colorfield = newcolor;
                }
            }
        }

        #endregion
        #region MENU BUTTONS

        private void Test_Click(object sender, RoutedEventArgs e)
        {
            string message = "";
            foreach (CardControl control in ListCardControls)
            {
                message += control.Textfield + "; " + control.Bookmarkfield + "; " + control.IDfield + "; " + control.Colorfield.ToString() + Environment.NewLine;
            }
            MessageBox.Show(message);
        }
        private void AddCard_Click(object sender, RoutedEventArgs e)
        {
            CardControl card = new CardControl();
            card.Textfield = "New Card";
            
            card.Colorfield = System.Windows.Media.Color.FromRgb(250, 160, 160);
            ListCardControls.Add(card);
            card.IDfield = FindIndexCard(card);
        }
        private void UpdateXML_Click(object sender, RoutedEventArgs e)
        {
            UpdateXMLFile();
        }

        #endregion



    }
}
