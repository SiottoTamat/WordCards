using System;
using System.Collections.Generic;
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
        }

        #region VARIABLES

        List<CardControl> ListCardControls = new List<CardControl>();

        #endregion
        #region MY METHODS
        private void AddCard_Click(object sender, RoutedEventArgs e)
        {
            CardControl card = new CardControl();
            card.Textfield = "Test textfield";
            card.Colorfield = System.Windows.Media.Color.FromRgb(250, 160, 160);
            AddCardtoUI(card);
        }
        public void AddCardtoUI(CardControl card)
        {
            this.StackPanelxaml.Children.Add(card);

        }

        public void AddListCardControltoUI()
        {
            foreach(CardControl cardcontrol in ListCardControls)
            {
                AddCardtoUI(cardcontrol);
            }
        }

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


            }

        }
        #endregion


    }
}
