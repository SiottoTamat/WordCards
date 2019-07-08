﻿using Word = Microsoft.Office.Interop.Word;
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

        public ObservableCollection<CardControl> ListCardControls = new ObservableCollection<CardControl>();
        System.Windows.Media.Color CopiedColor = new Color();
        #endregion
        #region MY METHODS

       /* public void RefreshCardsIDs()
        {
            int idx = 1;
            foreach (CardControl carditem in ListCardControls)
            {
                carditem.IDfield = idx.ToString();
                idx++;
            }
        }*/

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



        #region CARD RIGHT CLICK MENU METHODS


        public void ChangeCardColor()
        {
            
            ColorWindow Colordialog = new ColorWindow();
            if(Colordialog.ShowDialog() == true)
            {
                if (ListViewxaml.SelectedItems.Count > 0)
                {
                    foreach(CardControl item in ListViewxaml.SelectedItems)
                    {
                    item.Colorfield = Colordialog.AnweredColor;
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

        internal void DeleteCard()
        {
            if (ListViewxaml.SelectedItems.Count > 0)
            {
                string message = "Are you sure that you want to delete this card?";
                if (ListViewxaml.SelectedItems.Count > 1)
                {
                    message = "Are you sure that you want to delete these cards?";
                }
                if(MessageBox.Show(message,"Delete Cards",MessageBoxButton.YesNo, MessageBoxImage.Warning)==MessageBoxResult.Yes)
                    while (ListViewxaml.SelectedItems.Count>0)
                    {
                        ListCardControls.Remove((CardControl)ListViewxaml.SelectedItem);
                        //RefreshCardsIDs();
                        //ListViewxaml.Items.Refresh();
                    }
            }
        }

        public void MoveUpCard(object sender, RoutedEventArgs e)
        {
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("Move Card Up");
            CardControl item = null;
            int index = -1;

            if (ListViewxaml.SelectedItems.Count != 1) return;
            item = (CardControl)ListViewxaml.SelectedItems[0];
            index = ListCardControls.IndexOf(item);
            if (index > 0)
            {
                MoveText(ListCardControls[index - 1], ListCardControls[index]);
                ListCardControls.Move(index, index - 1);
                ListCardControls[index].IDfield = "change";
                
                //ListCardControls[index - 1].IDfield = (index-1).ToString();
                //ListCardControls[index].IDfield = (index).ToString();
                //RefreshCardsIDs();
                //ListViewxaml.Items.Refresh();

            }
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }

        public void MoveDownCard(object sender, RoutedEventArgs e)
        {
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("Move Card Down");
            CardControl item = null;
            int index = -1;

            if (ListViewxaml.SelectedItems.Count != 1) return;
            item = (CardControl)ListViewxaml.SelectedItems[0];
            index = ListCardControls.IndexOf(item);
            if (index < ListCardControls.Count - 1)
            {
                MoveText(ListCardControls[index], ListCardControls[index + 1]);
                ListCardControls.Move(index, index + 1);
                ListCardControls[index+1].IDfield = "change";


                //ListCardControls[index + 1].IDfield = (index + 1).ToString();
                //ListCardControls[index].IDfield = (index).ToString();
                //RefreshCardsIDs();
                //ListViewxaml.Items.Refresh();
            }
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }

        private void MoveText(CardControl TOPCARD, CardControl BOTTOMCARD)
        {
            //Globals.ThisAddIn.userControlWPF.ListCardControls.IndexOf(card)
            

            if (int.Parse(BOTTOMCARD.IDfield) != 1)// if it is not the first card
            {
            
                    if ((TOPCARD.Bookmarkfield != "None") && (BOTTOMCARD.Bookmarkfield != "None")) // only if both are linked to a bookmark
                    {

                        


                        Word.Bookmark TOPbmk = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks[TOPCARD.Bookmarkfield];
                        Word.Range TOPRange = TOPbmk.Range;

                        // move the text
                        Word.Bookmark BOTTOMbmk = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks[BOTTOMCARD.Bookmarkfield];
                        Word.Range BOTTOMRange = BOTTOMbmk.Range;
                        BOTTOMRange.Cut();

                        Word.Range temprange = Globals.ThisAddIn.Application.ActiveDocument.Range(TOPRange.Start, TOPRange.End);// store range of prevbookmark
                        int topbookmarkLenght = TOPRange.End - TOPRange.Start;// calculate lenght of prevbookmark subtracting end and start
                        string topbkmName = TOPbmk.Name;// save name of prevbookmark
                        TOPbmk.Delete();// delete prevbookmark
                        // add a temporary space character at the start of the previous bookmark
                        // paste the text
                        Word.Range newrange = Globals.ThisAddIn.Application.ActiveDocument.Range(temprange.Start, temprange.Start);
                        newrange.Paste();
                        Word.Range newbkmrkrange = Globals.ThisAddIn.Application.ActiveDocument.Range(newrange.End, newrange.End + topbookmarkLenght);
                        Word.Bookmark newbookmark = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add(topbkmName, newbkmrkrange);// create new bookmark with start at the end of the bookmark and end at start+lenght

                        //BOTTOMbmk.Range.Select();


                        

                    }
                
            }
        }

        public void LinkTextToCard(object sender, RoutedEventArgs e)
        {
            if (ListViewxaml.SelectedItems.Count == 1)
            {
                


            // select the range of text
                Microsoft.Office.Interop.Word.Range range = Globals.ThisAddIn.Application.Selection.Range;

                if (CheckRulesBookmark(range)) // controls that the paragraphs don't overimpose
                {
                    MessageBox.Show("You can't assing overlapping paragraphs to cards.");
                    ((CardControl)ListViewxaml.SelectedItem).Bookmarkfield = "None";
                }
                Microsoft.Office.Interop.Word.Paragraphs paragraphs = Globals.ThisAddIn.Application.Selection.Paragraphs;


                // this checks that is paragraphs and not middle phrases
                if (paragraphs.Count > 0 && ((range.Start == paragraphs.First.Range.Start) && (range.End == paragraphs.Last.Range.End)))
                {

                    // add a bookmark with a specific name calling another method the name will go into the label of the control
                    string nameBookMrk = AddBookmark(range);
                    ((CardControl)ListViewxaml.SelectedItem).Bookmarkfield = nameBookMrk;
                }
                else
                {
                    MessageBox.Show("You need to select one or more paragraphs to add to the Card");
                    ((CardControl)ListViewxaml.SelectedItem).Bookmarkfield = "None";
                }
                ((CardControl)ListViewxaml.SelectedItem).SetStats();


            }
        }

        public string AddBookmark(Microsoft.Office.Interop.Word.Range range)
        {
            string name = "CARD0";
            if (Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Count > 0)
            {
                int maxnum = 0;
                foreach (Microsoft.Office.Interop.Word.Bookmark bookmark in Globals.ThisAddIn.Application.ActiveDocument.Bookmarks) // BookmarkCollection)
                {
                    string nameB = bookmark.Name;
                    int num = int.Parse(nameB.Replace("CARD", ""));
                    if (num > maxnum) { maxnum = num; }
                }
                name = "CARD" + ((maxnum + 1).ToString());
            }

            try
            {
                Microsoft.Office.Interop.Word.Bookmark newbookmark = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks.Add(name, range);
                return name;
            }
            catch
            {
                return "None";
            }
        }

        public bool CheckRulesBookmark(Microsoft.Office.Interop.Word.Range range)// controls that the paragraphs don't overimpose FALSE= correct
        {
            int start = range.Start;
            int end = range.End;
            Microsoft.Office.Interop.Word.Paragraphs paragraphs = range.Paragraphs;
            foreach (Microsoft.Office.Interop.Word.Paragraph paragraph in paragraphs)
            {
                foreach (Microsoft.Office.Interop.Word.Bookmark bookmark in Globals.ThisAddIn.Application.ActiveDocument.Bookmarks)
                {
                    foreach (Microsoft.Office.Interop.Word.Paragraph bparagraph in bookmark.Range.Paragraphs)
                    {
                        string debug1 = bparagraph.ID;
                        string debug2 = paragraph.ID;
                        if (bparagraph.ParaID == paragraph.ParaID)
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }
        #endregion
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

        private void Test2_Click(object sender, RoutedEventArgs e)
        {
            
        }



        #endregion


    }
}
