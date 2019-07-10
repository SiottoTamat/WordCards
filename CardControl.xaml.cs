using System;
using System.Collections.Generic;
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
using Word = Microsoft.Office.Interop.Word;

namespace WordCards_WPF
{
    /// <summary>
    /// Interaction logic for CardControl.xaml
    /// </summary>
    public partial class CardControl : UserControl , INotifyPropertyChanged
    {

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyRaised(string propertyname)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyname));
            }
        }



        public CardControl()
        {
            InitializeComponent();
        }

        private string textfield;
        public string Textfield
        {
            get
            {
                return textfield;
            }
            set
            {
                textfield = value;
                //Textxaml.Text = textfield;
                OnPropertyRaised("Textfield");
            }
        }
        private string bookmarkfield;
        public string Bookmarkfield
        {

            get
            {
                if (bookmarkfield == null)
                {
                    return "None";
                }
                else
                {
                    return bookmarkfield;
                }
            }
            set
            {
                bookmarkfield = value;
                OnPropertyRaised("Bookmarkfield");

             


            }
        
        }

    private string idfield;
        public string IDfield
        {
            get
            {
                return (Globals.ThisAddIn.userControlWPF.ListCardControls.IndexOf(this) + 1).ToString(); 
            }
            set
            {
                idfield = value;
                OnPropertyRaised("IDfield");
            }
        }

        //System.Windows.Media.Color
        private System.Windows.Media.Color colorfield = System.Windows.Media.Color.FromRgb(160, 250, 160);
        public System.Windows.Media.Color Colorfield
        {
            get
            {
                return colorfield;
            }
            set
            {
                colorfield = value;
                System.Windows.Media.Brush Canvasbrush = new SolidColorBrush(colorfield);
                CardCanvas.Background = Canvasbrush;
                OnPropertyRaised("Colorfield");
            }
        }

        public void TestMethod(string test ="Eureka!")
        {
            MessageBox.Show(test);
        }





        #region METHODS

        public void SetStats()
        {
            try
            {
                Wordcountxaml.Content = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks[bookmarkfield].Range.ComputeStatistics(Word.WdStatistic.wdStatisticWords).ToString();
                int rangestrt = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks[bookmarkfield].Range.Start;
                int rangeend = Globals.ThisAddIn.Application.ActiveDocument.Bookmarks[bookmarkfield].Range.End;
                int range = (int)Globals.ThisAddIn.Application.ActiveDocument.Range(rangestrt, rangestrt).get_Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber); ;
                int range2 = (int)Globals.ThisAddIn.Application.ActiveDocument.Range(rangeend, rangeend).get_Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndPageNumber); ;
                Pagesxaml.Content = range.ToString() + "-" + range2.ToString();
            }
            catch { }
        }

        private void Choose_Color_Click(object sender, RoutedEventArgs e)
        {
            //TestMethod();
            Globals.ThisAddIn.userControlWPF.ChangeCardColor();
        }
        private void Copy_Color_Click(object sender, RoutedEventArgs e)
        {
            Globals.ThisAddIn.CopiedColor = this.Colorfield;
            //TestMethod(Globals.ThisAddIn.CopiedColor.ToString());
        }
        private void Paste_Color_Click(object sender, RoutedEventArgs e)
        {
            Color black = Color.FromRgb(0, 0, 0);
            if (Globals.ThisAddIn.CopiedColor != black)
            {
                Globals.ThisAddIn.userControlWPF.PasteCardColor(Globals.ThisAddIn.CopiedColor);
            }
        }

        private void MoveUp_Click(object sender, RoutedEventArgs e)
        {
            Globals.ThisAddIn.userControlWPF.MoveUpCard(sender, e);
        }

        private void MoveDown_Click(object sender, RoutedEventArgs e)
        {
            Globals.ThisAddIn.userControlWPF.MoveDownCard(sender, e);
        }

        private void LinkText_Click(object sender, RoutedEventArgs e)
        {
            Globals.ThisAddIn.userControlWPF.LinkTextToCard(sender, e);
            
            
        }

        private void UnlinkText_Click(object sender, RoutedEventArgs e)
        {
            Globals.ThisAddIn.userControlWPF.UnlinkTextFromCard(sender, e);
        }

        private void DeleteCardText_Click(object sender, RoutedEventArgs e)
        {
            Globals.ThisAddIn.userControlWPF.DeleteCard();
        }
        #endregion

        private void Add_Card_Above_Click(object sender, RoutedEventArgs e)
        {
            
            int idx = Globals.ThisAddIn.userControlWPF.ListCardControls.IndexOf(this);
            Globals.ThisAddIn.userControlWPF.NewCard(index:idx);
            
        }

        private void Add_Card_Below_Click(object sender, RoutedEventArgs e)
        {
            
            int idx = Globals.ThisAddIn.userControlWPF.ListCardControls.IndexOf(this)+1;
            Globals.ThisAddIn.userControlWPF.NewCard(index:idx);
        }

        private void Add_Card_Bottom_Click(object sender, RoutedEventArgs e)
        {
            Globals.ThisAddIn.userControlWPF.NewCard();
        }

        private void CardCanvas_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left && e.ClickCount == 2)
            {
                Globals.ThisAddIn.userControlWPF.FocusOnText(this);
            }
        }
    }
}