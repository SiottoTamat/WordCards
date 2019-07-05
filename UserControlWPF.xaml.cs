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

namespace WordCards_WPF
{
    /// <summary>
    /// Interaction logic for UserControlWPF.xaml
    /// </summary>
    public partial class UserControlWPF : UserControl
    {

        #region VARIABLES
        double cardheight = 75;
        #endregion
        public UserControlWPF()
        {
            InitializeComponent();
        }

        private void AddCard_Click(object sender, RoutedEventArgs e)
        {
            CardControl card = new CardControl();
            card.Text.Content = "Oddly satisfying?";
            Globals.ThisAddIn.stackpanelCards.Children.Add(card);
        }

        public void CreateCard(CardObj cardobj)
        {
            CardControl newcard = new CardControl();
            newcard.TextBlock.Text = cardobj.Text;
            newcard.TooltipCard.Content = cardobj.Text;
            newcard.BookmarkName.Content = cardobj.Bookmark;
            newcard.Id.Content = cardobj.Id;
            System.Windows.Media.Color newColor = System.Windows.Media.Color.FromArgb(cardobj.Color.A, cardobj.Color.R, cardobj.Color.G, cardobj.Color.B);
            SolidColorBrush brush = new SolidColorBrush(newColor);
            newcard.CardCanvas.Background = brush;
            Globals.ThisAddIn.stackpanelCards.Children.Add(newcard);
        }

        public void UpdatePanelCardsfromXML(List<CardObj> ListofCards)
        {
            if (ListofCards != null)
            {
                foreach (CardObj card in ListofCards)
                {
                    CreateCard(card);
                }
            }
        }


    }
}
