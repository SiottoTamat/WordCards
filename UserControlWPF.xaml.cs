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
        public UserControlWPF()
        {
            InitializeComponent();
        }

        private void AddCard_Click(object sender, RoutedEventArgs e)
        {
            CardControl card = new CardControl();
            card.Textfield = "Test textfield";
            card.Colorfield = System.Windows.Media.Color.FromRgb(250, 160, 160);
            AddCardtoUI(card);
        }
        public void AddCardtoUI(CardControl card)
        {
            Globals.ThisAddIn.stackpanelCards.Children.Add(card);

        }
    }
}
