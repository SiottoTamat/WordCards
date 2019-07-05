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
    /// Interaction logic for CardControl.xaml
    /// </summary>
    public partial class CardControl : UserControl
    {
        public string CardHeight { get; set; } = "75";

        public CardControl()
        {
            InitializeComponent();
        }

        

        private void CardCanvas_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            //((Canvas)sender).Background = Brushes.Azure;
        }
    }
}
