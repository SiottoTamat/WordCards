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
                Textxaml.Content = textfield;
            }
        }
        private string bookmarkfield;
        public string Bookmarkfield
        {
            get
            {
                return bookmarkfield;
            }
            set
            {
                bookmarkfield = value;
                BookmarkNamexaml.Content = bookmarkfield;
            }
        }
        private string idfield;
        public string IDfield
        {
            get
            {
                return textfield;
            }
            set
            {
                idfield = value;
                Idxaml.Content = idfield;
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
            }
        }

    }
}