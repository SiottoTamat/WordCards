using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordCards_WPF
{
    public class CardObj
    {
        public int Id { get; set; }
        public string Text { get; set; }
        public string Bookmark { get; set; }
        public string StartBkmPage { get; set; }
        public string EndBkmPage { get; set; }
        public string WordCount { get; set; }
        public System.Drawing.Color Color { get; set; }

    }
}
