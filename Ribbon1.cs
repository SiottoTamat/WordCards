﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using Microsoft.Office.Tools.Ribbon;

namespace WordCards_WPF
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void WPFUsrCtrl(object sender, RibbonControlEventArgs e)
        {
            
            Globals.ThisAddIn.InitializeCards();
        }
    }
}
