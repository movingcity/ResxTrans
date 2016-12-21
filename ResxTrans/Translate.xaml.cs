using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using MahApps.Metro.Controls;

namespace ResxTrans
{
    /// <summary>
    /// Interaction logic for Translate.xaml
    /// </summary>
    public partial class Translate : MetroWindow
    {
        public List<ResxEntity> ResxEntitys { get; set; }


        public Translate()
        {
            InitializeComponent();
            this.DataContext = this;
        }
     
    }
}
