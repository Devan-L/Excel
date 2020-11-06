using System.Collections.Generic;
using System.Windows;

namespace WpfApp2
{

    [PropertyChanged.AddINotifyPropertyChangedInterface]
    public partial class MainWindow : Window
    {
        public List<SomeClass> SomeItems { get; set; }
        public List<string> WindowList { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = this;
            this.WindowList = new List<string> { "A", "B", "C" };
            this.SomeItems = new List<SomeClass>
            {
                new SomeClass { Name = "a1"},
                new SomeClass { Name = "b2"},
                new SomeClass { Name = "c3", SelectedThing = "C" }
            };
        }
    }

    public class SomeClass
    {
        public string Name { get; set; }
        public string SelectedThing { get; set; }
    }
}
