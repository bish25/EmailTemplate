
using Microsoft.Win32;
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

namespace EmailTemplate
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {


        TextPointer caretPos;
        public MainWindow()
        {
            InitializeComponent();
            listView.Items.Add("##NAME##");
            listView.Items.Add("##FEE##");

        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openTemplate = new OpenFileDialog();
            var result = openTemplate.ShowDialog();
            if (result != false)
            {
                try {
                    richTextBox.Document.Blocks.Clear();
                    Microsoft.Office.Interop.Word.Application wordObject = new Microsoft.Office.Interop.Word.Application();
                    object File = openTemplate.FileName; //this is the path
                    object nullobject = System.Reflection.Missing.Value; Microsoft.Office.Interop.Word.Application wordobject = new Microsoft.Office.Interop.Word.Application();
                    wordobject.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone; Microsoft.Office.Interop.Word._Document docs = wordObject.Documents.Open(ref File, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject); docs.ActiveWindow.Selection.WholeStory();
                    docs.ActiveWindow.Selection.Copy();
                    this.richTextBox.Paste();
                    docs.Close(ref nullobject, ref nullobject, ref nullobject);
                    wordobject.Quit(ref nullobject, ref nullobject, ref nullobject);
                }
                catch
                {
                    MessageBox.Show("Failed To Open Template");
                }
            }
        }

        private void Insert(object sender, MouseButtonEventArgs e)
        {
            string value = listView.SelectedItem.ToString();
            richTextBox.Focus();
            richTextBox.CaretPosition = caretPos;
            Clipboard.SetText(value);
            richTextBox.Paste();
        }

        private void Save(object sender, RoutedEventArgs e)
        {

        }

        private void SaveCaretPostion(object sender, RoutedEventArgs e)
        {
            caretPos = richTextBox.CaretPosition;
        }
    }
}
