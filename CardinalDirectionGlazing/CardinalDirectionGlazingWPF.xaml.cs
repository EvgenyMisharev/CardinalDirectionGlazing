using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;

namespace CardinalDirectionGlazing
{
    public partial class CardinalDirectionGlazingWPF : Window
    {
        public RevitLinkInstance SelectedRevitLinkInstance;
        public CardinalDirectionGlazingWPF(List<RevitLinkInstance> revitLinkInstanceList)
        {
            InitializeComponent();

            listBox_RevitLinkInstance.ItemsSource = revitLinkInstanceList;
            listBox_RevitLinkInstance.DisplayMemberPath = "Name";
            listBox_RevitLinkInstance.SelectedItem = listBox_RevitLinkInstance.Items[0];
        }

        private void btn_Ok_Click(object sender, RoutedEventArgs e)
        {
            SelectedRevitLinkInstance = listBox_RevitLinkInstance.SelectedItem as RevitLinkInstance;
            this.DialogResult = true;
            this.Close();
        }
        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter || e.Key == Key.Space)
            {
                SelectedRevitLinkInstance = listBox_RevitLinkInstance.SelectedItem as RevitLinkInstance;
                this.DialogResult = true;
                this.Close();
            }

            else if (e.Key == Key.Escape)
            {
                this.DialogResult = false;
                this.Close();
            }
        }
        private void btn_Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}
