using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace CardinalDirectionGlazing
{
    public partial class CardinalDirectionGlazingWPF : Window
    {
        public RevitLinkInstance SelectedRevitLinkInstance;
        public string SpacesForProcessingButtonName;
        public string SpacesOrRoomsForProcessingButtonName;

        CardinalDirectionGlazingSettings CardinalDirectionGlazingSettingsItem = null;

        public CardinalDirectionGlazingWPF(List<RevitLinkInstance> revitLinkInstanceList)
        {
            InitializeComponent();
            CardinalDirectionGlazingSettingsItem = CardinalDirectionGlazingSettings.GetSettings();

            listBox_RevitLinkInstance.ItemsSource = revitLinkInstanceList;
            listBox_RevitLinkInstance.DisplayMemberPath = "Name";

            // Устанавливаем сохранённые настройки или значения по умолчанию
            if (CardinalDirectionGlazingSettingsItem != null)
            {
                if (revitLinkInstanceList.FirstOrDefault(li => li.Name == CardinalDirectionGlazingSettingsItem.SelectedRevitLinkInstanceName) != null)
                {
                    listBox_RevitLinkInstance.SelectedItem = revitLinkInstanceList.FirstOrDefault(li => li.Name == CardinalDirectionGlazingSettingsItem.SelectedRevitLinkInstanceName);
                }
                else
                {
                    listBox_RevitLinkInstance.SelectedItem = listBox_RevitLinkInstance.Items[0];
                }

                if (CardinalDirectionGlazingSettingsItem.SpacesForProcessingButtonName == "radioButton_Selected")
                {
                    radioButton_Selected.IsChecked = true;
                }
                else
                {
                    radioButton_All.IsChecked = true;
                }

                if (CardinalDirectionGlazingSettingsItem.SpacesOrRoomsForProcessingButtonName == "radioButton_Spaces")
                {
                    radioButton_Spaces.IsChecked = true;
                }
                else
                {
                    radioButton_Rooms.IsChecked = true;
                }
            }
            else
            {
                listBox_RevitLinkInstance.SelectedItem = listBox_RevitLinkInstance.Items[0];
                radioButton_All.IsChecked = true;
                radioButton_Spaces.IsChecked = true;
            }
        }

        private void btn_Ok_Click(object sender, RoutedEventArgs e)
        {
            SaveSettings();
            this.DialogResult = true;
            this.Close();
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter || e.Key == Key.Space)
            {
                SaveSettings();
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

        private void SaveSettings()
        {
            CardinalDirectionGlazingSettingsItem = new CardinalDirectionGlazingSettings();

            SelectedRevitLinkInstance = listBox_RevitLinkInstance.SelectedItem as RevitLinkInstance;
            if (SelectedRevitLinkInstance != null)
            {
                CardinalDirectionGlazingSettingsItem.SelectedRevitLinkInstanceName = SelectedRevitLinkInstance.Name;
            }

            SpacesForProcessingButtonName = (this.groupBox_SpacesForProcessing.Content as System.Windows.Controls.Grid)
                .Children.OfType<RadioButton>()
                .FirstOrDefault(rb => rb.IsChecked.Value == true)
                .Name;
            CardinalDirectionGlazingSettingsItem.SpacesForProcessingButtonName = SpacesForProcessingButtonName;

            SpacesOrRoomsForProcessingButtonName = (this.groupBox_SpacesOrRoomsForProcessing.Content as System.Windows.Controls.Grid)
                .Children.OfType<RadioButton>()
                .FirstOrDefault(rb => rb.IsChecked.Value == true)
                .Name;
            CardinalDirectionGlazingSettingsItem.SpacesOrRoomsForProcessingButtonName = SpacesOrRoomsForProcessingButtonName;

            CardinalDirectionGlazingSettingsItem.SaveSettings();
        }
    }
}
