using System.IO;
using System.Xml.Serialization;

namespace CardinalDirectionGlazing
{
    public class CardinalDirectionGlazingSettings
    {
        public string SelectedRevitLinkInstanceName { get; set; }
        public string SpacesForProcessingButtonName { get; set; }
        public string SpacesOrRoomsForProcessingButtonName { get; set; }

        public static CardinalDirectionGlazingSettings GetSettings()
        {
            CardinalDirectionGlazingSettings cardinalDirectionGlazingSettings = null;
            string assemblyPathAll = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string fileName = "CardinalDirectionGlazingSettings.xml";
            string assemblyPath = assemblyPathAll.Replace("CardinalDirectionGlazing.dll", fileName);

            if (File.Exists(assemblyPath))
            {
                using (FileStream fs = new FileStream(assemblyPath, FileMode.Open))
                {
                    XmlSerializer xSer = new XmlSerializer(typeof(CardinalDirectionGlazingSettings));
                    cardinalDirectionGlazingSettings = xSer.Deserialize(fs) as CardinalDirectionGlazingSettings;
                    fs.Close();
                }
            }
            else
            {
                cardinalDirectionGlazingSettings = null;
            }

            return cardinalDirectionGlazingSettings;
        }

        public void SaveSettings()
        {
            string assemblyPathAll = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string fileName = "CardinalDirectionGlazingSettings.xml";
            string assemblyPath = assemblyPathAll.Replace("CardinalDirectionGlazing.dll", fileName);

            if (File.Exists(assemblyPath))
            {
                File.Delete(assemblyPath);
            }

            using (FileStream fs = new FileStream(assemblyPath, FileMode.Create))
            {
                XmlSerializer xSer = new XmlSerializer(typeof(CardinalDirectionGlazingSettings));
                xSer.Serialize(fs, this);
                fs.Close();
            }
        }
    }
}
