using System;
using System.IO;
using System.Reflection;
using System.Xml.Serialization;

namespace TIA_Extract.Utility
{
    /// <summary>
    /// Settings of the add-in
    /// </summary>
    [XmlRoot]
    [XmlType]
    public class Settings
    {
        public string BlockExtension { get; set; }
        public string DefaultAlarmsClass { get; set; }
        public bool SimplifyTagname { get; set; }

        /// <summary>
        /// Path to the settings file
        /// </summary>
        private static readonly string SettingsFilePath;

        static Settings()
        {
            var settingsDirectoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TIA Add-Ins", Assembly.GetCallingAssembly().GetName().Name, Assembly.GetCallingAssembly().GetName().Version.ToString());
            var settingsDirectory = Directory.CreateDirectory(settingsDirectoryPath);
            SettingsFilePath = Path.Combine(settingsDirectory.FullName, string.Concat(typeof(Settings).Name, ".xml"));
        }

        public Settings()
        {
            BlockExtension = "_Defauts";
            DefaultAlarmsClass = "Alarm";
            SimplifyTagname = true;
        }

        /// <summary>
        /// Loads the Settings file into the setting object if file exist
        /// otherwise creates new file with default settings
        /// </summary>
        /// <returns>Settings object with loaded or default settings</returns>
        public static Settings Load()
        {
            if (File.Exists(SettingsFilePath) == false)
            {
                return new Settings();
            }

            try
            {
                using (FileStream readStream = new FileStream(SettingsFilePath, FileMode.Open))
                {
                    XmlSerializer serializer = new XmlSerializer(typeof(Settings));
                    return serializer.Deserialize(readStream) as Settings;
                }
            }
            catch
            {
                return new Settings();
            }

        }

        /// <summary>
        /// Saves the current configuration to the file
        /// </summary>
        public void Save()
        {
            try
            {
                using (FileStream writeStream = new FileStream(SettingsFilePath, FileMode.Create))
                {
                    XmlSerializer serializer = new XmlSerializer(typeof(Settings));
                    serializer.Serialize(writeStream, this);
                }
            }
            catch
            {
                // Ignore file operation. I know that changed settings will be lost
            }
        }
    }
}
