using System;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace WeldPropUtils.Settings
{
    // Per-user preferences (not per project): %AppData%\SPDS\WeldPropUtils\user.json.
    [DataContract]
    public class UserPreferences
    {
        public const string ThemeDark = "dark";
        public const string ThemeLight = "light";

        // "dark", "light", or empty = follow AutoCAD's COLORTHEME.
        [DataMember(Name = "theme", Order = 0)]
        public string Theme { get; set; }

        public static string FilePath
        {
            get
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                return Path.Combine(appData, "SPDS", "WeldPropUtils", "user.json");
            }
        }

        // Never throws: preferences are a convenience, a broken file just means defaults.
        public static UserPreferences Load()
        {
            try
            {
                if (File.Exists(FilePath))
                {
                    using FileStream stream = File.OpenRead(FilePath);
                    var prefs = (UserPreferences)new DataContractJsonSerializer(typeof(UserPreferences)).ReadObject(stream);
                    if (prefs != null) return prefs;
                }
            }
            catch (Exception)
            {
            }
            return new UserPreferences();
        }

        public void Save()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(FilePath));
                using FileStream stream = File.Create(FilePath);
                using var writer = JsonReaderWriterFactory.CreateJsonWriter(stream, new UTF8Encoding(false), false, true, "  ");
                new DataContractJsonSerializer(typeof(UserPreferences)).WriteObject(writer, this);
            }
            catch (Exception)
            {
            }
        }
    }
}
