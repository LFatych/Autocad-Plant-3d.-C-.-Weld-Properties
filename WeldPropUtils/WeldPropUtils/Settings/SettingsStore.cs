using Autodesk.ProcessPower.PlantInstance;
using Autodesk.ProcessPower.ProjectManager;
using System;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Text;

namespace WeldPropUtils.Settings
{
    // Reads and writes <Plant project folder>\SPDS\WeldPropUtils.json.
    // JSON via DataContractJsonSerializer: built into both .NET Framework 4.8 and .NET 8, so the plugin ships
    // no Newtonsoft.Json that could clash with the copy AutoCAD loads.
    public static class SettingsStore
    {
        public const string FolderName = "SPDS";
        public const string FileName = "WeldPropUtils.json";

        // Null when no Plant project is open.
        public static string GetSettingsPath()
        {
            PlantProject project = PlantApplication.CurrentProject;
            if (project == null || string.IsNullOrEmpty(project.ProjectFolderPath)) return null;
            return Path.Combine(project.ProjectFolderPath, FolderName, FileName);
        }

        // Missing file: defaults are written there so users have a template to edit.
        // Unreadable file: defaults are used and the file is left untouched; 'message' explains why.
        public static WeldPropSettings Load(string path, out string message)
        {
            message = null;
            if (!File.Exists(path))
            {
                WeldPropSettings defaults = WeldPropSettings.CreateDefault();
                try
                {
                    Save(defaults, path);
                    message = "Created default settings: " + path;
                }
                catch (Exception ex)
                {
                    message = "Using default settings; could not create " + path + ": " + ex.Message;
                }
                return defaults;
            }

            try
            {
                WeldPropSettings settings;
                using (FileStream stream = File.OpenRead(path))
                {
                    settings = (WeldPropSettings)CreateSerializer().ReadObject(stream);
                }
                if (settings == null) throw new InvalidDataException("File is empty.");
                settings.Normalize();
                return settings;
            }
            catch (Exception ex)
            {
                message = "Using default settings; could not read " + path + ": " + ex.Message;
                return WeldPropSettings.CreateDefault();
            }
        }

        // Writes to a temporary file first, so a failed write never leaves a half-written settings file.
        public static void Save(WeldPropSettings settings, string path)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            string tempPath = path + ".tmp";
            using (FileStream stream = File.Create(tempPath))
            using (var writer = JsonReaderWriterFactory.CreateJsonWriter(stream, new UTF8Encoding(false), false, true, "  "))
            {
                CreateSerializer().WriteObject(writer, settings);
                writer.Flush();
            }
            if (File.Exists(path))
                File.Replace(tempPath, path, null);
            else
                File.Move(tempPath, path);
        }

        private static DataContractJsonSerializer CreateSerializer()
        {
            return new DataContractJsonSerializer(typeof(WeldPropSettings));
        }
    }
}
