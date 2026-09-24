using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.ProcessPower.PlantInstance;
using Autodesk.ProcessPower.PnP3dObjects;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using WeldPropUtils.Settings;
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace WeldPropUtils
{
    // Fills the mapped properties of new welds automatically when the project settings have "autoUpdate": true.
    // Kept cheap on purpose:
    //  - Database.ObjectAppended only remembers the ObjectId of a new Connector (a type check, no database work);
    //  - when the command that created them ends, one Idle handler is attached; on the first Idle where AutoCAD is
    //    quiescent it processes only those connectors, then detaches again. No handler runs while nothing is pending.
    public static class WeldAutoUpdate
    {
        // Undo/redo re-create objects that already have their properties; our own commands write weld properties themselves.
        private static readonly HashSet<string> SkippedCommands = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "U", "UNDO", "REDO", "MREDO", "SETWELDPROP", "SETWELDNUMBER", "WELDPROPMAPPING", "WELDPROPAUTO"
        };

        // More new connectors than this at once (a big paste or insert): ask the user to run SetWeldProp instead.
        private const int MaxWeldsPerRun = 500;

        private static readonly Dictionary<Database, HashSet<ObjectId>> Pending = new Dictionary<Database, HashSet<ObjectId>>();
        private static bool _started;
        private static bool _idleAttached;
        private static bool _busy;

        // "autoUpdate" of the last settings file read, reused until the file changes.
        private static string _cachedPath;
        private static DateTime _cachedWriteTime;
        private static bool _cachedEnabled;

        public static void Start()
        {
            if (_started) return;
            _started = true;
            DocumentCollection docs = AcApp.DocumentManager;
            foreach (Document doc in docs) Attach(doc);
            docs.DocumentCreated += (sender, e) => Attach(e.Document);
            docs.DocumentToBeDestroyed += (sender, e) => Detach(e.Document);
        }

        private static void Attach(Document doc)
        {
            if (doc == null) return;
            doc.Database.ObjectAppended += OnObjectAppended;
            doc.CommandEnded += OnCommandFinished;
            doc.CommandCancelled += OnCommandFinished;   // e.g. routing ended with Esc: the routed parts stay
            doc.CommandFailed += OnCommandFinished;
        }

        private static void Detach(Document doc)
        {
            if (doc == null) return;
            doc.Database.ObjectAppended -= OnObjectAppended;
            doc.CommandEnded -= OnCommandFinished;
            doc.CommandCancelled -= OnCommandFinished;
            doc.CommandFailed -= OnCommandFinished;
            Pending.Remove(doc.Database);
        }

        // Runs for every object appended to a drawing: must stay trivial.
        private static void OnObjectAppended(object sender, ObjectEventArgs e)
        {
            if (_busy || !(e.DBObject is Connector)) return;
            Database db = e.DBObject.Database;
            if (db == null) return;
            if (!Pending.TryGetValue(db, out HashSet<ObjectId> ids))
                Pending[db] = ids = new HashSet<ObjectId>();
            ids.Add(e.DBObject.ObjectId);
        }

        private static void OnCommandFinished(object sender, CommandEventArgs e)
        {
            if (!(sender is Document doc) || !Pending.TryGetValue(doc.Database, out HashSet<ObjectId> ids) || ids.Count == 0) return;
            if (SkippedCommands.Contains(e.GlobalCommandName))
            {
                ids.Clear();
                return;
            }
            if (_idleAttached) return;
            AcApp.Idle += OnIdle;
            _idleAttached = true;
        }

        private static void OnIdle(object sender, EventArgs e)
        {
            if (_busy) return;
            Document doc = AcApp.DocumentManager.MdiActiveDocument;
            if (doc == null || !doc.Editor.IsQuiescent) return;   // still inside a command (e.g. picking points): wait
            if (!Pending.TryGetValue(doc.Database, out HashSet<ObjectId> pending) || pending.Count == 0)
            {
                // Work left only in other drawings waits until one of them is active.
                if (Pending.Values.All(p => p.Count == 0)) DetachIdle();
                return;
            }

            ObjectId[] ids = pending.ToArray();
            pending.Clear();
            if (Pending.Values.All(p => p.Count == 0)) DetachIdle();

            _busy = true;
            try
            {
                using (doc.LockDocument())
                {
                    Process(doc, ids);
                }
            }
            catch (System.Exception ex)
            {
                doc.Editor.WriteMessage("\nWeld auto-update failed: " + ex.Message + "\n");
            }
            finally
            {
                _busy = false;
            }
        }

        private static void DetachIdle()
        {
            AcApp.Idle -= OnIdle;
            _idleAttached = false;
        }

        private static void Process(Document doc, ObjectId[] ids)
        {
            if (PlantApplication.CurrentProject == null) return;
            string path = SettingsStore.GetSettingsPath();
            if (path == null || !IsEnabled(path)) return;

            ids = ids.Where(id => id.IsValid && !id.IsErased).ToArray();
            if (ids.Length == 0) return;
            if (ids.Length > MaxWeldsPerRun)
            {
                doc.Editor.WriteMessage("\n" + ids.Length + " new connectors: run SetWeldProp to fill the weld properties "
                    + "(auto-update handles up to " + MaxWeldsPerRun + " at once).\n");
                return;
            }

            WeldPropSettings settings = SettingsStore.Load(path, out string message);
            if (message != null) doc.Editor.WriteMessage("\n" + message);
            MappingProfile profile = WeldPropertiesHandler.WritableProfile(settings.GetActiveProfile());
            if (profile.Mappings.Count == 0) return;

            int count = WeldPropertiesHandler.ProcessWelds(ids, (connector, weld) => WeldPropertiesHandler.SetWeldProp(connector, weld, profile));
            if (count > 0) doc.Editor.WriteMessage("\nWeld properties filled for " + count + " new weld(s).\n");
        }

        // Reads "autoUpdate" only when the settings file changed; a missing file means off (no file is created here).
        private static bool IsEnabled(string path)
        {
            if (!File.Exists(path)) return false;
            DateTime writeTime = File.GetLastWriteTimeUtc(path);
            if (path == _cachedPath && writeTime == _cachedWriteTime) return _cachedEnabled;
            WeldPropSettings settings = SettingsStore.Load(path, out string message);
            _cachedPath = path;
            _cachedWriteTime = writeTime;
            _cachedEnabled = message == null && settings.AutoUpdate;
            return _cachedEnabled;
        }
    }
}
