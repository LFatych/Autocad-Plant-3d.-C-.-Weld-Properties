using Autodesk.AutoCAD.Runtime;

[assembly: ExtensionApplication(typeof(WeldPropUtils.PluginApp))]

namespace WeldPropUtils
{
    // Runs when the DLL is loaded (NETLOAD or autoload). Keep it light: there may be no drawing or project yet.
    public class PluginApp : IExtensionApplication
    {
        public void Initialize()
        {
            try
            {
                WeldAutoUpdate.Start();
            }
            catch (System.Exception ex)
            {
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument?.Editor
                    .WriteMessage("\nWeldPropUtils: weld auto-update could not start: " + ex.Message);
            }
        }

        public void Terminate()
        {
        }
    }
}
