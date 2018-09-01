using Syncfusion.Licensing;

namespace QuickLook.Plugin.OfficeViewer
{
    internal static class SyncfusionKey
    {
        public static void Register()
        {
            SyncfusionLicenseProvider.RegisterLicense("YOUR-LICENSE-KEY");
        }
    }
}