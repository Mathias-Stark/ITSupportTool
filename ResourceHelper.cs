using System.Resources;

namespace ITSupportToolGUI
{
    public static class ResourceHelper
    {
        private static readonly ResourceManager rm = new ResourceManager("ITSupportToolGUI.Resources.Strings", typeof(ResourceHelper).Assembly);

        public static string GetString(string key)
        {
            return rm.GetString(key);
        }
    }
}