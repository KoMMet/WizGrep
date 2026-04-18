using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Windows.ApplicationModel.Resources;

namespace WizGrep.Helpers
{
    public class ResourceLoaderHelper
    {
        private static readonly ResourceLoader Loader = new();

        public static string GetString(string resourceKey)
        {
            try
            {
                return Loader.GetString(resourceKey);
            }
            catch (Exception e)
            {
                LoggerHelper.Instance.LogError($"Error loading resource string for key '{resourceKey}': {e.StackTrace}");
            }
            return "";
        }
    }
}
