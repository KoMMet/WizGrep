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
            return Loader.GetString(resourceKey);
        }
    }
}
