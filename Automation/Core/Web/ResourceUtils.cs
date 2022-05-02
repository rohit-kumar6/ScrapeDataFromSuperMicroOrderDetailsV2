using System;
using System.IO;
using System.Reflection;

namespace Automation.Core.Web
{

    /// <summary>
    /// Utility file for handling resource file operations.
    /// </summary>
    public static class ResourceUtils
    {
        /// <summary>
        /// Get resource file content.
        /// </summary>
        /// <param name="resourceType">Type of resx resource file.</param>
        /// <param name="resourceName">Name of file inside resx.</param>
        /// <returns>Content as string.</returns>
        public static string GetResourceFileContent(Type resourceType, string resourceName)
        {
            if ((resourceType != null) && (resourceName != null))
            {
                PropertyInfo property = resourceType.GetProperty(resourceName, BindingFlags.Public | BindingFlags.Static | BindingFlags.NonPublic);

                if (property == null)
                {
                    throw new InvalidOperationException(string.Format("Resource Type Does Not Have Property"));
                }

                if (property.PropertyType != typeof(string))
                {
                    throw new InvalidOperationException(string.Format("Resource Property is Not String Type"));
                }

                return (string)property.GetValue(null, null);
            }

            return null;
        }

        /// <summary>
        /// Fetches file content from resource files of executing assembly.
        /// </summary>
        /// <param name="resourceName">Name of the file inside resources.</param>
        /// <returns>Resource File content.</returns>
        public static string GetFileContentFromResources(string resourceName)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                using (StreamReader reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }
    }
}
