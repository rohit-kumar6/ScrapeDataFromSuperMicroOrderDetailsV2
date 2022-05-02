namespace Automation.Utils
{
    using System;
    using Argument.Check;

    /// <summary>
    /// Utilities related to string operations.
    /// </summary>
    public static class StringUtils
    {
        /// <summary>
        /// Replaces first occurence of a string in a text.
        /// </summary>
        /// <param name="text">Original string.</param>
        /// <param name="search">String to replace.</param>
        /// <param name="replace">String to replace with.</param>
        /// <returns>Modified string.</returns>
        public static string ReplaceFirst(string text, string search, string replace)
        {
            if (string.IsNullOrEmpty(text))
            {
                throw new ArgumentException(string.Format("{0} cannot be null", "Original Text"));
            }

            if (string.IsNullOrEmpty(text))
            {
                throw new ArgumentException(string.Format("{0} cannot be null", "String to search"));
            }

            int pos = text.IndexOf(search);
            if (pos < 0)
            {
                return text;
            }

            return text.Substring(0, pos) + replace + text.Substring(pos + search.Length);
        }

        /// <summary>
        /// Append underscore as suffix to parameter.
        /// </summary>
        /// <param name="parameter">String value to be appended with underscore.</param>
        /// <returns>Modified prefix.</returns>
        public static string GetRenamePrefix(string parameter)
        {
            return string.IsNullOrEmpty(parameter) ? string.Empty : $"{parameter}_";
        }

        /// <summary>
        /// Method to check if string contains given value by ignoring case sensitivity.
        /// </summary>
        /// <param name="text">Input string.</param>
        /// <param name="value">Value to check.</param>
        /// <returns>Return true/false.</returns>
        public static bool CaseInsensitiveContains(string text, string value)
        {
            Throw.IfNull(() => text);
            StringComparison stringComparison = StringComparison.CurrentCultureIgnoreCase;
            return text.IndexOf(value, stringComparison) >= 0;
        }
    }
}
