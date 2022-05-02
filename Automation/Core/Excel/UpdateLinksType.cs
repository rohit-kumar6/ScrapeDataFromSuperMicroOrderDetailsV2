namespace Automation.Core.Excel
{
    /// <summary>
    /// Enum of UpdateLinks type.
    /// </summary>
    public enum UpdateLinksType
    {
        /// <summary>Prompted appears to user where he can specify how links will be updated.</summary>
        Default = -1,

        /// <summary>No Links will be updated.</summary>
        NoUpdates = 0,

        /// <summary>Only External Links will be updated.</summary>
        UpdateExternalLinks = 1,

        /// <summary>Only Remote Links will be updated.</summary>
        UpdateRemoteLinks = 2,

        /// <summary>Update All Links.</summary>
        UpdateAllLinks = 3,
    }
}
