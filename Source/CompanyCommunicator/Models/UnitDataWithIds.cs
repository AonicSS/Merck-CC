namespace Microsoft.Teams.Apps.CompanyCommunicator.Models
{
    /// <summary>
    /// Unit data model class.
    /// </summary>
    public class UnitDataWithIds : UnitData
    {
        /// <summary>
        /// Gets or sets user ids.
        /// </summary>
        public string[] UserIds { get; set; }

        /// <summary>
        /// Gets or sets group ids.
        /// </summary>
        public string[] GroupIds { get; set; }
    }
}