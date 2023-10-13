namespace Microsoft.Teams.Apps.CompanyCommunicator.Models
{
    /// <summary>
    /// Unit data model class.
    /// </summary>
    public class UnitDataWithDetails : UnitData
    {
        /// <summary>
        /// Gets or sets unit users.
        /// </summary>
        public UserData[] Users { get; set; }

        /// <summary>
        /// Gets or sets unit groups.
        /// </summary>
        public GroupData[] Groups { get; set; }
    }
}