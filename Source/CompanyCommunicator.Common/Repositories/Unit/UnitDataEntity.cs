namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.Unit
{
    using Microsoft.Azure.Cosmos.Table;

    /// <summary>
    /// Unit data entity class.
    /// </summary>
    public class UnitDataEntity : TableEntity
    {
        /// <summary>
        /// Gets or sets the unit's name.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the unit's user ids separated by comma.
        /// </summary>
        public string UserIds { get; set; }

        /// <summary>
        /// Gets or sets the unit's group ids separated by comma.
        /// </summary>
        public string GroupIds { get; set; }
    }
}
