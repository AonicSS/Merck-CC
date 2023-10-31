namespace Microsoft.Teams.Apps.CompanyCommunicator.Services
{
    using System;

    /// <summary>
    /// Unit Data Cache Item.
    /// </summary>
    public class UnitDataCacheItem
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="UnitDataCacheItem"/> class.
        /// </summary>
        /// <param name="validUntil">valid until.</param>
        /// <param name="isIncluded">is included.</param>
        public UnitDataCacheItem(DateTime validUntil, bool isIncluded)
        {
            this.ValidUntil = validUntil;
            this.IsIncluded = isIncluded;
        }

        /// <summary>
        /// Gets valid until.
        /// </summary>
        public DateTime ValidUntil { get; private set; }

        /// <summary>
        /// Gets a value indicating whether is included.
        /// </summary>
        public bool IsIncluded { get; private set; }
    }
}