namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.Unit
{
    using System.Collections.Generic;
    using System.Threading.Tasks;

    /// <summary>
    /// Interface for Unit data repository.
    /// </summary>
    public interface IUnitDataRepository
    {
        /// <summary>
        /// Get units by user id.
        /// </summary>
        /// <param name="userId">User Id.</param>
        /// <returns>Units.</returns>
        public Task<IEnumerable<UnitDataEntity>> GetByUserIdAsync(string userId);

        /// <summary>
        /// Get all units.
        /// </summary>
        /// <returns>Units.</returns>
        public Task<IEnumerable<UnitDataEntity>> GetAllAsync();

        /// <summary>
        /// Create or update unit data.
        /// </summary>
        /// <param name="unitDataEntity">Unit data entity.</param>
        /// <returns>unit data id.</returns>
        public Task<string> SaveAsync(UnitDataEntity unitDataEntity);

        /// <summary>
        /// Delete unit data.
        /// </summary>
        /// <param name="id">Unit data id.</param>
        /// <returns>True if deleted. False if not found.</returns>
        public Task<bool> TryDeleteAsync(string id);

        /// <summary>
        /// Get unit data by id.
        /// </summary>
        /// <param name="id">Unit data id.</param>
        /// <returns>Unit data.</returns>
        Task<UnitDataEntity> GetAsync(string id);
    }
}