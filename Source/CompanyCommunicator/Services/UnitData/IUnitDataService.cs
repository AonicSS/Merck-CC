using System.Threading.Tasks;

namespace Microsoft.Teams.Apps.CompanyCommunicator.Services.UnitData
{
    /// <summary>
    /// Unit Data Service.
    /// </summary>
    public interface IUnitDataService
    {
        /// <summary>
        /// Check if the user is in any unit.
        /// </summary>
        /// <param name="upn">The user's upn.</param>
        /// <returns>True if the user is in any unit, false otherwise.</returns>
        Task<bool> IsUserInAnyUnitAsync(string upn);
    }
}