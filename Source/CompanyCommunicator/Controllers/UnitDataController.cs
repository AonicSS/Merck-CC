using System.Collections.Generic;

namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Teams.Apps.CompanyCommunicator.Authentication;
    using Microsoft.Teams.Apps.CompanyCommunicator.Models;

    /// <summary>
    /// Controller for the unit data data API.
    /// </summary>
    [Route("api/unitData")]
    [Authorize(PolicyNames.MustBeValidUpnPolicy)]
    public class UnitDataController : ControllerBase
    {
        /// <summary>
        /// Create a new unit data.
        /// </summary>
        /// <param name="unitData">The unit data to be created.</param>
        /// <returns>The unit data id.</returns>
        [HttpPost]
        public Task<string> CreateUnitDataAsync([FromBody] UnitDataWithIds unitData)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Get a unit data.
        /// </summary>
        /// <param name="unitData">The unit data to be created.</param>
        /// <returns>The unit data id.</returns>
        public Task<string> UpdateUnitDataAsync([FromBody] UnitDataWithIds unitData)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Delete an existing unit data.
        /// </summary>
        /// <param name="id">The id of the unit data to be deleted.</param>
        /// <returns>If the passed in Id is invalid, it returns 404 not found error. Otherwise, it returns 200 OK.</returns>
        [HttpDelete("{id}")]
        public Task<IActionResult> DeleteUnitDataAsync(string id)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Get all unit data.
        /// </summary>
        /// <returns>A list of <see cref="UnitData"/> instances.</returns>
        [HttpGet]
        public Task<ActionResult<IEnumerable<UnitData>>> GetAllUnitDataAsync()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Get a unit data by id. 
        /// </summary>
        /// <param name="id">Unit data Id.</param>
        /// <returns>It returns the unit data with the passed in id.
        /// The returning value is wrapped in a ActionResult object.
        /// If the passed in id is invalid, it returns 404 not found error.</returns>
        [HttpGet("{id}")]
        public Task<ActionResult<UnitDataWithDetails>> GetUnitDataByIdAsync(string id)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Get a unit data by user id.
        /// </summary>
        /// <param name="userId">User Id.</param>
        /// <returns>It returns the unit data with the passed in user id.
        /// The returning value is wrapped in a ActionResult object.
        /// If the passed in user id is invalid, it returns 404 not found error.</returns>
        [HttpGet("user/{userId}")]
        public Task<ActionResult<ActionResult<IEnumerable<UnitData>>>> GetUnitDataByUserIdAsync(string userId)
        {
            throw new NotImplementedException();
        }
    }
}