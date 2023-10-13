namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Teams.Apps.CompanyCommunicator.Models;

    /// <summary>
    /// Controller for getting users.
    /// </summary>
    [Route("api/userData")]
    public class UserDataController : ControllerBase
    {
        /// <summary>
        /// Action method to get users.
        /// </summary>
        /// <param name="query">filter (see graph api).</param>
        /// <returns>list of user data.</returns>
        [HttpGet("search/{query}")]
        public Task<IEnumerable<UserData>> SearchAsync(string query)
        {
            throw new NotImplementedException();
        }


        /// <summary>
        /// Action method to get user.
        /// </summary>
        /// <param name="mail">user mail.</param>
        /// <returns>user data.</returns>
        [HttpGet("{mail:required}")]
        public Task<UserData> GetUserAsync(string mail)
        {
            throw new NotImplementedException();
        }
    }
}