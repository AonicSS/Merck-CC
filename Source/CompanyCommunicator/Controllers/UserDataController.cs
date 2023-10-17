namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Teams.Apps.CompanyCommunicator.Authentication;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph;
    using Microsoft.Teams.Apps.CompanyCommunicator.Models;

    /// <summary>
    /// Controller for getting users.
    /// </summary>
    [Route("api/userData")]
    [Authorize(PolicyNames.MustBeValidUpnPolicy)]
    public class UserDataController : ControllerBase
    {
        private readonly IUsersService usersService;

        /// <summary>
        /// Initializes a new instance of the <see cref="UserDataController"/> class.
        /// </summary>
        /// <param name="usersService">Microsoft Graph users service instance.</param>
        public UserDataController(IUsersService usersService)
        {
            this.usersService = usersService;
        }

        /// <summary>
        /// Action method to get users.
        /// </summary>
        /// <param name="filter">filter (see graph api).</param>
        /// <returns>list of user data.</returns>
        [HttpGet("search/{filter}")]
        public async Task<IEnumerable<UserData>> SearchAsync(string filter)
        {
            var usersInBatches = await this.usersService.GetUsersAsync(filter).ToListAsync();
            var users = usersInBatches.SelectMany(u => u);
            var results = users.Select(u => new UserData
            {
                Id = u.Id,
                Name = u.DisplayName,
                Mail = u.Mail,
            });
            return results;
        }

        /// <summary>
        /// Action method to get user.
        /// </summary>
        /// <param name="mail">user mail.</param>
        /// <returns>user data.</returns>
        [HttpGet("{mail:required}")]
        public async Task<UserData> GetUserAsync(string mail)
        {
            var user = await this.usersService.GetUserByMailAsync(mail);
            return new UserData
            {
                Id = user.Id,
                Name = user.DisplayName,
                Mail = user.Mail,
            };
        }
    }
}