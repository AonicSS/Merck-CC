namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Teams.Apps.CompanyCommunicator.Authentication;
    using Microsoft.Teams.Apps.CompanyCommunicator.Models;

    /// <summary>
    /// Controller for the notifications API.
    /// </summary>
    [Route("api/notifications")]
    [Authorize(PolicyNames.MustBeValidUpnPolicy)]
    public class NotificationsController : ControllerBase
    {
        /// <summary>
        /// Gets draft notifications for a unit.
        /// </summary>
        /// <param name="unitId">Unit Id.</param>
        /// <returns>It returns the notifications for the passed unit id and notification type.
        /// The returning value is wrapped in a ActionResult object.
        /// If the passed in unit id is invalid, it returns 404 not found error.</returns>
        [HttpGet("draft/{unitId}")]
        public Task<ActionResult<BaseNotification>> GetDraftNotificationsByUnitIdAsync(string unitId)
        {
            throw new System.NotImplementedException();
        }

        /// <summary>
        /// Gets sent notifications for a unit.
        /// </summary>
        /// <param name="unitId">Unit Id.</param>
        /// <returns>It returns the notifications for the passed unit id and notification type.
        /// The returning value is wrapped in a ActionResult object.
        /// If the passed in unit id is invalid, it returns 404 not found error.</returns>
        [HttpGet("sent/{unitId}")]
        public Task<ActionResult<BaseNotification>> GetSentNotificationsByUnitIdAsync(string unitId)
        {
            throw new System.NotImplementedException();
        }
    }
}