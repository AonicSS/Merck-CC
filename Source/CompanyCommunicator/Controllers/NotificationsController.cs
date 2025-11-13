namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System.Collections.Generic;
    using System.Linq;
    using System.Security.Claims;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Teams.Apps.CompanyCommunicator.Authentication;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Extensions;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.ExportData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.TeamData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph;
    using Microsoft.Teams.Apps.CompanyCommunicator.Models;

    /// <summary>
    /// Controller for the notifications API.
    /// </summary>
    [Route("api/notifications")]
    [Authorize(PolicyNames.MustBeValidUpnPolicy)]
    public class NotificationsController : ControllerBase
    {
        private readonly INotificationDataRepository notificationDataRepository;
        private readonly IGroupsService groupsService;
        private readonly IExportDataRepository exportDataRepository;
        private readonly ITeamDataRepository teamDataRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="NotificationsController"/> class.
        /// </summary>
        /// <param name="notificationDataRepository">Notification data repository instance.</param>
        /// <param name="groupsService">Groups service.</param>
        /// <param name="exportDataRepository">Export data repository instance.</param>
        /// <param name="teamDataRepository">Team data repository instance.</param>
        public NotificationsController(
            INotificationDataRepository notificationDataRepository,
            IGroupsService groupsService,
            IExportDataRepository exportDataRepository,
            ITeamDataRepository teamDataRepository)
        {
            this.notificationDataRepository = notificationDataRepository;
            this.groupsService = groupsService;
            this.exportDataRepository = exportDataRepository;
            this.teamDataRepository = teamDataRepository;
        }

        /// <summary>
        /// Gets draft notifications for a unit.
        /// </summary>
        /// <param name="unitId">Unit Id.</param>
        /// <returns>Draft notifications.</returns>
        [HttpGet("draft/{unitId}")]
        public async Task<IEnumerable<DraftNotification>> GetDraftNotificationsByUnitIdAsync(string unitId)
        {
            var entities = await this.notificationDataRepository.GetAllDraftNotificationsOfUnitAsync(unitId);
            var results = new List<DraftNotification>();

            foreach (var entity in entities)
            {
                var groupNames = await this.groupsService
                    .GetByIdsAsync(entity.Groups)
                    .Select(x => x.DisplayName)
                    .ToListAsync();

                var result = new DraftNotification
                {
                    Id = entity.Id,
                    Title = entity.Title,
                    ImageLink = entity.ImageLink,
                    ImageBase64BlobName = entity.ImageBase64BlobName,
                    Summary = entity.Summary,
                    Author = entity.Author,
                    ButtonTitle = entity.ButtonTitle,
                    ButtonLink = entity.ButtonLink,
                    CreatedDateTime = entity.CreatedDate,
                    Teams = entity.Teams,
                    Rosters = entity.Rosters,
                    Groups = entity.Groups,
                    AllUsers = entity.AllUsers,
                    UnitId = entity.UnitId,
                    UnitName = entity.UnitName,
                    CreatedBy = entity.CreatedBy,
                    GroupNames = groupNames,
                };

                // In case we have blob name instead of URL to public image.
                if (!string.IsNullOrEmpty(entity.ImageBase64BlobName))
                {
                    result.ImageLink = await this.notificationDataRepository.GetImageAsync(entity.ImageLink, entity.ImageBase64BlobName);
                }

                results.Add(result);
            }

            return results;
        }

        /// <summary>
        /// Gets scheduled notifications for a unit.
        /// </summary>
        /// <param name="unitId">Unit Id.</param>
        /// <returns>Scheduled Draft notifications.</returns>
        [HttpGet("scheduled/{unitId}")]
        public async Task<IEnumerable<DraftNotification>> GetScheduledNotificationsByUnitIdAsync(string unitId)
        {
            var entities = await this.notificationDataRepository.GetAllScheduledNotificationsOfUnitAsync(unitId);
            var results = new List<DraftNotification>();

            foreach (var entity in entities)
            {
                var groupNames = await this.groupsService
                    .GetByIdsAsync(entity.Groups)
                    .Select(x => x.DisplayName)
                    .ToListAsync();

                var result = new DraftNotification
                {
                    Id = entity.Id,
                    Title = entity.Title,
                    ImageLink = entity.ImageLink,
                    ImageBase64BlobName = entity.ImageBase64BlobName,
                    Summary = entity.Summary,
                    Author = entity.Author,
                    ButtonTitle = entity.ButtonTitle,
                    ButtonLink = entity.ButtonLink,
                    CreatedDateTime = entity.CreatedDate,
                    Teams = entity.Teams,
                    Rosters = entity.Rosters,
                    Groups = entity.Groups,
                    AllUsers = entity.AllUsers,
                    UnitId = entity.UnitId,
                    UnitName = entity.UnitName,
                    CreatedBy = entity.CreatedBy,
                    GroupNames = groupNames,
                };

                // In case we have blob name instead of URL to public image.
                if (!string.IsNullOrEmpty(entity.ImageBase64BlobName))
                {
                    result.ImageLink = await this.notificationDataRepository.GetImageAsync(entity.ImageLink, entity.ImageBase64BlobName);
                }

                results.Add(result);
            }

            return results;
        }

        /// <summary>
        /// Gets sent notifications for a unit.
        /// </summary>
        /// <param name="unitId">Unit Id.</param>
        /// <returns>Sent notifications.</returns>
        [HttpGet("sent/{unitId}")]
        public async Task<IEnumerable<SentNotification>> GetSentNotificationsByUnitIdAsync(string unitId)
        {
            var entities = await this.notificationDataRepository.GetAllSentNotificationsOfUnitAsync(unitId);
            var results = new List<SentNotification>();

            foreach (var entity in entities)
            {
                var groupNames = await this.groupsService
                    .GetByIdsAsync(entity.Groups)
                    .Select(x => x.DisplayName)
                    .ToListAsync();

                var userId = this.HttpContext.User.FindFirstValue(Common.Constants.ClaimTypeUserId);
                var userNotificationDownload = await this.exportDataRepository.GetAsync(userId, entity.Id);

                var result = new SentNotification
                {
                    Id = entity.Id,
                    Title = entity.Title,
                    ImageLink = entity.ImageLink,
                    ImageBase64BlobName = entity.ImageBase64BlobName,
                    Summary = entity.Summary,
                    Author = entity.Author,
                    ButtonTitle = entity.ButtonTitle,
                    ButtonLink = entity.ButtonLink,
                    CreatedDateTime = entity.CreatedDate,
                    SentDate = entity.SentDate,
                    Succeeded = entity.Succeeded,
                    Failed = entity.Failed,
                    Unknown = GetUnknownCount(entity),
                    Canceled = entity.Canceled > 0 ? entity.Canceled : null,
                    TeamNames = await this.teamDataRepository.GetTeamNamesByIdsAsync(entity.Teams),
                    RosterNames = await this.teamDataRepository.GetTeamNamesByIdsAsync(entity.Rosters),
                    GroupNames = groupNames,
                    AllUsers = entity.AllUsers,
                    SendingStartedDate = entity.SendingStartedDate,
                    ErrorMessage = entity.ErrorMessage,
                    WarningMessage = entity.WarningMessage,
                    CanDownload = userNotificationDownload == null,
                    SendingCompleted = entity.IsCompleted(),
                    CreatedBy = entity.CreatedBy,
                    UnitId = entity.UnitId,
                    UnitName = entity.UnitName,
                };

                // In case we have blob name instead of URL to public image.
                if (!string.IsNullOrEmpty(entity.ImageBase64BlobName)
                    && result.ImageLink.StartsWith(Common.Constants.ImageBase64Format))
                {
                    result.ImageLink =
                        await this.notificationDataRepository.GetImageAsync(result.ImageLink,
                            entity.ImageBase64BlobName);
                }

                results.Add(result);
            }

            return results;
        }

        private static int? GetUnknownCount(NotificationDataEntity notificationEntity)
        {
            var unknown = notificationEntity.Unknown;

            // In CC v2, the number of throttled recipients are counted and saved in NotificationDataEntity.Unknown property.
            // However, CC v1 saved the number of throttled recipients in NotificationDataEntity.Throttled property.
            // In order to make it backward compatible, we add the throttled number to the unknown variable.
            var throttled = notificationEntity.Throttled;
            if (throttled > 0)
            {
                unknown += throttled;
            }

            return unknown > 0 ? unknown : (int?)null;
        }
    }
}