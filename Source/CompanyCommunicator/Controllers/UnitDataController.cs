namespace Microsoft.Teams.Apps.CompanyCommunicator.Controllers
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.AspNetCore.Authorization;
    using Microsoft.AspNetCore.Mvc;
    using Microsoft.Teams.Apps.CompanyCommunicator.Authentication;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.Unit;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph;
    using Microsoft.Teams.Apps.CompanyCommunicator.Models;

    /// <summary>
    /// Controller for the unit data data API.
    /// </summary>
    [Route("api/unitData")]
    [Authorize(PolicyNames.MustBeValidUpnPolicy)]
    public class UnitDataController : ControllerBase
    {
        private readonly IUnitDataRepository unitDataRepository;
        private readonly IGroupsService groupsService;
        private readonly IUsersService usersService;

        /// <summary>
        /// Initializes a new instance of the <see cref="UnitDataController"/> class.
        /// </summary>
        /// <param name="unitDataRepository">Unit data repository.</param>
        /// <param name="groupsService">Microsoft Graph groups service instance.</param>
        /// <param name="usersService">Microsoft Graph users service instance.</param>
        public UnitDataController(IUnitDataRepository unitDataRepository, IGroupsService groupsService, IUsersService usersService)
        {
            this.unitDataRepository = unitDataRepository;
            this.groupsService = groupsService;
            this.usersService = usersService;
        }

        /// <summary>
        /// Create a new unit data.
        /// </summary>
        /// <param name="unitData">The unit data to be created.</param>
        /// <returns>The unit data id.</returns>
        [HttpPost]
        public async Task<string> CreateUnitDataAsync([FromBody] UnitDataWithIds unitData)
        {
            var unitDataEntity = new UnitDataEntity
            {
                RowKey = null,
                Name = unitData.Name,
                UserIds = string.Join(",", unitData.UserIds ?? Array.Empty<string>()),
                GroupIds = string.Join(",", unitData.GroupIds ?? Array.Empty<string>()),
            };
            await this.unitDataRepository.SaveAsync(unitDataEntity);
            return unitDataEntity.RowKey;
        }

        /// <summary>
        /// Get a unit data.
        /// </summary>
        /// <param name="unitData">The unit data to be created.</param>
        /// <returns>The unit data id.</returns>
        public async Task<string> UpdateUnitDataAsync([FromBody] UnitDataWithIds unitData)
        {
            var unitDataEntity = new UnitDataEntity
            {
                RowKey = unitData.Id,
                Name = unitData.Name,
                UserIds = string.Join(",", unitData.UserIds ?? Array.Empty<string>()),
                GroupIds = string.Join(",", unitData.GroupIds ?? Array.Empty<string>()),
            };
            await this.unitDataRepository.SaveAsync(unitDataEntity);
            return unitDataEntity.RowKey;
        }

        /// <summary>
        /// Delete an existing unit data.
        /// </summary>
        /// <param name="id">The id of the unit data to be deleted.</param>
        /// <returns>If the passed in Id is invalid, it returns 404 not found error. Otherwise, it returns 200 OK.</returns>
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteUnitDataAsync(string id)
        {
            var success = await this.unitDataRepository.TryDeleteAsync(id);
            return success ? (IActionResult)this.Ok() : this.NotFound();
        }

        /// <summary>
        /// Get all unit data.
        /// </summary>
        /// <returns>A list of <see cref="UnitData"/> instances.</returns>
        [HttpGet]
        public async Task<ActionResult<IEnumerable<UnitData>>> GetAllUnitDataAsync()
        {
            var unitDataEntities = await this.unitDataRepository.GetAllAsync();
            var results = await this.GetUnitsAsync(unitDataEntities);
            return this.Ok(results);
        }

        /// <summary>
        /// Get a unit data by id.
        /// </summary>
        /// <param name="id">Unit data Id.</param>
        /// <returns>It returns the unit data with the passed in id.
        /// The returning value is wrapped in a ActionResult object.
        /// If the passed in id is invalid, it returns 404 not found error.</returns>
        [HttpGet("{id}")]
        public async Task<ActionResult<UnitDataWithDetails>> GetUnitDataByIdAsync(string id)
        {
            var unitDataEntity = await this.unitDataRepository.GetAsync(id);
            if (unitDataEntity == null)
            {
                return this.NotFound();
            }

            var unitData = new UnitDataWithDetails
            {
                Id = unitDataEntity.RowKey,
                Name = unitDataEntity.Name,
                Users = Array.Empty<UserData>(),
                Groups = Array.Empty<GroupData>(),
            };

            if (unitDataEntity.UserIds != null && unitDataEntity.UserIds.Any())
            {
                var userIds = unitDataEntity.UserIds.Split(",").Select(e => e.Trim()).ToArray();
                var groupedUserIds = userIds
                    .Select((s, i) => new { Value = s, Index = i })
                    .GroupBy(x => x.Index / 15)
                    .Select(grp => grp.Select(x => x.Value).ToArray())
                    .ToArray();

                var users = await this.usersService.GetBatchByUserIds(groupedUserIds);
                unitData.Users = users.Select(u => new UserData
                {
                    Id = u.Id,
                    Name = u.DisplayName,
                    Mail = u.Mail,
                }).ToArray();
            }

            unitData.Size = unitData.Users.Length;
            if (unitDataEntity.GroupIds != null && unitDataEntity.GroupIds.Any())
            {
                var groupIds = unitDataEntity.GroupIds.Split(",").Select(e => e.Trim());
                var groups = await this.groupsService.GetByIdsAsync(groupIds).Select(g => new
                {
                    Size = g.Members.Count,
                    Group = new GroupData
                    {
                        Id = g.Id,
                        Name = g.DisplayName,
                        Mail = g.Mail,
                    },
                }).ToArrayAsync();
                unitData.Size += groups.Sum(g => g.Size);
                unitData.Groups = groups.Select(g => g.Group).ToArray();
            }

            return this.Ok(unitData);
        }

        /// <summary>
        /// Get a unit data by user id.
        /// </summary>
        /// <param name="userId">User Id.</param>
        /// <returns>It returns the unit data with the passed in user id.
        /// The returning value is wrapped in a ActionResult object.
        /// If the passed in user id is invalid, it returns 404 not found error.</returns>
        [HttpGet("user/{userId}")]
        public async Task<ActionResult<ActionResult<IEnumerable<UnitData>>>> GetUnitDataByUserIdAsync(string userId)
        {
            var unitDataEntities = await this.unitDataRepository.GetByUserIdAsync(userId);

            var results = this.GetUnitsAsync(unitDataEntities);
            return this.Ok(results);
        }

        private async Task<List<UnitData>> GetUnitsAsync(IEnumerable<UnitDataEntity> unitDataEntities)
        {
            var results = new List<UnitData>();
            foreach (var entity in unitDataEntities)
            {
                var unitData = new UnitData
                {
                    Id = entity.RowKey,
                    Name = entity.Name,
                    Size = (entity.UserIds ?? string.Empty).Split(",").Length,
                };

                if (entity.GroupIds != null && entity.GroupIds.Any())
                {
                    var groupIds = entity.GroupIds.Split(",").Select(id => id.Trim());
                    var groupCounts = await this.groupsService.GetByIdsAsync(groupIds).Select(g => g.Members.Count)
                        .ToListAsync();
                    groupCounts.ForEach(gc => unitData.Size += gc);
                }

                results.Add(unitData);
            }

            return results;
        }
    }
}