namespace Microsoft.Teams.Apps.CompanyCommunicator.Services
{
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.Unit;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.MicrosoftGraph;
    using Microsoft.Teams.Apps.CompanyCommunicator.Services.UnitData;

    /// <inheritdoc />
    public class UnitDataService : IUnitDataService
    {
        private static readonly UnitDataCache Cache = new UnitDataCache();

        private readonly IUnitDataRepository unitDataRepository;
        private readonly IGroupsService groupsService;
        private readonly IUsersService usersService;

        /// <summary>
        /// Initializes a new instance of the <see cref="UnitDataService"/> class.
        /// </summary>
        /// <param name="unitDataRepository">repository.</param>
        /// <param name="groupsService">groups service.</param>
        /// <param name="usersService">users service.</param>
        public UnitDataService(IUnitDataRepository unitDataRepository, IGroupsService groupsService, IUsersService usersService)
        {
            this.unitDataRepository = unitDataRepository;
            this.groupsService = groupsService;
            this.usersService = usersService;
        }

        /// <inheritdoc />
        public async Task<bool> IsUserInAnyUnitAsync(string upn)
        {
            if (Cache.TryGet(upn, out var isIncluded))
            {
                return isIncluded;
            }

            var user = await this.usersService.GetUserByUpnAsync(upn);
            var groupIds = await this.usersService.GetUsersGroupIdsAsync(user.Id);
            var unitDataEntities = await this.unitDataRepository.GetAllAsync();

            isIncluded = unitDataEntities.Any(u =>
                u.UserIds.Contains(user.Id) || groupIds.Any(g => u.GroupIds.Contains(g)));
            Cache.Put(upn, isIncluded);
            return isIncluded;
        }
    }
}