namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.Unit
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Logging;
    using Microsoft.Extensions.Options;

    /// <summary>
    /// Repository of the unit data in the table storage.
    /// </summary>
    public class UnitDataRepository : BaseRepository<UnitDataEntity>, IUnitDataRepository
    {
        private readonly TableRowKeyGenerator tableRowKeyGenerator;
        private readonly string partitionKey;

        /// <summary>
        /// Initializes a new instance of the <see cref="UnitDataRepository"/> class.
        /// </summary>
        /// <param name="logger">Logger.</param>
        /// <param name="repositoryOptions">Options.</param>
        /// <param name="tableRowKeyGenerator">Table row key generator.</param>
        public UnitDataRepository(
            ILogger<UnitDataRepository> logger,
            IOptions<RepositoryOptions> repositoryOptions,
            TableRowKeyGenerator tableRowKeyGenerator)
            : base(
                  logger,
                  storageAccountConnectionString: repositoryOptions.Value.StorageAccountConnectionString,
                  tableName: UnitDataTableNames.TableName,
                  defaultPartitionKey: UnitDataTableNames.UnitDataPartition,
                  ensureTableExists: repositoryOptions.Value.EnsureTableExists)
        {
            this.tableRowKeyGenerator = tableRowKeyGenerator ?? throw new ArgumentNullException(nameof(tableRowKeyGenerator));
            this.partitionKey = UnitDataTableNames.UnitDataPartition;
        }

        /// <inheritdoc />
        public async Task<UnitDataEntity> GetAsync(string id)
        {
            await this.EnsureUnitDataTableExistsAsync();
            return await this.GetAsync(this.partitionKey, id);
        }

        /// <inheritdoc />
        public async Task<IEnumerable<UnitDataEntity>> GetByUserIdAsync(string userId)
        {
            await this.EnsureUnitDataTableExistsAsync();
            var units = await this.GetAllAsync(this.partitionKey);
            return units.Where(u => u.UserIds.Contains(userId));
        }

        /// <inheritdoc />
        public async Task<IEnumerable<UnitDataEntity>> GetAllAsync()
        {
            await this.EnsureUnitDataTableExistsAsync();
            return await this.GetAllAsync(this.partitionKey);
        }

        /// <inheritdoc />
        public async Task<string> SaveAsync(UnitDataEntity unitDataEntity)
        {
            await this.EnsureUnitDataTableExistsAsync();
            unitDataEntity.PartitionKey = this.partitionKey;
            if (string.IsNullOrEmpty(unitDataEntity.RowKey))
            {
                unitDataEntity.RowKey = this.tableRowKeyGenerator.CreateNewKeyOrderingMostRecentToOldest();
            }

            await this.CreateOrUpdateAsync(unitDataEntity);
            return unitDataEntity.RowKey;
        }

        /// <inheritdoc />
        public async Task<bool> TryDeleteAsync(string id)
        {
            await this.EnsureUnitDataTableExistsAsync();
            var unitDataEntity = await this.GetAsync(this.partitionKey, id);
            if (unitDataEntity == null)
            {
                return false;
            }

            await this.DeleteAsync(unitDataEntity);
            return true;
        }

        private async Task EnsureUnitDataTableExistsAsync()
        {
            var exists = await this.Table.ExistsAsync();
            if (!exists)
            {
                await this.Table.CreateAsync();
            }
        }
    }
}