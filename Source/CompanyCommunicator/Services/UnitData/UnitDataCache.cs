namespace Microsoft.Teams.Apps.CompanyCommunicator.Services
{
    using System;
    using System.Collections.Concurrent;
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// Unit Data Cache.
    /// </summary>
    public class UnitDataCache
    {
        private DateTime lastCleanUp = DateTime.Now;

        private readonly ConcurrentDictionary<string, UnitDataCacheItem> cache =
            new ConcurrentDictionary<string, UnitDataCacheItem>();

        /// <summary>
        /// Try to get the value from the cache.
        /// </summary>
        /// <param name="upn">upn.</param>
        /// <param name="isIncluded">isIncluded.</param>
        /// <returns>True if the value is in the cache, false otherwise.</returns>
        public bool TryGet(string upn, out bool isIncluded)
        {
            isIncluded = false;
            if (!this.cache.TryGetValue(upn, out var item))
            {
                return false;
            }

            if (item.ValidUntil < DateTime.Now)
            {
                this.cache.Remove(upn, out _);
                return true;
            }

            isIncluded = item.IsIncluded;
            return true;
        }

        /// <summary>
        /// Put the value into the cache.
        /// </summary>
        /// <param name="upn">upn.</param>
        /// <param name="isIncluded">isIncluded.</param>
        public void Put(string upn, bool isIncluded)
        {
            this.CleanUpEveryFiveMinutes();
            var item = new UnitDataCacheItem(DateTime.Now.AddMinutes(5), isIncluded);
            this.cache[upn] = item;
        }

        private void CleanUpEveryFiveMinutes()
        {
            if (DateTime.Now - this.lastCleanUp <= TimeSpan.FromMinutes(5))
            {
                return;
            }

            this.lastCleanUp = DateTime.Now;
            var keysToRemove = this.cache.Where(c => c.Value.ValidUntil < DateTime.Now).Select(c => c.Key);
            foreach (var key in keysToRemove)
            {
                this.cache.Remove(key, out _);
            }
        }
    }
}