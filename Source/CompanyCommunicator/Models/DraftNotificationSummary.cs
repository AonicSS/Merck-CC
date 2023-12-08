// <copyright file="DraftNotificationSummary.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Models
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Draft Notification Summary model class.
    /// </summary>
    public class DraftNotificationSummary
    {
        /// <summary>
        /// Gets or sets Notification Id.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets Title value.
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets ScheduledDate value.
        /// </summary>
        public DateTime? ScheduledDate { get; set; }

        /// <summary>
        /// Gets or sets Unit Id value.
        /// </summary>
        public string UnitId { get; set; }

        /// <summary>
        /// Gets or sets a value of created by.
        /// </summary>
        public string CreatedBy { get; set; }

        /// <summary>
        /// Gets or sets a value of created on.
        /// </summary>
        public DateTime CreatedDateTime { get; set; }

        /// <summary>
        /// Gets or sets Groups audience name collection.
        /// </summary>
        public IEnumerable<string> GroupNames { get; set; }
    }
}
