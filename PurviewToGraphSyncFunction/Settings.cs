using Microsoft.Extensions.Configuration;

namespace PurviewToGraphSyncFunction;

// Simple POCO class to hold settings loaded from environment variables
public class Settings
{
    // Azure AD App Credentials
    public string? ClientID { get; set; }
    public string? ClientSecret { get; set; }
    public string? TenantID { get; set; }

    // Purview Settings
    public string? PurviewEndpoint { get; set; } // e.g., https://YourPurviewAccountName.purview.azure.com

    // Graph Connector Settings
    public string? GraphConnectionId { get; set; }

    // Azure Function / Timestamp Storage Settings
    public string? AzureWebJobsStorage { get; set; } // Standard Azure Functions connection string
    public string? TimestampContainerName { get; set; } // e.g., "purview-sync-timestamps"
} 