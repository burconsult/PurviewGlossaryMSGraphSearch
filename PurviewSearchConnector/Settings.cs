using Microsoft.Extensions.Configuration;

namespace PurviewSearchConnector;

public class Settings
{
    // Azure Credentials (match INI keys)
    public string? ClientID { get; set; }
    public string? ClientSecret { get; set; }
    public string? TenantID { get; set; }

    // Purview Settings (match INI key)
    public string? AccountName { get; set; }

    public static Settings LoadSettings()
    {
        // Build configuration
        IConfiguration config = new ConfigurationBuilder()
            .SetBasePath(System.IO.Directory.GetCurrentDirectory())
            .AddIniFile("config.ini", optional: false, reloadOnChange: true) // Ensure file exists
            .Build();

        // Bind sections to a new Settings object
        var settings = new Settings();
        config.GetSection("Azure").Bind(settings);
        config.GetSection("Purview").Bind(settings);

        // Validate essential settings using the new property names
        if (string.IsNullOrEmpty(settings.ClientID) ||
            string.IsNullOrEmpty(settings.ClientSecret) ||
            string.IsNullOrEmpty(settings.TenantID) ||
            string.IsNullOrEmpty(settings.AccountName))
        {
            throw new Exception("Could not load all required app settings from config.ini. Check [Azure] (ClientID, ClientSecret, TenantID) and [Purview] (AccountName) sections.");
        }

        return settings;
    }
} 