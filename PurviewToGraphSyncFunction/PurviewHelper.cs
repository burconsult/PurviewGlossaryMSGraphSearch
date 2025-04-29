using Azure.Identity;
using Azure.Analytics.Purview.DataMap;
using System;
using Azure;
using System.Text.RegularExpressions;
using Microsoft.Graph.Models.ExternalConnectors;
using Microsoft.Extensions.Logging;

namespace PurviewToGraphSyncFunction;

public static class PurviewHelper
{
    private static DataMapClient? _purviewClient;
    private static Settings? _settings;
    private static ILogger? _logger;

    public static void InitializePurview(Settings settings, ILogger logger)
    {
        _settings = settings;
        _logger = logger;

        _logger.LogInformation("Initializing Purview Data Map client...");

        if (string.IsNullOrEmpty(settings.TenantID) ||
            string.IsNullOrEmpty(settings.ClientID) ||
            string.IsNullOrEmpty(settings.ClientSecret) ||
            string.IsNullOrEmpty(settings.PurviewEndpoint))
        {
            _logger.LogError("Required Purview settings (TenantID, ClientID, ClientSecret, PurviewEndpoint) are missing.");
            throw new ArgumentNullException(nameof(settings), "Required Purview settings are missing.");
        }

        try
        {
            Uri purviewEndpointUri = new Uri(settings.PurviewEndpoint);
            var credential = new ClientSecretCredential(
                settings.TenantID,
                settings.ClientID,
                settings.ClientSecret);

            _purviewClient = new DataMapClient(purviewEndpointUri, credential);

            _logger.LogInformation("Purview Data Map client initialized successfully for endpoint {PurviewEndpoint}", settings.PurviewEndpoint);
        }
        catch (UriFormatException ex)
        {
            _logger.LogError(ex, "Invalid Purview Endpoint format: {PurviewEndpoint}", settings.PurviewEndpoint);
            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error initializing Purview client: {Message}", ex.Message);
            throw;
        }
    }

    public static async Task TestPurviewConnectionAsync()
    {
        if (_purviewClient == null || _settings == null || _logger == null)
        {
            throw new InvalidOperationException("Purview client or logger not initialized.");
        }

        _logger.LogInformation("Testing Purview connection by listing glossaries for endpoint {PurviewEndpoint}...", _settings.PurviewEndpoint);
        try
        {
            Glossary glossaryClient = _purviewClient.GetGlossaryClient();
            Response<IReadOnlyList<AtlasGlossary>> response = await glossaryClient.BatchGetAsync();

            _logger.LogInformation("Successfully called Purview BatchGetAsync, processing results...");
            int count = 0;
            if (response.Value != null)
            {
                foreach (AtlasGlossary glossary in response.Value)
                {
                    count++;
                    _logger.LogInformation("  - Found Glossary: {GlossaryName} (GUID: {GlossaryGuid})", glossary.Name, glossary.Guid);
                }
            }

            _logger.LogInformation("Successfully connected to Purview endpoint {PurviewEndpoint}. Found {GlossaryCount} glossaries.", _settings.PurviewEndpoint, count);
        }
        catch (RequestFailedException rfEx)
        {
            _logger.LogError(rfEx, "RequestFailedException testing Purview connection (Glossary): Status {StatusCode} - {ErrorMessage}", rfEx.Status, rfEx.Message);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "An unexpected error occurred during Purview glossary test: {ErrorMessage}", ex.Message);
        }
    }

    public static async Task<List<(AtlasGlossaryTerm Term, string GlossaryName)>> GetGlossaryTermsDataAsync(string? targetGlossaryId = null, DateTimeOffset? lastSyncTime = null)
    {
        if (_purviewClient == null || _settings == null || _logger == null)
        {
            throw new InvalidOperationException("Purview client or logger not initialized.");
        }

        var allTermsWithContext = new List<(AtlasGlossaryTerm Term, string GlossaryName)>();
        Glossary glossaryClient = _purviewClient.GetGlossaryClient();
        var glossariesToProcess = new List<AtlasGlossary>();

        if (!string.IsNullOrEmpty(targetGlossaryId))
        {
            _logger.LogInformation("\nFetching target glossary with ID: {GlossaryId}...", targetGlossaryId);
            try
            {
                Response<AtlasGlossary> targetGlossaryResponse = await glossaryClient.GetGlossaryAsync(targetGlossaryId);
                if (targetGlossaryResponse.Value != null)
                {
                    glossariesToProcess.Add(targetGlossaryResponse.Value);
                    _logger.LogInformation("Found target glossary: {GlossaryName}", targetGlossaryResponse.Value.Name);
                }
                else
                {
                    _logger.LogWarning("Target glossary with ID {GlossaryId} not found or response was empty.", targetGlossaryId);
                    return allTermsWithContext;
                }
            }
            catch (RequestFailedException rfEx)
            {
                _logger.LogError(rfEx, "RequestFailedException fetching target glossary {GlossaryId}: Status {StatusCode} - {ErrorMessage}", targetGlossaryId, rfEx.Status, rfEx.Message);
                return allTermsWithContext;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "ERROR fetching target glossary {GlossaryId}: {ErrorMessage}", targetGlossaryId, ex.Message);
                return allTermsWithContext;
            }
        }
        else
        {
            _logger.LogInformation("Fetching all glossaries...");
            try
            {
                Response<IReadOnlyList<AtlasGlossary>> glossariesResponse = await glossaryClient.BatchGetAsync();
                if (glossariesResponse.Value != null)
                {
                    glossariesToProcess.AddRange(glossariesResponse.Value);
                    _logger.LogInformation("Found {GlossaryCount} glossaries.", glossariesToProcess.Count);
                }
                else
                {
                    _logger.LogWarning("Could not retrieve glossaries list.");
                    return allTermsWithContext;
                }
            }
            catch (RequestFailedException rfEx)
            {
                _logger.LogError(rfEx, "RequestFailedException fetching all glossaries: Status {StatusCode} - {ErrorMessage}", rfEx.Status, rfEx.Message);
                return allTermsWithContext;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "ERROR fetching all glossaries: {ErrorMessage}", ex.Message);
                return allTermsWithContext;
            }
        }

        _logger.LogInformation("Fetching terms for {GlossaryCount} selected glossaries...", glossariesToProcess.Count);
        foreach (var glossary in glossariesToProcess)
        {
            if (string.IsNullOrEmpty(glossary.Guid))
            {
                _logger.LogWarning("Skipping glossary '{GlossaryName}' due to missing GUID.", glossary.Name ?? "<null_name>");
                continue;
            }

            string currentGlossaryName = glossary.Name ?? "Unknown Glossary";

            _logger.LogInformation("  Fetching terms for glossary: {GlossaryName} ({GlossaryGuid})...", currentGlossaryName, glossary.Guid);
            try
            {
                Response<IReadOnlyList<AtlasGlossaryTerm>> termsResponse = await glossaryClient.GetTermsAsync(glossary.Guid);

                if (termsResponse.Value != null && termsResponse.Value.Count > 0)
                {
                    int termsFound = termsResponse.Value.Count;
                    int termsAdded = 0;
                    _logger.LogInformation("    Found {TermsFoundCount} terms in glossary {GlossaryName}. Checking last sync time: {LastSyncTime}", termsFound, currentGlossaryName, lastSyncTime?.ToString("o") ?? "None");

                    foreach (var term in termsResponse.Value)
                    {
                        bool processTerm = true;
                        if (lastSyncTime.HasValue && term.UpdateTime.HasValue)
                        {
                            try
                            {
                                DateTimeOffset termUpdateTime = DateTimeOffset.FromUnixTimeMilliseconds(term.UpdateTime.Value);
                                if (termUpdateTime <= lastSyncTime.Value)
                                {
                                    processTerm = false;
                                }
                            }
                            catch (ArgumentOutOfRangeException ex)
                            {
                                _logger.LogWarning(ex, "Could not convert UpdateTime '{PurviewUpdateTime}' for term {TermGuid} ({TermName}). Processing term anyway.", term.UpdateTime.Value, term.Guid, term.Name);
                            }
                        }
                        else if (lastSyncTime.HasValue && !term.UpdateTime.HasValue)
                        {
                            _logger.LogWarning("Term {TermGuid} ({TermName}) has no UpdateTime. Processing term anyway for incremental sync.", term.Guid, term.Name);
                        }

                        if (processTerm)
                        {
                            allTermsWithContext.Add((term, currentGlossaryName));
                            termsAdded++;
                        }
                    }
                    _logger.LogInformation("    Added {TermsAddedCount} terms from glossary {GlossaryName} to be processed (based on last sync time).", termsAdded, currentGlossaryName);
                }
                else
                {
                    _logger.LogInformation("    No terms found in glossary {GlossaryName}.", currentGlossaryName);
                }
            }
            catch (RequestFailedException rfEx)
            {
                _logger.LogError(rfEx, "RequestFailedException fetching terms for glossary {GlossaryName} ({GlossaryGuid}): Status {StatusCode} - {ErrorMessage}", currentGlossaryName, glossary.Guid, rfEx.Status, rfEx.Message);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "ERROR fetching terms for glossary {GlossaryName} ({GlossaryGuid}): {ErrorMessage}", currentGlossaryName, glossary.Guid, ex.Message);
            }
        }

        _logger.LogInformation("Finished fetching terms. Total terms to process: {TotalTermCount}", allTermsWithContext.Count);
        return allTermsWithContext;
    }

    public static string CleanHtml(string? htmlInput)
    {
        if (string.IsNullOrEmpty(htmlInput))
        {
            return string.Empty;
        }

        string noHtml = Regex.Replace(htmlInput, "<.*?>", string.Empty);
        string decoded = System.Net.WebUtility.HtmlDecode(noHtml);

        return decoded.Trim();
    }

    public static ExternalItem MapPurviewTermToExternalItem(AtlasGlossaryTerm term, string glossaryName, string tenantId, ILogger logger)
    {
        ArgumentNullException.ThrowIfNull(term);
        if (string.IsNullOrEmpty(term.Guid))
        {
            throw new ArgumentException("Term must have a GUID.", nameof(term));
        }

        string termName = term.Name ?? "Unknown Term";
        string definition = term.LongDescription ?? term.ShortDescription ?? string.Empty;
        
        string originalStatus = term.Status?.ToString() ?? "Unknown";
        string translatedStatus = originalStatus;

        var statusTranslations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "Approved", "Anbefalt" },
            { "Draft", "Foreslått" },
            { "Alert", "Frarådet" },
            { "Expired", "Utgått" }
        };

        if (statusTranslations.TryGetValue(originalStatus, out string? norwegianStatus))
        {
            translatedStatus = norwegianStatus;
        }
        else if (originalStatus == "Unknown")
        {
            translatedStatus = "Ukjent";
        }

        string status = translatedStatus;

        string acronym = term.Abbreviation ?? string.Empty;
        string termGuid = term.Guid;

        // Construct the modern Purview portal URL directly, as the UI is now unified at purview.microsoft.com
        // The previous logic tried to convert the API endpoint, leading to incorrect URLs.
        const string purviewTermUrlFormat = "https://purview.microsoft.com/datacatalog/governance/main/catalog/glossary/term?termGuid={0}";
        string purviewUrl = string.Format(purviewTermUrlFormat, termGuid);

        string cleanedDefinition = CleanHtml(definition);
        
        var properties = new Microsoft.Graph.Models.ExternalConnectors.Properties();
        properties.AdditionalData = new Dictionary<string, object>
        {
            { "termName", termName },
            { "definition", cleanedDefinition },
            { "status", status },
            { "acronym", acronym },
            { "glossaryName", glossaryName },
            { "purviewUrl", purviewUrl },
            { "termGuid", termGuid }
        };

        if (term.UpdateTime.HasValue)
        {
            try
            {
                DateTimeOffset lastModifiedOffset = DateTimeOffset.FromUnixTimeMilliseconds(term.UpdateTime.Value);
                properties.AdditionalData.Add("lastModifiedTime", lastModifiedOffset.ToString("o"));
            }
            catch (ArgumentOutOfRangeException ex)
            {
                logger.LogWarning(ex, "Could not convert UpdateTime '{PurviewUpdateTime}' for term {TermGuid}. Skipping lastModifiedTime property.", term.UpdateTime.Value, termGuid);
            }
        }

        var externalItem = new ExternalItem
        {
            Id = termGuid,
            Content = new ExternalItemContent
            {
                Type = ExternalItemContentType.Text,
                Value = cleanedDefinition
            },
            Acl = new List<Acl>
            {
                new Acl { AccessType = AccessType.Grant, Type = AclType.Everyone, Value = tenantId! }
            },
            Properties = properties
        };

        return externalItem;
    }

    public static async Task<string?> GetGlossaryIdByNameAsync(string glossaryName)
    {
        if (_purviewClient == null || _settings == null || _logger == null)
        {
            throw new InvalidOperationException("Purview client or logger not initialized.");
        }
        ArgumentException.ThrowIfNullOrEmpty(glossaryName);

        _logger.LogInformation("Attempting to find GUID for glossary name: {GlossaryName}", glossaryName);
        try
        {
            Glossary glossaryClient = _purviewClient.GetGlossaryClient();
            Response<IReadOnlyList<AtlasGlossary>> response = await glossaryClient.BatchGetAsync();
            
            if (response.Value != null)
            {
                foreach (AtlasGlossary glossary in response.Value)
                {
                    if (string.Equals(glossary.Name, glossaryName, StringComparison.OrdinalIgnoreCase))
                    {
                        _logger.LogInformation("Found matching glossary: {GlossaryName} with GUID {GlossaryGuid}", glossaryName, glossary.Guid);
                        return glossary.Guid;
                    }
                }
            }
            _logger.LogWarning("Glossary with name {GlossaryName} not found.", glossaryName);
            return null;
        }
        catch (RequestFailedException rfEx)
        {
            _logger.LogError(rfEx, "RequestFailedException finding glossary GUID for {GlossaryName}: Status {StatusCode} - {ErrorMessage}", glossaryName, rfEx.Status, rfEx.Message);
            return null;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error finding glossary GUID for {GlossaryName}: {ErrorMessage}", glossaryName, ex.Message);
            return null;
        }
    }
} 