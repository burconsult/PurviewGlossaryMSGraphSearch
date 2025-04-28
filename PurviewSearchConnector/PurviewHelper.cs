using Azure.Identity;
using Azure.Analytics.Purview.DataMap;
using System;
using Azure;
using System.Text.RegularExpressions;
using Microsoft.Graph.Models.ExternalConnectors;

namespace PurviewSearchConnector;

public static class PurviewHelper
{
    private static DataMapClient? _purviewClient;
    private static Settings? _settings;

    public static void InitializePurview(Settings settings)
    {
        _settings = settings;

        if (string.IsNullOrEmpty(settings.TenantID) ||
            string.IsNullOrEmpty(settings.ClientID) ||
            string.IsNullOrEmpty(settings.ClientSecret) ||
            string.IsNullOrEmpty(settings.AccountName))
        {
            throw new Exception("Required settings for Purview client initialization are missing.");
        }

        try
        {
            Uri purviewEndpoint = new($"https://{settings.AccountName}.purview.azure.com");
            var credential = new ClientSecretCredential(
                settings.TenantID,
                settings.ClientID,
                settings.ClientSecret);

            _purviewClient = new DataMapClient(purviewEndpoint, credential);

            Console.WriteLine("Purview Data Map client initialized successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error initializing Purview client: {ex.Message}");
            throw;
        }
    }

    public static async Task TestPurviewConnectionAsync()
    {
        if (_purviewClient == null || _settings == null)
        {
            throw new InvalidOperationException("Purview client not initialized.");
        }

        // === Original Glossary Test (Uncommented) ===
        Console.WriteLine("Testing Purview connection by listing glossaries...");
        try
        {
            Glossary glossaryClient = _purviewClient.GetGlossaryClient();
            
            // Use BatchGetAsync() and await it
            Response<IReadOnlyList<AtlasGlossary>> response = await glossaryClient.BatchGetAsync(); // Use async version
            
            Console.WriteLine("Successfully called BatchGetAsync, processing results...");
            int count = 0;
            if (response.Value != null)
            {
                foreach (AtlasGlossary glossary in response.Value)
                {
                    // Process each AtlasGlossary object
                    count++;
                    // You can now access properties directly, e.g., glossary.Name, glossary.Guid
                    Console.WriteLine($"  - Found Glossary: {glossary.Name} (GUID: {glossary.Guid})"); 
                }
            }
            
            Console.WriteLine($"Successfully connected to Purview account '{_settings.AccountName}'. Found {count} glossaries.");
        }
        catch (Azure.RequestFailedException rfEx)
        {
            Console.WriteLine($"Error testing Purview connection (Glossary): {rfEx.Status} - {rfEx.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred during Purview glossary test: {ex.Message}");
        }

        // === Type Definition Test (Commented Out) ===
        // Console.WriteLine("Testing Purview connection by getting 'AtlasGlossary' type definition...");
        // try
        // {
        //     TypeDefinition typeDefinitionClient = _purviewClient.GetTypeDefinitionClient();
        //     Response response = await typeDefinitionClient.GetByNameAsync("AtlasGlossary", new Azure.RequestContext());
        //     if (response.Status >= 200 && response.Status < 300)
        //     {
        //         Console.WriteLine($"Successfully connected to Purview account '{_settings.PurviewAccountName}' and retrieved type definition (Status: {response.Status}).");
        //     }
        //     else
        //     {
        //         Console.WriteLine($"Connected to Purview, but failed to get type definition. Status: {response.Status}, Reason: {response.ReasonPhrase}");
        //     }
        // }
        // catch (Azure.RequestFailedException rfEx)
        // {
        //     Console.WriteLine($"Error testing Purview connection (Type Definition): {rfEx.Status} - {rfEx.Message}");
        // }
        // catch (Exception ex)
        // {
        //     Console.WriteLine($"An unexpected error occurred during Purview type definition test: {ex.Message}");
        // }
    }

    // Method to fetch terms and their details, optionally for a single glossary
    // Returns a list of tuples, pairing each term with its glossary name
    public static async Task<List<(AtlasGlossaryTerm Term, string GlossaryName)>> GetGlossaryTermsDataAsync(string? targetGlossaryId = null)
    {
        if (_purviewClient == null || _settings == null)
        {
            throw new InvalidOperationException("Purview client not initialized.");
        }

        var allTermsWithContext = new List<(AtlasGlossaryTerm Term, string GlossaryName)>();
        Glossary glossaryClient = _purviewClient.GetGlossaryClient();
        var glossariesToProcess = new List<AtlasGlossary>();

        if (!string.IsNullOrEmpty(targetGlossaryId))
        {
            Console.WriteLine($"\nFetching target glossary with ID: {targetGlossaryId}...");
            try
            {
                Response<AtlasGlossary> targetGlossaryResponse = await glossaryClient.GetGlossaryAsync(targetGlossaryId);
                if (targetGlossaryResponse.Value != null)
                {
                    glossariesToProcess.Add(targetGlossaryResponse.Value);
                    Console.WriteLine($"Found target glossary: {targetGlossaryResponse.Value.Name}");
                }
                else
                {
                    Console.WriteLine($"Target glossary with ID {targetGlossaryId} not found or response was empty.");
                    return allTermsWithContext; // Return empty if target not found
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ERROR fetching target glossary {targetGlossaryId}: {ex.Message}");
                return allTermsWithContext; // Return empty on error
            }
        }
        else
        {
            Console.WriteLine("\nFetching all glossaries...");
            try
            {
                Response<IReadOnlyList<AtlasGlossary>> glossariesResponse = await glossaryClient.BatchGetAsync();
                if (glossariesResponse.Value != null)
                {
                    glossariesToProcess.AddRange(glossariesResponse.Value);
                    Console.WriteLine($"Found {glossariesToProcess.Count} glossaries.");
                }
                else
                {
                     Console.WriteLine("Could not retrieve glossaries.");
                     return allTermsWithContext;
                }
            }
            catch (Exception ex)
            {
                 Console.WriteLine($"ERROR fetching all glossaries: {ex.Message}");
                 return allTermsWithContext;
            }
        }

        Console.WriteLine("Fetching terms for selected glossaries...");
        foreach (var glossary in glossariesToProcess)
        {
            if (string.IsNullOrEmpty(glossary.Guid))
            {
                Console.WriteLine($"Skipping glossary '{glossary.Name}' due to missing GUID.");
                continue;
            }

            string currentGlossaryName = glossary.Name ?? "Unknown Glossary";

            Console.WriteLine($"  Fetching terms for glossary: {currentGlossaryName} ({glossary.Guid})...");
            try
            {
                Response<IReadOnlyList<AtlasGlossaryTerm>> termsResponse = 
                    await glossaryClient.GetTermsAsync(glossary.Guid); 

                if (termsResponse.Value != null && termsResponse.Value.Count > 0)
                {
                    Console.WriteLine($"    Found {termsResponse.Value.Count} terms.");
                    foreach (var term in termsResponse.Value)
                    {
                        allTermsWithContext.Add((term, currentGlossaryName));
                    }
                }
                else
                {
                    Console.WriteLine("    No terms found in this glossary.");
                }
            }
            catch (Exception glossEx)
            {
                Console.WriteLine($"  ERROR fetching terms for glossary {glossary.Name}: {glossEx.Message}");
                // Continue to the next glossary
            }
        }

        Console.WriteLine($"\nFinished fetching. Total terms retrieved: {allTermsWithContext.Count}");
        return allTermsWithContext;
    }

    // Utility method to remove HTML tags from a string
    public static string CleanHtml(string? htmlInput)
    {
        if (string.IsNullOrEmpty(htmlInput))
        {
            return string.Empty;
        }

        // Simple regex to remove HTML tags
        return Regex.Replace(htmlInput, "<.*?>", string.Empty);
        
        // Consider more robust HTML parsing/sanitization library if complex HTML is present
        // e.g., HtmlAgilityPack (would require adding NuGet package)
    }

    // Add method later to fetch glossary terms
    // public static async Task<List<string>> GetGlossaryTermsAsync() { ... }

    // Function to map Purview term data to a Graph ExternalItem
    // Moved from Program.cs
    public static ExternalItem MapPurviewTermToExternalItem(AtlasGlossaryTerm term, string glossaryName, string tenantId)
    {
        // Basic null checks
        ArgumentNullException.ThrowIfNull(term);
        if (string.IsNullOrEmpty(term.Guid)) 
        { 
            throw new ArgumentException("Term must have a GUID.", nameof(term)); 
        }

        // Retrieve properties safely
        string termName = term.Name ?? "Unknown Term";
        string definition = term.LongDescription ?? term.ShortDescription ?? string.Empty;
        
        // Get original status and translate
        string originalStatus = term.Status?.ToString() ?? "Unknown";
        string translatedStatus = originalStatus; // Default to original if no translation found

        var statusTranslations = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) 
        {
            { "Approved", "Anbefalt" },
            { "Draft", "Foreslått" },
            { "Alert", "Frarådet" }, 
            { "Expired", "Utgått" }
            // Add other statuses if needed
        };

        if (statusTranslations.TryGetValue(originalStatus, out string? norwegianStatus))
        {
            translatedStatus = norwegianStatus;
        }
        else if (originalStatus == "Unknown")
        {
            translatedStatus = "Ukjent"; // Translate default Unknown
        }

        string status = translatedStatus; // Use the translated status

        string acronym = term.Abbreviation ?? string.Empty;
        string termGuid = term.Guid; // Already checked for null/empty

        // Glossary name is passed as parameter
        // Construct Purview URL
        string purviewUrl = $"https://purview.microsoft.com/datacatalog/governance/main/catalog/glossary/term?termGuid={termGuid}";

        string cleanedDefinition = CleanHtml(definition); // Use CleanHtml from this class
        
        var properties = new Microsoft.Graph.Models.ExternalConnectors.Properties(); 
        properties.AdditionalData = new Dictionary<string, object>
        {
            { "termName", termName },
            { "definition", cleanedDefinition }, 
            { "status", status }, // Add translated status to properties
            { "acronym", acronym },
            { "glossaryName", glossaryName }, 
            { "purviewUrl", purviewUrl },
            { "termGuid", termGuid }
        };

        // Re-add lastModifiedTime conversion and addition
        if (term.UpdateTime.HasValue)
        {
            try 
            {
                DateTimeOffset lastModifiedOffset = DateTimeOffset.FromUnixTimeMilliseconds(term.UpdateTime.Value);
                properties.AdditionalData.Add("lastModifiedTime", lastModifiedOffset.ToString("o")); 
            }
            catch (ArgumentOutOfRangeException ex)
            {
                // Consider logging this warning differently in Azure Function context (e.g., using ILogger)
                Console.WriteLine($"Warning: Could not convert UpdateTime '{term.UpdateTime.Value}' for term {termGuid}. Skipping lastModifiedTime property. Error: {ex.Message}");
            }
        }

        // --- Create ExternalItem with properties and ACL ---
        var externalItem = new ExternalItem
        {
            Id = termGuid, 
            Content = new ExternalItemContent
            {
                Type = ExternalItemContentType.Text,
                Value = cleanedDefinition 
            },
            Acl = new List<Acl> // Add ACL back
            {
                new Acl { AccessType = AccessType.Grant, Type = AclType.Everyone, Value = tenantId! } // Use non-null assertion for tenantId
            },
            Properties = properties
        };

        return externalItem;
    }
} 