using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models.ExternalConnectors;
using Microsoft.Graph.Models.ODataErrors;

namespace PurviewSearchConnector;

public static class GraphHelper
{
    private static GraphServiceClient? _graphClient;

    public static void InitializeGraph(Settings settings)
    {
        if (string.IsNullOrEmpty(settings.TenantID) ||
            string.IsNullOrEmpty(settings.ClientID) ||
            string.IsNullOrEmpty(settings.ClientSecret))
        {
            throw new Exception("Required settings (TenantID, ClientID, ClientSecret) missing from config.ini.");
        }

        var credential = new ClientSecretCredential(
            settings.TenantID,
            settings.ClientID,
            settings.ClientSecret);

        _graphClient = new GraphServiceClient(credential);
    }

    // We will add methods here later to interact with Graph API
    // Example placeholder:
    public static async Task<string?> GetTenantNameAsync()
    {
        if (_graphClient == null) throw new InvalidOperationException("Graph client not initialized.");

        try
        {
            var tenantInfo = await _graphClient.Organization
                                               .GetAsync(requestConfiguration =>
                                                {
                                                    requestConfiguration.QueryParameters.Select = new []{ "displayName" };
                                                });
            // Assuming the response contains at least one organization and it has a display name
            return tenantInfo?.Value?.FirstOrDefault()?.DisplayName;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting tenant name: {ex.Message}");
            return null;
        }
    }

    // Method to create a new external connection
    public static async Task<ExternalConnection?> CreateConnectionAsync(string connectionId, string connectionName, string connectionDescription)
    {
        if (_graphClient == null) throw new InvalidOperationException("Graph client not initialized.");

        Console.WriteLine($"Creating connection '{connectionName}' with ID '{connectionId}'...");

        var newConnection = new ExternalConnection
        {
            Id = connectionId,
            Name = connectionName,
            Description = connectionDescription
            // ConnectorId is typically required if using a Microsoft-provided connector, 
            // but usually not needed or set by Graph for a custom connector upon creation.
            // We might need to add ActivitySettings later depending on requirements.
        };

        try
        {
            // POST /external/connections
            var createdConnection = await _graphClient.External.Connections
                .PostAsync(newConnection);
            
            Console.WriteLine("Connection created successfully.");
            return createdConnection;
        }
        catch (ODataError odataError)
        {
            Console.WriteLine($"Error creating connection: {odataError.ResponseStatusCode} {odataError.Error?.Code} {odataError.Error?.Message}");
            return null;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating connection: {ex.Message}");
            return null;
        }
    }

    // Method to get existing connections
    public static async Task<List<ExternalConnection>?> GetConnectionsAsync()
    {
        if (_graphClient == null) throw new InvalidOperationException("Graph client not initialized.");

        Console.WriteLine("Getting existing connections...");

        try
        {
            // GET /external/connections
            // This retrieves connections owned by the application (due to OwnedBy permission)
            var connectionsResponse = await _graphClient.External.Connections.GetAsync();

            if (connectionsResponse?.Value != null)
            {
                Console.WriteLine($"Found {connectionsResponse.Value.Count} connections.");
                return connectionsResponse.Value;
            }
            else
            {
                Console.WriteLine("No connections found or response was empty.");
                return new List<ExternalConnection>(); // Return empty list
            }
        }
        catch (ODataError odataError)
        {
            Console.WriteLine($"Error getting connections: {odataError.ResponseStatusCode} {odataError.Error?.Code} {odataError.Error?.Message}");
            return null; // Indicate error
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting connections: {ex.Message}");
            return null; // Indicate error
        }
    }

    // Method to get the state of a specific connection
    public static async Task<ConnectionState?> GetConnectionStateAsync(string connectionId)
    {
        if (_graphClient == null) throw new InvalidOperationException("Graph client not initialized.");
        if (string.IsNullOrEmpty(connectionId)) throw new ArgumentNullException(nameof(connectionId));

        try
        {
            // GET /external/connections/{connectionId}?$select=state
            var connection = await _graphClient.External.Connections[connectionId]
                .GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Select = new string[] { "state" };
                });
            
            return connection?.State;
        }
        catch (ODataError odataError)
        {
            Console.WriteLine($"Error getting connection state: {odataError.ResponseStatusCode} {odataError.Error?.Code} {odataError.Error?.Message}");
            return null;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting connection state: {ex.Message}");
            return null;
        }
    }

    // Method to delete a connection
    public static async Task<bool> DeleteConnectionAsync(string connectionId)
    {
        if (_graphClient == null) throw new InvalidOperationException("Graph client not initialized.");
        if (string.IsNullOrEmpty(connectionId)) throw new ArgumentNullException(nameof(connectionId));

        Console.WriteLine($"Deleting connection with ID '{connectionId}'...");

        try
        {
            // DELETE /external/connections/{connectionId}
            await _graphClient.External.Connections[connectionId].DeleteAsync();
            
            Console.WriteLine("Connection deleted successfully.");
            return true;
        }
        catch (ODataError odataError)
        {
            // Handle case where connection might not be found (e.g., already deleted)
            if (odataError.ResponseStatusCode == 404)
            {
                 Console.WriteLine($"Connection with ID '{connectionId}' not found. Already deleted?");
                 return true; // Consider deletion successful if not found
            }
            Console.WriteLine($"Error deleting connection: {odataError.ResponseStatusCode} {odataError.Error?.Code} {odataError.Error?.Message}");
            return false;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error deleting connection: {ex.Message}");
            return false;
        }
    }

    // Method to register the schema for a connection
    public static async Task RegisterSchemaAsync(string connectionId)
    {
        if (_graphClient == null) throw new InvalidOperationException("Graph client not initialized.");
        if (string.IsNullOrEmpty(connectionId)) throw new ArgumentNullException(nameof(connectionId));

        Console.WriteLine("Registering schema...");

        var schema = new Schema
        {
            BaseType = "microsoft.graph.externalItem", // Required base type
            Properties = new List<Property>
            {
                new Property { Name = "termName", Type = PropertyType.String, IsSearchable = true, IsRetrievable = true, IsQueryable = true },
                new Property { Name = "definition", Type = PropertyType.String, IsSearchable = true, IsRetrievable = true, IsQueryable = true },
                new Property { Name = "status", Type = PropertyType.String, IsRetrievable = true, IsRefinable = true, IsQueryable = true },
                new Property { Name = "acronym", Type = PropertyType.String, IsSearchable = true, IsRetrievable = true, IsQueryable = true },
                new Property { Name = "glossaryName", Type = PropertyType.String, IsRetrievable = true, IsRefinable = true, IsQueryable = true },
                new Property { Name = "purviewUrl", Type = PropertyType.String, IsRetrievable = true },
                new Property { Name = "termGuid", Type = PropertyType.String, IsRetrievable = true, IsQueryable = true },
                new Property { Name = "lastModifiedTime", Type = PropertyType.DateTime, IsRetrievable = true, IsQueryable = true }
            }
        };

        try
        {
            // POST /external/connections/{connectionId}/schema
            // The SDK uses PUT for schema operations (Create or Update)
            await _graphClient.External.Connections[connectionId].Schema
                .PatchAsync(schema); 
                // Note: Schema provisioning is async on the service side. 
                // The call returns quickly, but the schema state becomes 'ready' later.
            
            Console.WriteLine("Schema registration request submitted successfully.");
            Console.WriteLine("Schema provisioning takes time. Check connection status later.");
        }
        catch (ODataError odataError)
        {
            Console.WriteLine($"Error registering schema: {odataError.ResponseStatusCode} {odataError.Error?.Code} {odataError.Error?.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error registering schema: {ex.Message}");
        }
    }

    // Method to get the schema for a connection
    public static async Task<Schema?> GetSchemaAsync(string connectionId)
    {
        if (_graphClient == null) throw new InvalidOperationException("Graph client not initialized.");
        if (string.IsNullOrEmpty(connectionId)) throw new ArgumentNullException(nameof(connectionId));

        Console.WriteLine($"Getting schema for connection ID '{connectionId}'...");

        try
        {
            // GET /external/connections/{connectionId}/schema
            var schema = await _graphClient.External.Connections[connectionId].Schema.GetAsync();
            
            Console.WriteLine("Schema retrieved successfully.");
            return schema;
        }
        catch (ODataError odataError)
        {
            // Handle case where connection might not be found 
            if (odataError.ResponseStatusCode == 404)
            {
                 Console.WriteLine($"Connection with ID '{connectionId}' not found.");
            }
            else
            {
                Console.WriteLine($"Error getting schema: {odataError.ResponseStatusCode} {odataError.Error?.Code} {odataError.Error?.Message}");
            }
            return null;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting schema: {ex.Message}");
            return null;
        }
    }

    // Method to add or update an external item
    // Returns true if successful, false otherwise
    public static async Task<bool> AddOrUpdateItemAsync(string connectionId, ExternalItem item)
    {
        if (_graphClient == null) throw new InvalidOperationException("Graph client not initialized.");
        if (string.IsNullOrEmpty(connectionId)) throw new ArgumentNullException(nameof(connectionId));
        ArgumentNullException.ThrowIfNull(item);
        if (string.IsNullOrEmpty(item.Id)) throw new ArgumentException("ExternalItem must have an ID.", nameof(item));

        Console.WriteLine($"Upserting item {item.Id}...");
        try
        {
            var result = await _graphClient.External.Connections[connectionId].Items[item.Id]
                            .PutAsync(item);
            Console.WriteLine($"Item {item.Id} upserted successfully.");
            return true; // Indicate success
        }
        catch (ODataError odataError)
        {
             // Log more details from the error object if available
             Console.WriteLine($"ERROR upserting item {item.Id}: {odataError.ResponseStatusCode} {odataError.Error?.Code}");
             Console.WriteLine($"  Message: {odataError.Error?.Message}");
             if (odataError.Error?.Details != null)
             {
                Console.WriteLine("  Details:");
                foreach(var detail in odataError.Error.Details)
                {
                    Console.WriteLine($"    Code: {detail.Code}, Message: {detail.Message}, Target: {detail.Target}");
                }
             }
             // Also log inner error if present
             if (odataError.Error?.InnerError != null)
             {
                Console.WriteLine($"  Inner Error: {odataError.Error.InnerError}"); // May contain stack trace or more info
             }
             return false; // Indicate failure
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ERROR upserting item {item.Id}: {ex.Message}");
            return false; // Indicate failure
        }
    }

    // Method to delete an external item
    // Returns true if successful or item not found, false otherwise
    public static async Task<bool> DeleteItemAsync(string connectionId, string itemId)
    {
        if (_graphClient == null) throw new InvalidOperationException("Graph client not initialized.");
        if (string.IsNullOrEmpty(connectionId)) throw new ArgumentNullException(nameof(connectionId));
        if (string.IsNullOrEmpty(itemId)) throw new ArgumentNullException(nameof(itemId));

        Console.WriteLine($"Deleting item {itemId}...");
        try
        {
            await _graphClient.External.Connections[connectionId].Items[itemId]
                .DeleteAsync();
            Console.WriteLine($"Item {itemId} delete request submitted successfully.");
            return true; // Indicate success
        }
        catch (ODataError odataError)
        {
            if (odataError.ResponseStatusCode == 404)
            {
                Console.WriteLine($"Item {itemId} not found. Already deleted?");
                return true; // Consider it a success if already gone
            }
            else
            {
                Console.WriteLine($"ERROR deleting item {itemId}: {odataError.ResponseStatusCode} {odataError.Error?.Code} {odataError.Error?.Message}");
                return false; // Indicate failure
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"ERROR deleting item {itemId}: {ex.Message}");
            return false; // Indicate failure
        }
    }

    // Method to get all item IDs within a connection (handles pagination)
    public static async Task<List<string>?> GetConnectionItemIdsAsync(string connectionId)
    {
        if (_graphClient == null) throw new InvalidOperationException("Graph client not initialized.");
        if (string.IsNullOrEmpty(connectionId)) throw new ArgumentNullException(nameof(connectionId));

        Console.WriteLine($"Getting all item IDs for connection '{connectionId}'...");
        var itemIds = new List<string>();

        try
        {
            // GET /external/connections/{connectionId}/items?$select=id
            // Use the SDK's iterator pattern if available, or handle paging manually
            var itemsResponse = await _graphClient.External.Connections[connectionId].Items
                .GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Select = new string[] { "id" };
                    // Optionally add $top for page size if needed, e.g.:
                    // requestConfiguration.QueryParameters.Top = 100;
                });

            // The Graph SDK typically provides ways to automatically handle pagination.
            // Check if itemsResponse or a subsequent call handles iteration.
            // Assuming itemsResponse.Value directly contains the items for the first page
            // and itemsResponse.OdataNextLink provides the URL for the next page.

            var pageIterator = PageIterator<ExternalItem, ExternalItemCollectionResponse>
                .CreatePageIterator(
                    _graphClient, 
                    itemsResponse, 
                    (item) => { itemIds.Add(item.Id ?? string.Empty); return true; } 
                    // Process each item - add its ID to our list
                );
                
            await pageIterator.IterateAsync(); // Iterate through all pages

            // Filter out any potential empty strings if an item somehow had no ID
            itemIds.RemoveAll(string.IsNullOrEmpty);

            Console.WriteLine($"Found {itemIds.Count} item IDs in the connection.");
            return itemIds;
        }
        catch (ODataError odataError)
        {
             // Log more details from the error object
             Console.WriteLine($"ERROR getting item IDs: {odataError.ResponseStatusCode} {odataError.Error?.Code}");
             Console.WriteLine($"  Message: {odataError.Error?.Message}");
             if (odataError.Error?.Details != null)
             {
                Console.WriteLine("  Details:");
                foreach(var detail in odataError.Error.Details)
                {
                    Console.WriteLine($"    Code: {detail.Code}, Message: {detail.Message}, Target: {detail.Target}");
                }
             }
             if (odataError.Error?.InnerError != null)
             {
                Console.WriteLine($"  Inner Error: {odataError.Error.InnerError}");
             }
             return null;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting item IDs: {ex.Message}");
            return null;
        }
    }
} 