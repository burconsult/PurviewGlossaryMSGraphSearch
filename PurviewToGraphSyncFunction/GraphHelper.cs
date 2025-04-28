using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models.ExternalConnectors;
using Microsoft.Graph.Models.ODataErrors;
using Microsoft.Extensions.Logging;

namespace PurviewToGraphSyncFunction;

public static class GraphHelper
{
    private static GraphServiceClient? _graphClient;
    private static ILogger? _logger;

    public static void InitializeGraph(Settings settings, ILogger logger)
    {
        _logger = logger;

        _logger.LogInformation("Initializing Graph client...");

        if (string.IsNullOrEmpty(settings.TenantID) ||
            string.IsNullOrEmpty(settings.ClientID) ||
            string.IsNullOrEmpty(settings.ClientSecret))
        {
            _logger.LogError("Required Graph settings (TenantID, ClientID, ClientSecret) are missing.");
            throw new ArgumentNullException(nameof(settings), "Required Graph settings (TenantID, ClientID, ClientSecret) are missing.");
        }

        var credential = new ClientSecretCredential(
            settings.TenantID,
            settings.ClientID,
            settings.ClientSecret);

        _graphClient = new GraphServiceClient(credential);
        _logger.LogInformation("Graph client initialized successfully.");
    }

    public static async Task<string?> GetTenantNameAsync()
    {
        if (_graphClient == null || _logger == null) throw new InvalidOperationException("Graph client or logger not initialized.");

        _logger.LogInformation("Attempting to retrieve tenant display name...");
        try
        {
            var tenantInfo = await _graphClient.Organization
                                               .GetAsync(requestConfiguration =>
                                                {
                                                    requestConfiguration.QueryParameters.Select = new []{ "displayName" };
                                                });
            var displayName = tenantInfo?.Value?.FirstOrDefault()?.DisplayName;
             _logger.LogInformation("Successfully retrieved tenant display name: {DisplayName}", displayName ?? "N/A");
            return displayName;
        }
        catch (ODataError odataError)
        {
            _logger?.LogError(odataError, "OData Error getting tenant name: {StatusCode} {Code} {Message}", odataError.ResponseStatusCode, odataError.Error?.Code, odataError.Error?.Message);
             return null;
        }
        catch (Exception ex)
        {
            _logger?.LogError(ex, "Error getting tenant name: {Message}", ex.Message);
            return null;
        }
    }

    // Method to create a new external connection
    public static async Task<ExternalConnection?> CreateConnectionAsync(string connectionId, string connectionName, string connectionDescription)
    {
        if (_graphClient == null || _logger == null) throw new InvalidOperationException("Graph client or logger not initialized.");

        _logger.LogInformation("Creating connection {ConnectionName} with ID {ConnectionId}...", connectionName, connectionId);

        var newConnection = new ExternalConnection
        {
            Id = connectionId,
            Name = connectionName,
            Description = connectionDescription
        };

        try
        {
            // POST /external/connections
            var createdConnection = await _graphClient.External.Connections
                .PostAsync(newConnection);

            _logger.LogInformation("Connection {ConnectionId} created successfully.", connectionId);
            return createdConnection;
        }
        catch (ODataError odataError)
        {
            _logger.LogError(odataError, "OData Error creating connection {ConnectionId}: {StatusCode} {Code} {Message}", connectionId, odataError.ResponseStatusCode, odataError.Error?.Code, odataError.Error?.Message);
            return null;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error creating connection {ConnectionId}: {Message}", connectionId, ex.Message);
            return null;
        }
    }

    // Method to get existing connections
    public static async Task<List<ExternalConnection>?> GetConnectionsAsync()
    {
         if (_graphClient == null || _logger == null) throw new InvalidOperationException("Graph client or logger not initialized.");

        _logger.LogInformation("Getting existing connections...");

        try
        {
            // GET /external/connections
            var connectionsResponse = await _graphClient.External.Connections.GetAsync();

            if (connectionsResponse?.Value != null)
            {
                _logger.LogInformation("Found {ConnectionCount} connections.", connectionsResponse.Value.Count);
                return connectionsResponse.Value;
            }
            else
            {
                _logger.LogInformation("No connections found or response was empty.");
                return new List<ExternalConnection>(); // Return empty list
            }
        }
        catch (ODataError odataError)
        {
            _logger.LogError(odataError, "OData Error getting connections: {StatusCode} {Code} {Message}", odataError.ResponseStatusCode, odataError.Error?.Code, odataError.Error?.Message);
            return null; // Indicate error
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error getting connections: {Message}", ex.Message);
            return null; // Indicate error
        }
    }

    // Method to get the state of a specific connection
    public static async Task<ConnectionState?> GetConnectionStateAsync(string connectionId)
    {
        if (_graphClient == null || _logger == null) throw new InvalidOperationException("Graph client or logger not initialized.");
        if (string.IsNullOrEmpty(connectionId))
        {
             _logger.LogError("Connection ID cannot be null or empty when getting connection state.");
             throw new ArgumentNullException(nameof(connectionId));
        }

         _logger.LogInformation("Getting state for connection {ConnectionId}...", connectionId);

        try
        {
            // GET /external/connections/{connectionId}?$select=state
            var connection = await _graphClient.External.Connections[connectionId]
                .GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Select = new string[] { "state" };
                });

            _logger.LogInformation("State for connection {ConnectionId} is {State}.", connectionId, connection?.State);
            return connection?.State;
        }
        catch (ODataError odataError)
        {
            _logger.LogError(odataError, "OData Error getting state for connection {ConnectionId}: {StatusCode} {Code} {Message}", connectionId, odataError.ResponseStatusCode, odataError.Error?.Code, odataError.Error?.Message);
            return null;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error getting state for connection {ConnectionId}: {Message}", connectionId, ex.Message);
            return null;
        }
    }

    // Method to delete a connection
    public static async Task<bool> DeleteConnectionAsync(string connectionId)
    {
        if (_graphClient == null || _logger == null) throw new InvalidOperationException("Graph client or logger not initialized.");
        if (string.IsNullOrEmpty(connectionId))
        {
            _logger.LogError("Connection ID cannot be null or empty when deleting connection.");
            throw new ArgumentNullException(nameof(connectionId));
        }

        _logger.LogInformation("Deleting connection with ID {ConnectionId}...", connectionId);

        try
        {
            // DELETE /external/connections/{connectionId}
            await _graphClient.External.Connections[connectionId].DeleteAsync();

            _logger.LogInformation("Connection {ConnectionId} deleted successfully.", connectionId);
            return true;
        }
        catch (ODataError odataError)
        {
            // Handle case where connection might not be found (e.g., already deleted)
            if (odataError.ResponseStatusCode == 404)
            {
                 _logger.LogWarning("Connection with ID {ConnectionId} not found during delete attempt. Already deleted?", connectionId);
                 return true; // Consider deletion successful if not found
            }
            _logger.LogError(odataError, "OData Error deleting connection {ConnectionId}: {StatusCode} {Code} {Message}", connectionId, odataError.ResponseStatusCode, odataError.Error?.Code, odataError.Error?.Message);
            return false;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error deleting connection {ConnectionId}: {Message}", connectionId, ex.Message);
            return false;
        }
    }

    // Method to register the schema for a connection
    public static async Task<bool> RegisterSchemaAsync(string connectionId, Schema schema)
    {
        if (_graphClient == null || _logger == null) throw new InvalidOperationException("Graph client or logger not initialized.");
        if (string.IsNullOrEmpty(connectionId))
        {
             _logger.LogError("Connection ID cannot be null or empty when registering schema.");
             throw new ArgumentNullException(nameof(connectionId));
        }
        ArgumentNullException.ThrowIfNull(schema);

        _logger.LogInformation("Registering schema for connection {ConnectionId}...", connectionId);

        try
        {
            // Ensure the state is 'draft' before attempting to register
            var state = await GetConnectionStateAsync(connectionId);
            if (state != ConnectionState.Draft && state != null) // Allow registration if state couldn't be determined (null) or is draft
            {
                 _logger.LogWarning("Schema cannot be registered for connection {ConnectionId} because its state is {State}. It must be 'draft'.", connectionId, state);
                 // Optionally, attempt to reset to draft if needed, but for now, just return false.
                 return false;
            }
            else if (state == null)
            {
                _logger.LogWarning("Could not determine state for connection {ConnectionId}. Proceeding with schema registration attempt.", connectionId);
            }

            // PATCH /external/connections/{connectionId}/schema
             await _graphClient.External.Connections[connectionId].Schema
                .PatchAsync(schema); // Use PatchAsync for schema update/registration

            _logger.LogInformation("Schema registration command sent successfully for connection {ConnectionId}. Waiting for provisioning...", connectionId);

            // Wait for the schema provisioning to complete (optional, but recommended)
            return await WaitForSchemaProvisioningAsync(connectionId);

        }
        catch (ODataError odataError)
        {
            _logger.LogError(odataError, "OData Error registering schema for connection {ConnectionId}: {StatusCode} {Code} {Message}", connectionId, odataError.ResponseStatusCode, odataError.Error?.Code, odataError.Error?.Message);
            return false;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error registering schema for connection {ConnectionId}: {Message}", connectionId, ex.Message);
            return false;
        }
    }

    // Helper method to wait for schema provisioning
    private static async Task<bool> WaitForSchemaProvisioningAsync(string connectionId, int timeoutSeconds = 120, int delaySeconds = 5)
    {
        if (_graphClient == null || _logger == null) throw new InvalidOperationException("Graph client or logger not initialized.");
        _logger.LogInformation("Waiting up to {TimeoutSeconds} seconds for schema provisioning on connection {ConnectionId}...", timeoutSeconds, connectionId);

        var startTime = DateTime.UtcNow;
        while (DateTime.UtcNow - startTime < TimeSpan.FromSeconds(timeoutSeconds))
        {
            try
            {
                var connection = await _graphClient.External.Connections[connectionId].GetAsync(config => config.QueryParameters.Select = new[] { "state" });

                if (connection?.State == ConnectionState.Ready)
                {
                    _logger.LogInformation("Schema for connection {ConnectionId} provisioned successfully (State: Ready).", connectionId);
                    return true;
                }
                 if (connection?.State == ConnectionState.LimitExceeded || connection?.State == ConnectionState.UnknownFutureValue)
                {
                     _logger.LogError("Schema provisioning failed for connection {ConnectionId}. State: {State}", connectionId, connection.State);
                    return false;
                }

                _logger.LogDebug("Connection {ConnectionId} state is {State}. Waiting {DelaySeconds} seconds...", connectionId, connection?.State, delaySeconds);
                await Task.Delay(TimeSpan.FromSeconds(delaySeconds));
            }
             catch (ODataError odataError)
            {
                 _logger.LogError(odataError, "OData error while checking schema provisioning status for {ConnectionId}: {StatusCode} {Code} {Message}", connectionId, odataError.ResponseStatusCode, odataError.Error?.Code, odataError.Error?.Message);
                 // Decide if we should retry or fail based on the error (e.g., maybe fail on 404)
                 await Task.Delay(TimeSpan.FromSeconds(delaySeconds)); // Wait before retrying
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Unexpected error while waiting for schema provisioning for {ConnectionId}. Aborting wait.", connectionId);
                return false; // Stop waiting on unexpected errors
            }
        }

        _logger.LogError("Schema provisioning for connection {ConnectionId} timed out after {TimeoutSeconds} seconds.", connectionId, timeoutSeconds);
        return false;
    }


    // Method to get the current schema
    public static async Task<Schema?> GetSchemaAsync(string connectionId)
    {
        if (_graphClient == null || _logger == null) throw new InvalidOperationException("Graph client or logger not initialized.");
        if (string.IsNullOrEmpty(connectionId))
        {
            _logger.LogError("Connection ID cannot be null or empty when getting schema.");
            throw new ArgumentNullException(nameof(connectionId));
        }

        _logger.LogInformation("Getting schema for connection {ConnectionId}...", connectionId);

        try
        {
            // GET /external/connections/{connectionId}/schema
            var schema = await _graphClient.External.Connections[connectionId].Schema.GetAsync();
            _logger.LogInformation("Schema retrieved successfully for connection {ConnectionId}.", connectionId);
            return schema;
        }
        catch (ODataError odataError)
        {
             // Check if the error is 'Resource not found' (might mean schema not registered yet)
            if (odataError.ResponseStatusCode == 404)
            {
                 _logger.LogWarning("Schema not found for connection {ConnectionId}. It might not be registered yet.", connectionId);
                 return null;
            }
            _logger.LogError(odataError, "OData Error getting schema for {ConnectionId}: {StatusCode} {Code} {Message}", connectionId, odataError.ResponseStatusCode, odataError.Error?.Code, odataError.Error?.Message);
            return null;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error getting schema for {ConnectionId}: {Message}", connectionId, ex.Message);
            return null;
        }
    }

    // Method to add or update a single item
    public static async Task<bool> AddOrUpdateItemAsync(string connectionId, ExternalItem item)
    {
         if (_graphClient == null || _logger == null) throw new InvalidOperationException("Graph client or logger not initialized.");
         if (string.IsNullOrEmpty(connectionId))
         {
             _logger.LogError("Connection ID cannot be null or empty when adding/updating item.");
             throw new ArgumentNullException(nameof(connectionId));
         }
         ArgumentNullException.ThrowIfNull(item);
         if (string.IsNullOrEmpty(item.Id))
         {
             _logger.LogError("ExternalItem must have an ID.");
             throw new ArgumentException("ExternalItem must have an ID.", nameof(item));
         }

        _logger.LogInformation("Adding/updating item {ItemId} in connection {ConnectionId}...", item.Id, connectionId);

        try
        {
            // PUT /external/connections/{connectionId}/items/{itemId}
            await _graphClient.External.Connections[connectionId].Items[item.Id].PutAsync(item);
            _logger.LogInformation("Successfully added/updated item {ItemId} in connection {ConnectionId}.", item.Id, connectionId);
            return true;
        }
        catch (ODataError odataError)
        {
             _logger.LogError(odataError, "OData Error adding/updating item {ItemId} in {ConnectionId}: {StatusCode} {Code} {Message}", item.Id, connectionId, odataError.ResponseStatusCode, odataError.Error?.Code, odataError.Error?.Message);
            return false;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error adding/updating item {ItemId} in {ConnectionId}: {Message}", item.Id, connectionId, ex.Message);
            return false;
        }
    }

     // Method to add or update items in bulk (using Graph SDK's batching isn't straightforward for external items, handle sequentially for now)
    public static async Task<int> AddOrUpdateItemsAsync(string connectionId, List<ExternalItem> items, int batchSize = 10) // Added batchSize, though not using Graph Batching yet
    {
        if (_graphClient == null || _logger == null) throw new InvalidOperationException("Graph client or logger not initialized.");
        if (string.IsNullOrEmpty(connectionId))
        {
            _logger.LogError("Connection ID cannot be null or empty when adding/updating items.");
            throw new ArgumentNullException(nameof(connectionId));
        }
        ArgumentNullException.ThrowIfNull(items);

        _logger.LogInformation("Starting bulk add/update for {ItemCount} items in connection {ConnectionId}...", items.Count, connectionId);
        int successCount = 0;
        int failureCount = 0;

        // Process items sequentially for simplicity.
        // TODO: Implement proper batching if performance becomes an issue and SDK/API supports it well.
        for (int i = 0; i < items.Count; i++)
        {
            var item = items[i];
             if (string.IsNullOrEmpty(item.Id))
            {
                _logger.LogWarning("Skipping item at index {Index} because it has no ID.", i);
                failureCount++;
                continue;
            }

            if (await AddOrUpdateItemAsync(connectionId, item))
            {
                successCount++;
                // Optional: Add a small delay between requests to avoid throttling if needed
                // await Task.Delay(50);
            }
            else
            {
                 failureCount++;
                 _logger.LogWarning("Failed to add/update item {ItemId} (Index {Index}). See previous error for details.", item.Id, i);
                 // Consider adding retry logic here for transient errors
            }
        }

        _logger.LogInformation("Bulk add/update finished for connection {ConnectionId}. Success: {SuccessCount}, Failed: {FailureCount}",
            connectionId, successCount, failureCount);

        return successCount; // Return the number of successfully processed items
    }


    // Method to delete a single item
    public static async Task<bool> DeleteItemAsync(string connectionId, string itemId)
    {
         if (_graphClient == null || _logger == null) throw new InvalidOperationException("Graph client or logger not initialized.");
         if (string.IsNullOrEmpty(connectionId) || string.IsNullOrEmpty(itemId))
         {
             _logger.LogError("Connection ID and Item ID cannot be null or empty when deleting item.");
             throw new ArgumentNullException(string.IsNullOrEmpty(connectionId) ? nameof(connectionId) : nameof(itemId));
         }

        _logger.LogInformation("Deleting item {ItemId} from connection {ConnectionId}...", itemId, connectionId);

        try
        {
            // DELETE /external/connections/{connectionId}/items/{itemId}
            await _graphClient.External.Connections[connectionId].Items[itemId].DeleteAsync();
            _logger.LogInformation("Successfully deleted item {ItemId} from connection {ConnectionId}.", itemId, connectionId);
            return true;
        }
        catch (ODataError odataError)
        {
             // Handle 404 Not Found gracefully
             if (odataError.ResponseStatusCode == 404)
             {
                 _logger.LogWarning("Item {ItemId} not found in connection {ConnectionId} during delete. Assuming already deleted.", itemId, connectionId);
                 return true; // Consider deletion successful if not found
             }
            _logger.LogError(odataError, "OData Error deleting item {ItemId} from {ConnectionId}: {StatusCode} {Code} {Message}", itemId, connectionId, odataError.ResponseStatusCode, odataError.Error?.Code, odataError.Error?.Message);
            return false;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error deleting item {ItemId} from {ConnectionId}: {Message}", itemId, connectionId, ex.Message);
            return false;
        }
    }

    // Method to delete multiple items (handle sequentially for now)
    public static async Task<int> DeleteItemsAsync(string connectionId, List<string> itemIds)
    {
        if (_graphClient == null || _logger == null) throw new InvalidOperationException("Graph client or logger not initialized.");
        if (string.IsNullOrEmpty(connectionId))
        {
             _logger.LogError("Connection ID cannot be null or empty when deleting items.");
            throw new ArgumentNullException(nameof(connectionId));
        }
         ArgumentNullException.ThrowIfNull(itemIds);

        _logger.LogInformation("Starting bulk delete for {ItemCount} items in connection {ConnectionId}...", itemIds.Count, connectionId);
        int successCount = 0;
        int failureCount = 0;

        foreach (var itemId in itemIds)
        {
             if (string.IsNullOrEmpty(itemId))
             {
                 _logger.LogWarning("Skipping delete for a null or empty item ID.");
                 failureCount++;
                 continue;
             }

            if (await DeleteItemAsync(connectionId, itemId))
            {
                successCount++;
            }
            else
            {
                failureCount++;
                _logger.LogWarning("Failed to delete item {ItemId}. See previous error for details.", itemId);
            }
        }

         _logger.LogInformation("Bulk delete finished for connection {ConnectionId}. Success: {SuccessCount}, Failed: {FailureCount}",
            connectionId, successCount, failureCount);

        return successCount;
    }


    // Method to get all item IDs currently in the connection (potential for pagination needed)
    public static async Task<List<string>?> GetConnectionItemIdsAsync(string connectionId)
    {
        if (_graphClient == null || _logger == null) throw new InvalidOperationException("Graph client or logger not initialized.");
         if (string.IsNullOrEmpty(connectionId))
         {
             _logger.LogError("Connection ID cannot be null or empty when getting item IDs.");
             throw new ArgumentNullException(nameof(connectionId));
         }

        _logger.LogInformation("Getting all item IDs for connection {ConnectionId}...", connectionId);
        var allItemIds = new List<string>();

        try
        {
            // GET /external/connections/{connectionId}/items?$select=id
            // Use iteration to handle potential pagination
             var itemsPage = await _graphClient.External.Connections[connectionId].Items
                 .GetAsync(requestConfiguration =>
                 {
                     requestConfiguration.QueryParameters.Select = new[] { "id" };
                     // Adjust top based on API limits / performance needs
                     requestConfiguration.QueryParameters.Top = 100;
                 });

            while (itemsPage?.Value != null)
            {
                allItemIds.AddRange(itemsPage.Value.Select(item => item.Id).Where(id => !string.IsNullOrEmpty(id))!); // Add non-null IDs

                // Check if there is a next page
                if (!string.IsNullOrEmpty(itemsPage.OdataNextLink))
                {
                     _logger.LogDebug("Fetching next page of item IDs for connection {ConnectionId}...", connectionId);
                     itemsPage = await _graphClient.External.Connections[connectionId].Items
                         .WithUrl(itemsPage.OdataNextLink) // Use the next link URL
                         .GetAsync();
                }
                else
                {
                    break; // No more pages
                }
            }

            _logger.LogInformation("Retrieved {ItemCount} item IDs for connection {ConnectionId}.", allItemIds.Count, connectionId);
            return allItemIds;
        }
        catch (ODataError odataError)
        {
            _logger.LogError(odataError, "OData Error getting item IDs for {ConnectionId}: {StatusCode} {Code} {Message}", connectionId, odataError.ResponseStatusCode, odataError.Error?.Code, odataError.Error?.Message);
            return null; // Indicate error
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error getting item IDs for {ConnectionId}: {Message}", connectionId, ex.Message);
            return null; // Indicate error
        }
    }
} 