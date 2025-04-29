using PurviewSearchConnector;
using Microsoft.Graph.Models.ExternalConnectors;
using Microsoft.Graph.Models.ODataErrors;
using Azure.Analytics.Purview.DataMap;

Console.WriteLine("Purview Glossary Search Connector\n");

var settings = Settings.LoadSettings();

// Wrap initialization in try-catch
try
{
    // Initialize Graph
    Console.WriteLine("Initializing Microsoft Graph client...");
    GraphHelper.InitializeGraph(settings);
    Console.WriteLine("Graph client initialized.");

    // Initialize Purview
    Console.WriteLine("Initializing Purview Data Map client...");
    PurviewHelper.InitializePurview(settings);
    // PurviewHelper already prints success/failure

    // Test Purview Connection
    await PurviewHelper.TestPurviewConnectionAsync();
}
catch (Exception ex)
{
    Console.WriteLine($"\nInitialization failed: {ex.Message}");
    Console.WriteLine("Exiting application.");
    return; // Exit if initialization fails
}

ExternalConnection? currentConnection = null;

int choice = -1;

while (choice != 0)
{
    Console.WriteLine($"Current connection: {(currentConnection == null ? "NONE" : currentConnection.Id)} - {currentConnection?.Name}\n"); // Added ID and Name
    Console.WriteLine("Please choose one of the following options:");
    Console.WriteLine("0. Exit");
    Console.WriteLine("1. Create a connection");
    Console.WriteLine("2. Select an existing connection");
    Console.WriteLine("3. Delete current connection");
    Console.WriteLine("4. Register schema for current connection");
    Console.WriteLine("5. View schema for current connection");
    Console.WriteLine("6. Push UPDATED items to current connection");
    Console.WriteLine("7. Push ALL items to current connection");
    Console.WriteLine("8. Delete ALL items from current connection");
    Console.WriteLine("9. Check Connection State");
    Console.Write("Selection: ");

    try
    {
        choice = int.Parse(Console.ReadLine() ?? string.Empty);
    }
    catch (FormatException)
    {
        // Set to invalid value
        choice = -1;
    }

    try 
    { 
        switch(choice)
        {
            case 0:
                // Exit the program
                Console.WriteLine("Goodbye...");
                break;
            case 1:
                currentConnection = await CreateConnectionAsync();
                break;
            case 2:
                currentConnection = await SelectExistingConnectionAsync();
                break;
            case 3:
                await DeleteCurrentConnectionAsync(currentConnection);
                currentConnection = null;
                break;
            case 4:
                await RegisterSchemaAsync(currentConnection);
                break;
            case 5:
                 await ViewSchemaAsync(currentConnection);
                 break;
            case 6:
                await UpdateItemsAsync(currentConnection, settings.TenantID, true);
                break;
            case 7:
                await UpdateItemsAsync(currentConnection, settings.TenantID, false);
                break;
            case 8:
                await DeleteAllItemsAsync(currentConnection);
                break;
            case 9:
                await CheckConnectionStateAsync(currentConnection);
                break;
            default:
                Console.WriteLine("Invalid choice! Please try again.");
                break;
        }
    }
    catch (ODataError odataError)
    {
        Console.WriteLine($"ERROR: {odataError.Error?.Code} {odataError.Error?.Message}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"ERROR: {ex.Message}");
    }
}

// Placeholder functions - To be implemented
async Task<ExternalConnection?> CreateConnectionAsync()
{
    // Prompt user for connection details
    Console.Write("Enter a unique ID for the connection (alphanumeric, 3-32 chars): ");
    string? connectionId = Console.ReadLine();
    // Basic validation (add more robust validation if needed)
    if (string.IsNullOrWhiteSpace(connectionId) || connectionId.Length < 3 || connectionId.Length > 32 || !System.Text.RegularExpressions.Regex.IsMatch(connectionId, "^[a-zA-Z0-9]*$"))
    {
        Console.WriteLine("Invalid ID. Must be 3-32 alphanumeric characters.");
        return null;
    }

    // Convert ID to lowercase before using
    connectionId = connectionId.ToLowerInvariant();
    Console.WriteLine($"Using connection ID: {connectionId}"); // Show the user the enforced ID

    Console.Write("Enter a name for the connection: ");
    string? connectionName = Console.ReadLine();
    if (string.IsNullOrWhiteSpace(connectionName))
    {
        Console.WriteLine("Connection name cannot be empty.");
        return null;
    }

    Console.Write("Enter a description for the connection: ");
    string? connectionDescription = Console.ReadLine();
    if (string.IsNullOrWhiteSpace(connectionDescription))
    {
        Console.WriteLine("Connection description cannot be empty.");
        return null;
    }

    // Call GraphHelper to create the connection
    // Need to add using Microsoft.Graph.Models.ExternalConnectors at the top
    return await GraphHelper.CreateConnectionAsync(connectionId, connectionName, connectionDescription);
}

async Task<ExternalConnection?> SelectExistingConnectionAsync()
{
    Console.WriteLine("Getting existing connections...");
    var connections = await GraphHelper.GetConnectionsAsync();

    if (connections == null)
    {
        Console.WriteLine("Error retrieving connections.");
        return null;
    }

    if (connections.Count == 0)
    {
        Console.WriteLine("No connections found for this application.");
        return null;
    }

    Console.WriteLine("Available connections:");
    for (int i = 0; i < connections.Count; i++)
    {
        Console.WriteLine($" {i + 1}. {connections[i].Name} ({connections[i].Id})");
    }

    ExternalConnection? selectedConnection = null;
    while (selectedConnection == null)
    {
        Console.Write("Select connection number: ");
        try
        {
            if (int.TryParse(Console.ReadLine(), out int selection) && selection > 0 && selection <= connections.Count)
            {
                selectedConnection = connections[selection - 1];
            }
            else
            {
                Console.WriteLine("Invalid selection.");
            }
        }
        catch
        {
            Console.WriteLine("Invalid input.");
        }
    }

    Console.WriteLine($"Selected connection: {selectedConnection.Name}");
    return selectedConnection;
}

async Task DeleteCurrentConnectionAsync(ExternalConnection? connection)
{
    if (connection == null)
    {
        Console.WriteLine("No connection selected.");
        return;
    }

    Console.WriteLine($"WARNING: You are about to delete connection '{connection.Name}' ({connection.Id}).");
    Console.Write("This cannot be undone. Are you sure? (y/N): ");
    string? confirmation = Console.ReadLine();

    if (confirmation?.Trim().ToLower() == "y")
    {
        bool success = await GraphHelper.DeleteConnectionAsync(connection.Id ?? string.Empty); 
        if (success)
        {
            Console.WriteLine("Connection deletion process completed.");
            // Note: currentConnection will be set to null in the main loop after this returns
        }
        else
        {
            Console.WriteLine("Connection deletion failed. See errors above.");
        }
    }
    else
    {
        Console.WriteLine("Deletion cancelled.");
    }
}

// Add RegisterSchemaAsync placeholder implementation
async Task RegisterSchemaAsync(ExternalConnection? connection)
{
    if (connection == null || string.IsNullOrEmpty(connection.Id))
    {
        Console.WriteLine("No connection selected. Please create or select a connection first.");
        return;
    }

    Console.WriteLine($"Registering schema for connection '{connection.Name}' ({connection.Id})...");
    await GraphHelper.RegisterSchemaAsync(connection.Id);
    // Output messages handled by GraphHelper method
}

// Implement ViewSchemaAsync function
async Task ViewSchemaAsync(ExternalConnection? connection)
{
    if (connection == null || string.IsNullOrEmpty(connection.Id))
    {
        Console.WriteLine("No connection selected. Please create or select a connection first.");
        return;
    }

    Console.WriteLine($"Getting schema for connection '{connection.Name}' ({connection.Id})...");
    var schema = await GraphHelper.GetSchemaAsync(connection.Id);

    if (schema != null)
    {
        Console.WriteLine("\n--- Schema Details ---");
        Console.WriteLine($"BaseType: {schema.BaseType}");
        Console.WriteLine("Properties:");
        if (schema.Properties != null && schema.Properties.Count > 0)
        {
            foreach (var prop in schema.Properties)
            {
                Console.WriteLine($"  - Name: {prop.Name}");
                Console.WriteLine($"    Type: {prop.Type}");
                Console.WriteLine($"    IsSearchable: {prop.IsSearchable ?? false}");
                Console.WriteLine($"    IsRetrievable: {prop.IsRetrievable ?? false}");
                Console.WriteLine($"    IsRefinable: {prop.IsRefinable ?? false}");
                Console.WriteLine($"    IsQueryable: {prop.IsQueryable ?? false}");
                // Add other attributes like Aliases if needed
            }
        }
        else
        {
            Console.WriteLine("  (No properties defined in schema)");
        }
        Console.WriteLine("----------------------");

        // Also check the connection state itself
        // We could add a separate GraphHelper method to get just the connection status
        // For now, let's just remind the user to check status manually or implement later
        Console.WriteLine("\nReminder: Check connection status separately (e.g., via Graph Explorer or implement a status check feature). Schema registration needs the connection state to be 'ready'.");
    }
    else
    {
        Console.WriteLine("Could not retrieve schema details. See errors above.");
    }
}

// --- Timestamp Management ---
const string TimestampFile = "lastSync.txt";

DateTimeOffset GetLastSyncTime()
{
    try
    {
        if (File.Exists(TimestampFile))
        {
            string timestampStr = File.ReadAllText(TimestampFile);
            if (DateTimeOffset.TryParse(timestampStr, out DateTimeOffset lastSync))
            {
                Console.WriteLine($"Last sync time loaded: {lastSync}");
                return lastSync;
            }
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Warning: Could not read last sync time: {ex.Message}");
    }
    // Default to a very old time if file doesn't exist or is invalid
    return DateTimeOffset.MinValue; 
}

void SaveLastSyncTime(DateTimeOffset syncTime)
{
    try
    {
        File.WriteAllText(TimestampFile, syncTime.ToString("o")); // ISO 8601 format
        Console.WriteLine($"Current sync time saved: {syncTime}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error saving sync time: {ex.Message}");
    }
}

// --- Item Ingestion Orchestration ---

async Task UpdateItemsAsync(ExternalConnection? connection, string? tenantId, bool pushModifiedOnly)
{
    if (connection == null || string.IsNullOrEmpty(connection.Id) || string.IsNullOrEmpty(tenantId))
    {
        Console.WriteLine("Connection not selected or tenant ID missing. Cannot push items.");
        return;
    }

    Console.WriteLine($"\nStarting item push ({(pushModifiedOnly ? "INCREMENTAL" : "FULL")}) for connection '{connection.Name}'...");

    // Ask user if they want to target a specific glossary
    string targetGlossaryId = null;
    Console.Write("Process ALL glossaries? (Y/n - 'n' to specify one): ");
    string? processAllChoice = Console.ReadLine();
    if (processAllChoice?.Trim().ToLower() == "n")
    {
        Console.Write("Enter the GUID of the glossary to process: ");
        targetGlossaryId = Console.ReadLine();
        if (string.IsNullOrWhiteSpace(targetGlossaryId))
        {
            Console.WriteLine("Invalid Glossary GUID entered. Aborting push.");
            return;
        }
        Console.WriteLine($"Targeting Glossary ID: {targetGlossaryId}");
    }
    else
    {
        Console.WriteLine("Processing ALL glossaries.");
    }

    DateTimeOffset lastSyncTime = DateTimeOffset.MinValue;
    if (pushModifiedOnly)
    {
        lastSyncTime = GetLastSyncTime();
    }
    DateTimeOffset currentSyncStartTime = DateTimeOffset.UtcNow; 

    // 1. Fetch Data from Purview - Returns list of tuples
    List<(AtlasGlossaryTerm Term, string GlossaryName)> termData = await PurviewHelper.GetGlossaryTermsDataAsync(targetGlossaryId);

    if (termData == null || termData.Count == 0)
    {
        Console.WriteLine("No terms retrieved from Purview. Aborting push.");
        return;
    }

    Console.WriteLine($"Retrieved {termData.Count} total terms from Purview.");

    // 2. Filter, Map, and Push to Graph
    int itemsProcessed = 0;
    int itemsSkipped = 0;
    int itemsPushed = 0;

    // Iterate through the list of tuples
    foreach (var (term, glossaryName) in termData) // Deconstruct tuple
    {
        itemsProcessed++;
        if (string.IsNullOrEmpty(term.Guid)) 
        {
            Console.WriteLine("Skipping term with missing GUID.");
            itemsSkipped++;
            continue;
        }

        // Check modification time 
        DateTimeOffset? termLastModified = term.UpdateTime.HasValue 
            ? DateTimeOffset.FromUnixTimeMilliseconds(term.UpdateTime.Value) 
            : null;
        if (pushModifiedOnly && termLastModified.HasValue && termLastModified.Value <= lastSyncTime)
        {
            itemsSkipped++;
            continue; 
        }

        // Map Purview term to Graph ExternalItem, passing glossaryName
        ExternalItem graphItem;
        try
        {
            graphItem = PurviewHelper.MapPurviewTermToExternalItem(term, glossaryName, tenantId!); // Pass tenantId safely
        }
        catch (Exception mapEx)
        {
            Console.WriteLine($"ERROR mapping term {term.Guid}: {mapEx.Message}");
            itemsSkipped++;
            continue;
        }

        // Push item to Graph connection and check result
        bool success = await GraphHelper.AddOrUpdateItemAsync(connection.Id, graphItem);
        if (success)
        {
            itemsPushed++;
        }
        else
        {
             // Failure already logged by GraphHelper
             itemsSkipped++; // Count failures as skipped for summary
        }
    }

    Console.WriteLine("\n--- Push Summary ---");
    Console.WriteLine($"Items Processed: {itemsProcessed}");
    Console.WriteLine($"Items Skipped (errors/unmodified): {itemsSkipped}");
    Console.WriteLine($"Items Pushed (add/update): {itemsPushed}");
    Console.WriteLine("------------------");
    if (itemsPushed > 0 || !pushModifiedOnly) 
    {
        SaveLastSyncTime(currentSyncStartTime);
    }
    else if (itemsSkipped == itemsProcessed)
    {
         Console.WriteLine("No new or modified items found to push.");
    }
    Console.WriteLine("Item push process finished.");
}

// --- Delete All Items Function ---

async Task DeleteAllItemsAsync(ExternalConnection? connection)
{
    if (connection == null || string.IsNullOrEmpty(connection.Id))
    {
        Console.WriteLine("No connection selected. Please create or select a connection first.");
        return;
    }

    Console.WriteLine($"WARNING: This will attempt to delete ALL items from connection '{connection.Name}' ({connection.Id}).");
    Console.Write("Are you sure? (y/N): ");
    string? confirmation = Console.ReadLine();

    if (confirmation?.Trim().ToLower() != "y")
    {
        Console.WriteLine("Deletion cancelled.");
        return;
    }

    // 1. Get all item IDs from the connection
    List<string>? itemIds = await GraphHelper.GetConnectionItemIdsAsync(connection.Id);

    if (itemIds == null)
    {
        Console.WriteLine("Error retrieving item IDs. Cannot proceed with deletion.");
        return;
    }

    if (itemIds.Count == 0)
    {
        Console.WriteLine("No items found in the connection to delete.");
        return;
    }

    Console.WriteLine($"Found {itemIds.Count} items. Proceeding with deletion...");

    // 2. Loop and delete each item
    int deleteSuccessCount = 0;
    int deleteFailCount = 0;
    foreach (string itemId in itemIds)
    {
        bool success = await GraphHelper.DeleteItemAsync(connection.Id, itemId);
        if (success) 
        {
            deleteSuccessCount++;
        }
        else
        {
            deleteFailCount++;
        }
        // Optional: Add delay if hitting throttling limits
        // await Task.Delay(50);
    }

    Console.WriteLine("\n--- Deletion Summary ---");
    Console.WriteLine($"Deletion requests submitted: {deleteSuccessCount}");
    Console.WriteLine($"Deletion requests failed: {deleteFailCount}");
    Console.WriteLine("----------------------");
}

// --- Check Connection State Function ---

async Task CheckConnectionStateAsync(ExternalConnection? connection)
{
    if (connection == null || string.IsNullOrEmpty(connection.Id))
    {
        Console.WriteLine("No connection selected.");
        return;
    }

    Console.WriteLine($"Checking state for connection '{connection.Name}' ({connection.Id})...");
    ConnectionState? state = await GraphHelper.GetConnectionStateAsync(connection.Id);

    if (state != null)
    {
        Console.WriteLine($"Connection state: {state}");
        if (state == ConnectionState.Ready)
        {
            Console.WriteLine("Connection is ready for item ingestion/querying.");
        }
        else if (state == ConnectionState.Draft)
        {
             Console.WriteLine("Connection is still provisioning (Draft state). Wait and check again.");
        }
        else // Handles Error, LimitExceeded, UnknownFutureValue
        {
             Console.WriteLine($"Connection is in a non-ready state ({state}). Check for errors or configuration issues.");
        }
    }
    else
    {
        Console.WriteLine("Could not retrieve connection state.");
    }
}
