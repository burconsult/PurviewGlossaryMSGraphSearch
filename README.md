# Purview Glossary to MS Graph Search Synchronization System

This document describes the components and workflow of the system. It makes glossary terms from Microsoft Purview Glossaries (Classic type) available for search in Microsoft 365 by synchronizing them to MS Graph external items.

## Overview

The system consists of two main components:
1.  **Connector Console App  (`PurviewSearchConnector`)**: Initializes the MS Graph connector and can be used to run tests and initialize partial or full Purview Glossary terms synchronization.
2.  **Azure Function App (`PurviewSyncFunc`)**: Listens for change events and synchronizes them accordingly.

## Components

### Connector App

The `PurviewSearchConnector` is a .NET console application designed to bridge Microsoft Purview and Microsoft Graph Search. Its primary function is to create a custom Microsoft Graph Connector that fetches glossary terms (classic types) from a specified Purview account and indexes them as `ExternalItem` objects within Microsoft Graph. This makes the Purview glossary terms discoverable through Microsoft Search across the M365 ecosystem.

**Functionality:**

*   **Connects to Purview:** Uses the Azure SDK (`Azure.Analytics.Purview.DataMap`) and Azure AD application credentials (`ClientID`, `ClientSecret`, `TenantID`) specified in `config.ini` to connect to the Purview Data Map API for a given Purview account (`AccountName`).
*   **Fetches Glossary Terms:** Retrieves terms from all glossaries within the specified Purview account or optionally from a specific glossary. (`PurviewHelper.cs`)
*   **Manages Graph Connector:** Interacts with the Microsoft Graph API (`Microsoft.Graph` SDK) to:
    *   Create, view, select, and delete `ExternalConnection` resources.
    *   Define and register a `Schema` for the connection, specifying how Purview term properties (like name, description, acronym, contacts, related terms, definition) map to searchable Graph properties.
    *   Manage the state of the connection. (`GraphHelper.cs`)
*   **Indexes Items in Graph:**
    *   Maps fetched Purview terms to `ExternalItem` objects using the defined schema. It cleans potential HTML from descriptions during mapping. (`PurviewHelper.cs`)
    *   Pushes these `ExternalItem` objects to the selected Graph connection, either performing a full sync (all terms), a partial sync (specific glossary) or an incremental sync (only terms updated since the last run, based on `lastSync.txt`).
    *   Can also delete all indexed items from a connection. (`Program.cs`, `GraphHelper.cs`)
*   **Configuration:** Requires Azure AD app registration details and the Purview account name to be configured in `config.ini`. The `Settings.cs` file handles loading these settings.
*   **Operation:** Runs as an interactive console application, providing a menu to perform the various tasks (connection management, schema registration, item synchronization). It tracks the last sync timestamp in `lastSync.txt` for incremental updates.

### Azure Function App (`PurviewSyncFunc`)

The `PurviewToGraphSyncFunction` is an Azure Function App designed to automatically synchronize glossary terms from Microsoft Purview to a Microsoft Graph connection on a scheduled basis.

**Functionality:**

*   **Trigger:** Uses a Timer Trigger (`[TimerTrigger("%TimerSchedule%")]`). The schedule (e.g., every 30 minutes) is defined by the `TimerSchedule` application setting.
*   **Connects to Purview & Graph:** Similar to the `PurviewSearchConnector`, it uses Azure AD application credentials (`ClientID`, `ClientSecret`, `TenantID`) stored as application settings to initialize connections to the Purview Data Map API (`PurviewEndpoint` setting) and the Microsoft Graph API. It utilizes adapted versions of `PurviewHelper.cs` and `GraphHelper.cs`, incorporating `ILogger` for logging within the Azure Functions environment.
*   **Incremental Synchronization:** 
    *   Reads a timestamp from a blob (`lastSyncTimestamp.txt`) stored in an Azure Blob Storage container (`TimestampContainerName` setting, connection string from `AzureWebJobsStorage` setting) to determine the last successful sync time.
    *   Fetches only those Purview glossary terms that have been created or modified since the `lastSyncTime`. If no timestamp exists (first run or previous error), it performs a full sync.
*   **Maps and Pushes Items:**
    *   Maps the fetched Purview terms to `ExternalItem` objects using the same logic as the connector app (`PurviewHelper.MapPurviewTermToExternalItem`).
    *   Pushes the `ExternalItem` objects to a pre-defined Microsoft Graph Connection specified by the `GraphConnectionId` application setting (`GraphHelper.AddOrUpdateItemsAsync`).
*   **Timestamp Update:** Upon successful completion of fetching, mapping, and pushing items to Graph, it updates the `lastSyncTimestamp.txt` blob with the timestamp of when the current function execution began.
*   **Configuration:** Relies entirely on Azure Function Application Settings for configuration (Azure AD details, Purview endpoint, Graph Connection ID, Storage Account details, Timer Schedule).
*   **Operation:** Runs automatically based on the timer schedule. It does *not* create or manage the Graph connection or schema; it assumes the connection (identified by `GraphConnectionId`) and its schema have already been created (likely using the `PurviewSearchConnector` console app).

## Workflow

1.  **Initial Setup (Manual):**
    *   An Azure AD Application is registered and granted necessary permissions for both Purview and Microsoft Graph.
    *   The `PurviewSearchConnector` console application is configured (`config.ini`) and run locally.
    *   Using the console app, an administrator creates a new Microsoft Graph `ExternalConnection`.
    *   The schema defining how Purview terms map to Graph `ExternalItem` properties is registered for the connection using the console app.
    *   The ID of the created Graph connection is noted.
    *   The `PurviewToGraphSyncFunction` Azure Function App is deployed.
    *   The Function App's Application Settings are configured with the Azure AD credentials, Purview endpoint, the created Graph Connection ID, Azure Storage details, and the desired timer schedule.
    *   Optionally, an initial full synchronization of all Purview glossary terms can be pushed using the console app (menu option 7).

2.  **Scheduled Synchronization (Automated):**
    *   The `PurviewToGraphSyncFunction` Azure Function triggers based on its `TimerSchedule`.
    *   The function reads the `lastSyncTimestamp.txt` blob from Azure Storage to find the last successful run time.
    *   It queries the Purview Data Map API for glossary terms created or updated since the last sync time.
    *   Fetched terms are mapped to `ExternalItem` objects according to the pre-registered schema.
    *   These `ExternalItem` objects are pushed (added or updated) to the specified Graph `ExternalConnection`.
    *   If the process is successful, the function updates the `lastSyncTimestamp.txt` blob with the start time of the current run.

3.  **Discovery:** Users within the Microsoft 365 environment can now discover the indexed Purview glossary terms through Microsoft Search.

## Setup

1.  **Azure AD Application Registration:**
    *   Register a new application in Azure Active Directory.
    *   Create a Client Secret for the application.
    *   Note the Application (client) ID, Directory (tenant) ID, and the Client Secret value.
2.  **Permissions:**
    *   **Microsoft Purview:** Assign the `Data Curator` role to the registered application's service principal on your Purview account (or at a specific collection level if desired).
    *   **Microsoft Graph:** Grant the following Application permissions to the registered application in Azure AD:
        *   `ExternalConnection.ReadWrite.OwnedBy`
        *   `ExternalItem.ReadWrite.OwnedBy`
    *   Ensure admin consent is granted for the Graph permissions.
3.  **`PurviewSearchConnector` Console App:**
    *   Clone or download the `PurviewSearchConnector` project.
    *   Create a `config.ini` file in the project's output directory (e.g., `bin/Debug/netX.X`) with the following structure:
        ```ini
        [Azure]
        ClientID = YOUR_APP_CLIENT_ID
        ClientSecret = YOUR_APP_CLIENT_SECRET
        TenantID = YOUR_TENANT_ID

        [Purview]
        AccountName = YOUR_PURVIEW_ACCOUNT_NAME 
        ```
    *   **Important:** Create this file from the `PurviewSearchConnector/config.ini.template` and add your actual credentials. **Do not commit `config.ini` to source control** as it contains secrets.
    *   Build the project (`dotnet build`).
4.  **Create Graph Connection & Schema:**
    *   Run the `PurviewSearchConnector` executable (`dotnet run` or run the compiled `.exe`).
    *   Choose option `1` to create a new connection. Provide a unique ID (e.g., `purviewglossarysearch`), name, and description.
    *   **Important:** Note the `connectionId` you provided.
    *   Choose option `4` to register the schema for the newly created connection.
    *   Optionally, choose option `7` to push all current terms for an initial population.
5.  **Azure Function App Deployment:**
    *   Clone or download the `PurviewToGraphSyncFunction` project.
    *   Deploy the Function App to your Azure subscription (using Azure CLI, Portal, Visual Studio, VS Code, etc.). Choose a suitable hosting plan (e.g., Consumption).
6.  **Function App Configuration:**
    *   In the Azure portal, navigate to the deployed Function App.
    *   Go to `Configuration` -> `Application settings`.
    *   Add the following settings:
        *   `ClientID`: YOUR_APP_CLIENT_ID
        *   `ClientSecret`: YOUR_APP_CLIENT_SECRET
        *   `TenantID`: YOUR_TENANT_ID
        *   `PurviewEndpoint`: `https://YOUR_PURVIEW_ACCOUNT_NAME.purview.azure.com`
        *   `GraphConnectionId`: The `connectionId` noted in step 4.
        *   `AzureWebJobsStorage`: The connection string for an Azure Storage account (this is often created automatically with the Function App, or you can use an existing one). This storage account will hold the sync timestamp.
        *   `TimestampContainerName`: `purview-sync-timestamps` (or your preferred container name).
        *   `TimerSchedule`: The CRON expression for the sync schedule (e.g., `0 */30 * * * *` for every 30 minutes, `0 0 2 * * *` for 2 AM daily).
    *   **Important:** For local development, configure these settings in `PurviewToGraphSyncFunction/local.settings.json`. Create this file from `PurviewToGraphSyncFunction/local.settings.json.template`. **Do not commit `local.settings.json` to source control** if it contains secrets (like ClientSecret or AzureWebJobsStorage).
    *   Save the settings. The Function App may restart.

## Usage

*   **Automatic Synchronization:** Once set up, the `PurviewToGraphSyncFunction` runs automatically on the defined schedule, keeping the Graph `ExternalItems` updated with changes from the Purview glossary.
*   **Monitoring:** Monitor the execution of the Azure Function App through the Azure Portal (`Monitor` section of the Function App). Check for successful runs and investigate any logged errors.
*   **Manual Operations (via `PurviewSearchConnector`):**
    *   **Check Connection State:** Use option `9` to verify the Graph connection status.
    *   **Force Full Sync:** Use option `7` to re-index all terms if needed.
    *   **Delete Items:** Use option `8` to remove all indexed items from the Graph connection.
    *   **Delete Connection:** Use option `3` to delete the Graph connection entirely (requires re-running setup steps if you want to use the sync again).
*   **Search:** Search for Purview glossary terms via Microsoft Search in applications like SharePoint, Office.com, etc. The terms should appear as results linked back to their source (though the console app/function doesn't currently set a specific `Url` in the `ExternalItem` properties, Graph might generate a default one).

## Notes

*   **Security:** The `PurviewSearchConnector/config.ini` and `PurviewToGraphSyncFunction/local.settings.json` files contain sensitive credentials. Ensure they are excluded from source control (handled by the root `.gitignore` file). Use the provided `.template` files as a starting point.
*   **State File:** The `PurviewSearchConnector/lastSync.txt` file stores the timestamp of the last manual sync performed by the console app. It is generated at runtime and should generally not be committed to source control.
*   **Translations:** The `PurviewSearchConnector/PurviewHelper.cs` file contains an example of rudimentary translation functionality for the status of a term which can be expanded (look for `statusTranslations` ).
*   **Testing:** The file may contain testing mechanisms which have been commented out, feel free to reactivate them for verbose debugging.
*   **Dependencies:** Make sure to install the necessary packages from NuGet. Check MSLearn documentation for more details.
