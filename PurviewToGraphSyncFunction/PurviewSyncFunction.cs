using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Graph.Models.ExternalConnectors;
using Azure;
using Azure.Analytics.Purview.DataMap;

namespace PurviewToGraphSyncFunction
{
    public class PurviewSyncFunction
    {
        private readonly ILogger _logger;
        private const string TimestampBlobName = "lastSyncTimestamp.txt";

        public PurviewSyncFunction(ILoggerFactory loggerFactory)
        {
            _logger = loggerFactory.CreateLogger<PurviewSyncFunction>();
        }

        [Function("PurviewSyncFunction")]
        public async Task Run([TimerTrigger("%TimerSchedule%")] TimerInfo myTimer)
        {
            var functionStartTime = DateTimeOffset.UtcNow;
            _logger.LogInformation($"Purview-Graph Sync Function triggered at: {functionStartTime:o}");

            try
            {
                _logger.LogInformation("Loading settings from environment variables...");
                var settings = LoadSettingsFromEnvironment();
                _logger.LogInformation("Settings loaded successfully.");
                _logger.LogInformation("Target Graph Connection ID: {GraphConnectionId}", settings.GraphConnectionId);
                _logger.LogInformation("Purview Endpoint: {PurviewEndpoint}", settings.PurviewEndpoint);
                _logger.LogInformation("Timestamp Container Name: {TimestampContainerName}", settings.TimestampContainerName);

                GraphHelper.InitializeGraph(settings, _logger);
                PurviewHelper.InitializePurview(settings, _logger);

                _logger.LogInformation("Connecting to Azure Blob Storage for timestamp management...");
                BlobServiceClient blobServiceClient = new BlobServiceClient(settings.AzureWebJobsStorage);
                BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient(settings.TimestampContainerName);
                await containerClient.CreateIfNotExistsAsync();
                BlobClient timestampBlobClient = containerClient.GetBlobClient(TimestampBlobName);

                DateTimeOffset? lastSyncTime = await ReadLastSyncTimeAsync(timestampBlobClient);
                if (lastSyncTime.HasValue)
                {
                    _logger.LogInformation("Last successful sync timestamp read from blob: {LastSyncTime}", lastSyncTime.Value.ToString("o"));
                }
                else
                {
                    _logger.LogInformation("No previous sync timestamp found (or failed to read). Performing a full sync.");
                }

                _logger.LogInformation("Fetching glossary terms from Purview...");
                List<(AtlasGlossaryTerm Term, string GlossaryName)> termsToSync = 
                    await PurviewHelper.GetGlossaryTermsDataAsync(targetGlossaryId: null, lastSyncTime: lastSyncTime);

                if (termsToSync == null || termsToSync.Count == 0)
                {
                    _logger.LogInformation("No new or modified terms found in Purview since the last sync time.");
                    await UpdateLastSyncTimeAsync(timestampBlobClient, functionStartTime);
                    _logger.LogInformation("Purview-Graph Sync Function finished successfully at: {EndTime}", DateTimeOffset.UtcNow.ToString("o"));
                    return;
                }
                
                _logger.LogInformation("Found {TermCount} terms to sync.", termsToSync.Count);

                _logger.LogInformation("Mapping Purview terms to Graph ExternalItems...");
                List<ExternalItem> externalItems = new List<ExternalItem>();
                foreach (var (term, glossaryName) in termsToSync)
                {
                    try
                    {
                        ExternalItem? item = PurviewHelper.MapPurviewTermToExternalItem(term, glossaryName, settings.TenantID!, _logger);
                        if (item != null)
                        {
                            externalItems.Add(item);
                        }
                    }
                    catch (Exception ex)
                    {
                         _logger.LogError(ex, "Error mapping term {TermGuid} ('{TermName}') from glossary {GlossaryName}. Skipping this term.", term.Guid, term.Name, glossaryName);
                    }
                }
                _logger.LogInformation("Successfully mapped {ItemCount} terms to ExternalItems.", externalItems.Count);

                if (externalItems.Count == 0)
                {
                    _logger.LogWarning("Although terms were found, none could be successfully mapped to ExternalItems. Check mapping errors above.");
                    await UpdateLastSyncTimeAsync(timestampBlobClient, functionStartTime);
                    _logger.LogInformation("Purview-Graph Sync Function finished with mapping issues at: {EndTime}", DateTimeOffset.UtcNow.ToString("o"));
                    return;
                }

                _logger.LogInformation("Pushing {ItemCount} items to Graph Connection ID {ConnectionId}...", externalItems.Count, settings.GraphConnectionId);
                int successCount = await GraphHelper.AddOrUpdateItemsAsync(settings.GraphConnectionId!, externalItems);
                 _logger.LogInformation("Graph push complete. Successfully added/updated {SuccessCount} / {TotalCount} items.", successCount, externalItems.Count);

                if (successCount > 0 || termsToSync.Count == 0)
                {
                    await UpdateLastSyncTimeAsync(timestampBlobClient, functionStartTime);
                }
                else
                {
                     _logger.LogError("Failed to push any items to Microsoft Graph. Timestamp will not be updated to retry these items on the next run.");
                }
            }
            catch (ArgumentNullException argEx)
            {
                 _logger.LogError(argEx, "Configuration Error: A required setting is missing. Parameter: {ParamName}", argEx.ParamName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "An unhandled error occurred during the sync process: {ErrorMessage}", ex.Message);
            }
            finally
            {
                 _logger.LogInformation("Purview-Graph Sync Function finished execution at: {EndTime}", DateTimeOffset.UtcNow.ToString("o"));
                 if (myTimer.ScheduleStatus is not null)
                 {
                     _logger.LogInformation($"Next timer schedule at: {myTimer.ScheduleStatus.Next:o}");
                 }
            }
        }

        private Settings LoadSettingsFromEnvironment()
        {
            var settings = new Settings
            {
                ClientID = Environment.GetEnvironmentVariable("ClientID", EnvironmentVariableTarget.Process),
                ClientSecret = Environment.GetEnvironmentVariable("ClientSecret", EnvironmentVariableTarget.Process),
                TenantID = Environment.GetEnvironmentVariable("TenantID", EnvironmentVariableTarget.Process),
                PurviewEndpoint = Environment.GetEnvironmentVariable("PurviewEndpoint", EnvironmentVariableTarget.Process),
                GraphConnectionId = Environment.GetEnvironmentVariable("GraphConnectionId", EnvironmentVariableTarget.Process),
                AzureWebJobsStorage = Environment.GetEnvironmentVariable("AzureWebJobsStorage", EnvironmentVariableTarget.Process),
                TimestampContainerName = Environment.GetEnvironmentVariable("TimestampContainerName", EnvironmentVariableTarget.Process) ?? "purview-sync-timestamps"
            };

            if (string.IsNullOrEmpty(settings.ClientID)) throw new ArgumentNullException(nameof(settings.ClientID), "ClientID setting is missing.");
            if (string.IsNullOrEmpty(settings.ClientSecret)) throw new ArgumentNullException(nameof(settings.ClientSecret), "ClientSecret setting is missing.");
            if (string.IsNullOrEmpty(settings.TenantID)) throw new ArgumentNullException(nameof(settings.TenantID), "TenantID setting is missing.");
            if (string.IsNullOrEmpty(settings.PurviewEndpoint)) throw new ArgumentNullException(nameof(settings.PurviewEndpoint), "PurviewEndpoint setting is missing.");
            if (string.IsNullOrEmpty(settings.GraphConnectionId)) throw new ArgumentNullException(nameof(settings.GraphConnectionId), "GraphConnectionId setting is missing.");
            if (string.IsNullOrEmpty(settings.AzureWebJobsStorage)) throw new ArgumentNullException(nameof(settings.AzureWebJobsStorage), "AzureWebJobsStorage setting is missing.");

            return settings;
        }

        private async Task<DateTimeOffset?> ReadLastSyncTimeAsync(BlobClient blobClient)
        {
            try
            {
                if (await blobClient.ExistsAsync())
                {
                    _logger.LogInformation("Reading last sync timestamp from blob: {BlobName}", blobClient.Name);
                    Response<BlobDownloadResult> downloadResult = await blobClient.DownloadContentAsync();
                    string timestampString = downloadResult.Value.Content.ToString();
                    
                    if (DateTimeOffset.TryParseExact(timestampString, "o", CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTimeOffset lastSync))
                    {
                        return lastSync;
                    }
                    else
                    {
                         _logger.LogWarning("Could not parse timestamp string '{TimestampString}' from blob {BlobName}. Performing full sync.", timestampString, blobClient.Name);
                        return null;
                    }
                }
                else
                {
                    _logger.LogInformation("Timestamp blob {BlobName} does not exist. Assuming first run.", blobClient.Name);
                    return null;
                }
            }
            catch (RequestFailedException rfEx) when (rfEx.Status == 404)
            {
                 _logger.LogInformation("Timestamp blob {BlobName} not found (404). Assuming first run.", blobClient.Name);
                 return null;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error reading last sync time from blob {BlobName}. Performing full sync as a precaution.", blobClient.Name);
                return null;
            }
        }

        private async Task UpdateLastSyncTimeAsync(BlobClient blobClient, DateTimeOffset syncTime)
        {
            try
            {
                string timestampString = syncTime.ToString("o", CultureInfo.InvariantCulture);
                 _logger.LogInformation("Updating last sync timestamp in blob {BlobName} to: {SyncTime}", blobClient.Name, timestampString);
                using var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(timestampString));
                await blobClient.UploadAsync(stream, overwrite: true);
                _logger.LogInformation("Timestamp blob updated successfully.");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to update timestamp blob {BlobName}. Last successful sync time may be outdated for the next run.", blobClient.Name);
            }
        }
    }
}
