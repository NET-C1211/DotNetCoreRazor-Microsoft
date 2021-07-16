
using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System.Linq;

namespace DotNetCoreRazor_MSGraph.Graph
{
    public class GraphFilesClient
    {
        private readonly ILogger<GraphFilesClient> _logger;
        private readonly GraphServiceClient _graphServiceClient;

        public GraphFilesClient(ILogger<GraphFilesClient> logger, GraphServiceClient graphServiceClient)
        {
            _logger = logger;
            _graphServiceClient = graphServiceClient;
        }

        public async Task<IDriveItemChildrenCollectionPage> GetFiles()
        {
            try 
            {
                return await _graphServiceClient.Me.Drive.Root.Children
                            .Request()
                            .Select(file => new {
                                file.Id,
                                file.Name,
                                file.Folder,
                                file.Package
                            })
                            .GetAsync();
            }
            catch (Exception ex)
            {
                _logger.LogInformation($"Error calling Graph /me/drive/root/children: {ex.Message}");
                throw;
            }
        }

        public async Task<Stream> DownloadFile(string fileId) {
            try 
            {
                return await _graphServiceClient
                            .Me.Drive.Items[fileId].Content
                            .Request()
                            .GetAsync();
            }
            catch (Exception ex)
            {
                _logger.LogInformation($"Error downloading file: {ex.Message}");
                throw;
            }
        }

        public async Task<DriveItem> UploadFile(Stream fileStream) {
            try 
            {
                return await _graphServiceClient
                                .Users["upn or userID"]
                                .Drive.Items["{item-id}"].Content
                                .Request()
                                .PutAsync<DriveItem>(fileStream);
            }
            catch (Exception ex)
            {
                _logger.LogInformation($"Error uploading file: {ex.Message}");
                throw;
            }
        }

    }
}