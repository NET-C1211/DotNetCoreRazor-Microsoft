
using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System.Linq;
using System.Net;

namespace DotNetCoreRazor_MSGraph.Graph
{
    public class GraphUtils
    {
        private readonly ILogger<GraphUtils> _logger;
        private readonly GraphServiceClient _graphServiceClient;

        public GraphUtils(ILogger<GraphUtils> logger, GraphServiceClient graphServiceClient)
        {
            _logger = logger;
            _graphServiceClient = graphServiceClient;
        }
        public async Task<string> GetUserDisplayName()
        {
            User currentUser = null;

            try
            {
                currentUser = await _graphServiceClient.Me.Request().GetAsync();
                return currentUser.DisplayName;
            }
            // Catch CAE exception from Graph SDK
            catch (ServiceException ex) when (ex.Message.Contains("Continuous access evaluation resulted in claims challenge"))
            {
                _logger.LogInformation($"/me Continuous access evaluation resulted in claims challenge: {ex.Message}");
                throw;
            }
        }
        public async Task<string> GetUserProfileImage()
        {
            try
            {
                // Get user photo
                using (var photoStream = await _graphServiceClient.Me.Photo.Content.Request().GetAsync())
                {
                    byte[] photo = ((MemoryStream)photoStream).ToArray();
                    return Convert.ToBase64String(photo);
                }
            }
            catch (Exception ex)
            {
                _logger.LogInformation($"Error calling Graph /me/photo: {ex.Message}");
                throw;
            }
        }

        public async Task<IList<Message>> GetUserMessages()
        {
            try
            {
                var emails = await _graphServiceClient.Me.Messages
                            .Request()
                            .Select(msg => new
                            {
                                msg.Subject,
                                msg.Body,
                                msg.BodyPreview,
                                msg.ReceivedDateTime
                            })
                            .OrderBy("receivedDateTime")
                            .Top(10)
                            .GetAsync();

                return emails.CurrentPage;
            }
            catch (Exception ex)
            {
                _logger.LogInformation($"Error calling Graph /me/messages: {ex.Message}");
                throw;
            }
        }

        public async Task<(IList<Message> Messages, int Skip)> GetUserMessagesPage(int pageSize, int skip = 0)
        {
            var pagedMessages = await _graphServiceClient.Me.Messages
                    .Request()
                    .Select(msg => new
                    {
                        msg.Subject,
                        msg.Body,
                        msg.BodyPreview,
                        msg.ReceivedDateTime
                    })
                    .Top(pageSize)
                    .Skip(skip)
                    .OrderBy("receivedDateTime")
                    .GetAsync();

            var skipValue = pagedMessages
                .NextPageRequest?
                .QueryOptions?
                .FirstOrDefault(
                    x => string.Equals("$skip", WebUtility.UrlDecode(x.Name), StringComparison.InvariantCultureIgnoreCase))?
                .Value ?? "0";

            _logger.LogInformation($"skipValue: {skipValue}");

            return (Messages: pagedMessages, Skip: int.Parse(skipValue));
        }

    }
}