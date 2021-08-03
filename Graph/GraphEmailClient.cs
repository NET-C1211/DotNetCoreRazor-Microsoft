
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
    public class GraphEmailClient
    {
        private readonly ILogger<GraphEmailClient> _logger;
        private readonly GraphServiceClient _graphServiceClient;

        public GraphEmailClient(ILogger<GraphEmailClient> logger, GraphServiceClient graphServiceClient)
        {
            _logger = logger;
            _graphServiceClient = graphServiceClient;
        }

        public async Task<IEnumerable<Message>> GetUserMessages()
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

        public async Task<(IEnumerable<Message> Messages, int Skip)> GetUserMessagesPage(int pageSize, int skip = 0)
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

            // var httpRequest = new HttpRequestMessage(HttpMethod.Get, "https://graph.microsoft.com/v1.0/me/messages?%24select=subject%2cbody%2cbodyPreview%2creceivedDateTime&%24orderby=receivedDateTime&%24top=5&%24skip=5");
            // var response = await _graphServiceClient.HttpProvider.SendAsync(httpRequest);

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