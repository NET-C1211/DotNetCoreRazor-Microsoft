
using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System.Linq;
using System.Net;
using System.Net.Http;

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
                _logger.LogError($"Error calling Graph /me/messages: {ex.Message}");
                throw;
            }
        }

        public async Task<(IEnumerable<Message> Messages, string NextLink)> GetUserMessagesPage(
            string nextPageLink = null, int top = 5)
        {
            IUserMessagesCollectionPage pagedMessages;
            
            if (nextPageLink == null) 
            {
                // Get initial page of messages
                pagedMessages = await _graphServiceClient.Me.Messages
                        .Request()
                        .Select(msg => new
                        {
                            msg.Subject,
                            msg.BodyPreview,
                            msg.ReceivedDateTime
                        })
                        .Top(top)
                        .OrderBy("receivedDateTime desc")
                        .GetAsync();
            }
            else 
            {
                // Use nextLink value to get the page of messages
                UserMessagesCollectionRequest messagesCollectionRequest = 
                    new UserMessagesCollectionRequest(nextPageLink, _graphServiceClient, null);
                pagedMessages = await messagesCollectionRequest.GetAsync();
            }

            return (Messages: pagedMessages, NextLink: GetNextLink(pagedMessages));
        }

        private string GetNextLink(IUserMessagesCollectionPage pagedMessages) {
            if (pagedMessages.AdditionalData.TryGetValue("@odata.nextLink", out object value)) {
                return value.ToString();
            }
            return null;
        }

    }
}

// Couldn't get this working to make a "raw" API call due to token not being pass error
// _logger.LogInformation(nextPageLink);
// var httpRequest = new HttpRequestMessage(HttpMethod.Get, nextPageLink);
// var response = await _graphServiceClient.HttpProvider.SendAsync(httpRequest);
// Console.WriteLine(response.StatusCode);
// var stream = await response.Content.ReadAsStreamAsync();
// pagedMessages =  _graphServiceClient.HttpProvider.Serializer.DeserializeObject<IUserMessagesCollectionPage>(stream);
// nextLink = pagedMessages.AdditionalData["@odata.nextLink"].ToString();
// Console.WriteLine(pagedMessages.Count);

// var skipValue = pagedMessages
//     .NextPageRequest?
//     .QueryOptions?
//     .FirstOrDefault(
//         x => string.Equals("$skip", WebUtility.UrlDecode(x.Name), StringComparison.InvariantCultureIgnoreCase))?
//     .Value ?? "0";

// _logger.LogInformation($"skipValue: {skipValue}");

//return (Messages: pagedMessages, Skip: int.Parse(skipValue));