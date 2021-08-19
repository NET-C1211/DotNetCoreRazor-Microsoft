
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

        public async Task<(IEnumerable<Message> Messages, string NextLink)> GetUserMessagesPage(string nextPageLink = null)
        {
            int top = 5;
            IUserMessagesCollectionPage pagedMessages;
            string nextLink;
            
            if (nextPageLink == null) 
            {
                // Get initial page of messages
                pagedMessages = await _graphServiceClient.Me.Messages
                        .Request()
                        .Select(msg => new
                        {
                            msg.Subject,
                            msg.Body,
                            msg.BodyPreview,
                            msg.ReceivedDateTime
                        })
                        .Top(top)
                        .OrderBy("receivedDateTime")
                        .GetAsync();
                nextLink = pagedMessages.AdditionalData["@odata.nextLink"].ToString();
                _logger.LogInformation(nextLink); 
            }
            else 
            {
                // Use nextLink value to get the page of messages
                _logger.LogInformation(nextPageLink);
                var httpRequest = new HttpRequestMessage(HttpMethod.Get, nextPageLink);
                var response = await _graphServiceClient.HttpProvider.SendAsync(httpRequest);
                Console.WriteLine(response.StatusCode);
                var stream = await response.Content.ReadAsStreamAsync();
                pagedMessages =  _graphServiceClient.HttpProvider.Serializer.DeserializeObject<IUserMessagesCollectionPage>(stream);
                nextLink = pagedMessages.AdditionalData["@odata.nextLink"].ToString();
                Console.WriteLine(pagedMessages.Count);
            }

            return (Messages: pagedMessages, NextLink: nextLink);


            // var skipValue = pagedMessages
            //     .NextPageRequest?
            //     .QueryOptions?
            //     .FirstOrDefault(
            //         x => string.Equals("$skip", WebUtility.UrlDecode(x.Name), StringComparison.InvariantCultureIgnoreCase))?
            //     .Value ?? "0";

            // _logger.LogInformation($"skipValue: {skipValue}");

            //return (Messages: pagedMessages, Skip: int.Parse(skipValue));
        }

    }
}