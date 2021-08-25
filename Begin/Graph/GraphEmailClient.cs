
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
        private readonly ILogger<GraphEmailClient> _logger = null;
        private readonly GraphServiceClient _graphServiceClient = null;

        public GraphEmailClient()
        {

        }

        public async Task<IEnumerable<Message>> GetUserMessages()
        {
            // Remove this code
            return await Task.FromResult<IEnumerable<Message>>(null);
        }

        public async Task<(IEnumerable<Message> Messages, string NextLink)> GetUserMessagesPage(
            string nextPageLink = null)
        {
            // Remove this code
            return await Task.FromResult<
                (IEnumerable<Message> Messages, string NextLink)>((Messages:null, NextLink:null));

            int top = 5;
            IUserMessagesCollectionPage pagedMessages;
            
            try 
            {
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
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error calling Graph /me/messages to page messages: {ex.Message}");
                throw;
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