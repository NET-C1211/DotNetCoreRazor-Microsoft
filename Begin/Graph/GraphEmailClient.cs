
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
            string nextPageLink = null, int top = 10)
        {
            // Remove this code
            return await Task.FromResult<
                (IEnumerable<Message> Messages, string NextLink)>((Messages:null, NextLink:null));
        }

        private string GetNextLink(IUserMessagesCollectionPage pagedMessages) {
            if (pagedMessages.AdditionalData.TryGetValue("@odata.nextLink", out object value)) {
                return value.ToString();
            }
            return null;
        }

    }
}