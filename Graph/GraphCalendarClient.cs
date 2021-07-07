
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
    public class GraphCalendarClient
    {
        private readonly ILogger<GraphCalendarClient> _logger;
        private readonly GraphServiceClient _graphServiceClient;

        public GraphCalendarClient(ILogger<GraphCalendarClient> logger, GraphServiceClient graphServiceClient)
        {
            _logger = logger;
            _graphServiceClient = graphServiceClient;
        }

        public async Task<Calendar> GetUserCalendar()
        {
            try
            {
                var calendar = await _graphServiceClient.Me
                            .Calendar
                            .Request()
                            .GetAsync();

                return calendar;
            }
            catch (Exception ex)
            {
                _logger.LogInformation($"Error calling Graph /me/calendar: {ex.Message}");
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