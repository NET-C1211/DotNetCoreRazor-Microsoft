using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Web;

namespace DotNetCoreRazor_MSGraph.Pages
{
    [AuthorizeForScopes(Scopes = new[] { "user.read" })]
    [Authorize]
    public class IndexModel : PageModel
    {
        private readonly ILogger<IndexModel> _logger;
        private readonly GraphServiceClient _graphServiceClient;
        public string DisplayName { get; private set; } = "";
        public string Photo { get; private set; }

        public IndexModel(ILogger<IndexModel> logger, GraphServiceClient graphServiceClient)
        {
            _logger = logger;
            _graphServiceClient = graphServiceClient;
        }

        public async Task OnGetAsync()
        {
            User currentUser = null;

            try
            {
                currentUser = await _graphServiceClient.Me.Request().GetAsync();
                DisplayName = currentUser.DisplayName;
            }
            // Catch CAE exception from Graph SDK
            catch (ServiceException svcex) when (svcex.Message.Contains("Continuous access evaluation resulted in claims challenge"))
            {
                _logger.LogInformation("Error calling Graph /me");
            }

            try
            {
                // Get user photo
                using (var photoStream = await _graphServiceClient.Me.Photo.Content.Request().GetAsync())
                {
                    byte[] photo = ((MemoryStream)photoStream).ToArray();
                    Photo = Convert.ToBase64String(photo);
                }
            }
            catch (Exception pex)
            {
                Console.WriteLine($"{pex.Message}");
            }
        }
    }
}
