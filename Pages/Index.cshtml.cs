using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DotNetCoreRazor_MSGraph.Graph;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Web;

namespace DotNetCoreRazor_MSGraph.Pages
{
    [AuthorizeForScopes(ScopeKeySection = "DownstreamApi:Scopes")]
    public class IndexModel : PageModel
    {
        private readonly ILogger<IndexModel> _logger;
        private readonly GraphUtils _graphUtils;
        public string UserDisplayName { get; private set; } = "";
        public string UserPhoto { get; private set; }

        public IndexModel(GraphUtils graphUtils)
        {
            _graphUtils = graphUtils;
        }

        public async Task OnGetAsync()
        {
            var displayName = await _graphUtils.GetUserDisplayName(); 
            UserDisplayName = displayName.Split(' ')[0];
            UserPhoto = await _graphUtils.GetUserProfileImage();
        }
    }
}
