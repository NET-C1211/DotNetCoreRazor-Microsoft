using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DotNetCoreRazor_MSGraph.Graph;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Web;

namespace DotNetCoreRazor_MSGraph.Pages
{
    [AuthorizeForScopes(ScopeKeySection = "DownstreamApi:Scopes")]
    public class EmailModel : PageModel
    {
        private readonly GraphUtils _graphUtils;
        
        [BindProperty(SupportsGet = true)]
        public int Skip { get; set; }
        public IList<Message> Messages  { get; private set; }

        public EmailModel(GraphUtils graphUtils)
        {
            _graphUtils = graphUtils;
        }

        public async Task OnGetAsync()
        {
            var messagesPagingData = await _graphUtils.GetUserMessagesPage(5, Skip); 
            Messages = messagesPagingData.Messages;
            Skip = messagesPagingData.Skip;
        }
    }
}
