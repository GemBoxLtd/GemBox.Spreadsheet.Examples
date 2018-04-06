using Microsoft.AspNetCore.Http;
using System.Threading.Tasks;

namespace Spreadsheet.Samples.Core
{
    public class FrameAllowMiddleware
    {
        private readonly RequestDelegate next;

        public FrameAllowMiddleware(RequestDelegate next)
        {
            this.next = next;
        }

        public Task Invoke(HttpContext httpContext)
        {
            httpContext.Response.Headers["X-Frame-Options"] = "ALLOWALL";
            return next(httpContext);
        }
    }
}
