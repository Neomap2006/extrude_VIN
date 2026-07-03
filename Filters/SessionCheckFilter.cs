using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;

namespace VelastoProductionSystem.Filters
{
    public class SessionCheckFilter : ActionFilterAttribute
    {
        public override void OnActionExecuting(ActionExecutingContext context)
        {
            var controllerName = context.RouteData.Values["controller"]?.ToString();
            
            // Bypass session check for Account controller (Login/Logout/OfflineSession)
            // and for API routes and OfflineSync (public snapshot endpoint)
            var path = context.HttpContext.Request.Path.Value ?? "";
            if (controllerName == "Account" ||
                path.StartsWith("/api/", StringComparison.OrdinalIgnoreCase) ||
                path.StartsWith("/OfflineSync/", StringComparison.OrdinalIgnoreCase))
            {
                base.OnActionExecuting(context);
                return;
            }

            var userName = context.HttpContext.Session.GetString("UserName");
            if (string.IsNullOrEmpty(userName))
            {
                // Redirect to login if session is empty
                context.Result = new RedirectToActionResult("Login", "Account", null);
            }

            base.OnActionExecuting(context);
        }
    }
}
