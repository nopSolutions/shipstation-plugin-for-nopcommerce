using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Routing;
using Nop.Web.Framework.Mvc.Routing;

namespace Nop.Plugin.Shipping.ShipStation
{
    public partial class RouteProvider : IRouteProvider
    {
        public void RegisterRoutes(IRouteBuilder routeBuilder)
        {
            //Webhook
            routeBuilder.MapRoute("Plugin.Payments.ShipStation.WebhookHandler", "Plugins/ShipStation/Webhook",
                new { controller = "ShipStation", action = "Webhook" });
        }

        public int Priority => 0;
    }
}
