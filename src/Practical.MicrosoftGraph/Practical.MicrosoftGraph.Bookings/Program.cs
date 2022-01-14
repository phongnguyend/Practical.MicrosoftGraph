using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using System.Threading.Tasks;

namespace Practical.MicrosoftGraph.Bookings
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            var config = new ConfigurationBuilder()
            //.AddJsonFile("appsettings.json")
            .AddUserSecrets("473ed7c3-3710-46ab-a7f1-816a98fe18c6")
            .Build();

            // The client credentials flow requires that you request the
            // /.default scope, and preconfigure your permissions on the
            // app registration in Azure. An administrator must grant consent
            // to those permissions beforehand.
            var scopes = new[] { "Bookings.Read.All", "BookingsAppointment.ReadWrite.All", "Bookings.ReadWrite.All", "Bookings.Manage.All" };

            // Multi-tenant apps can use "common",
            // single-tenant apps must use the tenant ID from the Azure portal
            var tenantId = config["TenantId"];

            // Values from app registration
            var clientId = config["ClientId"];

            // using Azure.Identity;
            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            var userName = config["Bookings:UserName"];
            var password = config["Bookings:Password"];

            // https://docs.microsoft.com/dotnet/api/azure.identity.usernamepasswordcredential
            var userNamePasswordCredential = new UsernamePasswordCredential(
                userName, password, tenantId, clientId, options);

            var graphClient = new GraphServiceClient(userNamePasswordCredential, scopes);

            var bookingId = $"Test@{config["Domain"]}";
            var serviceId = "03d719b8-1dd6-437b-a10f-42c376046df6";

            var bookings = await graphClient.Solutions.BookingBusinesses[bookingId].Services[serviceId].Request().GetAsync();
            var staffs = await graphClient.Solutions.BookingBusinesses[bookingId].StaffMembers["14dda61e-513e-4c4c-a0ec-1837ec8f8987"].Request().GetAsync();

        }
    }
}
