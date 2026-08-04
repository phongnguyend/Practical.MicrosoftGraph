// Load configuration from user secrets
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Practical.MicrosoftGraph.Sites;
using System;
using System.IO;
using System.Text;
using UriHelper;

SharePointOnlineClientOptions options = new SharePointOnlineClientOptions();
IMemoryCache memoryCache;

var configuration = new ConfigurationBuilder()
    .AddUserSecrets<Program>()
    .Build();

configuration.GetSection("SharePointOnline").Bind(options);

// Create IMemoryCache instance using DI container
var services = new ServiceCollection();
services.AddMemoryCache();
var provider = services.BuildServiceProvider();
memoryCache = provider.GetService<IMemoryCache>();

var sharePointOnlineClient = new SharePointOnlineClient(options, memoryCache);

var subscriptions = await sharePointOnlineClient.GetWebhookSubscriptionsAsync();

foreach (var subscription in subscriptions)
{
    Console.WriteLine($"Subscription: {subscription.Id}, ExpirationDateTime: {subscription.ExpirationDateTime}, ClientState: {subscription.ClientState}");

    //await sharePointOnlineClient.DeleteWebhookSubscriptionAsync(subscription.Id);
}

//var createdSubscription = await sharePointOnlineClient.CreateWebhookSubscriptionAsync($"https://ddddotnet-webhook-server.azurewebsites.net/tenants/{options.TenantId}/topics/sharepoint", DateTimeOffset.Now.AddMinutes(43200), "test client state");

//var fileLocation = UriPath.Combine(DateTime.Now.ToString("yyyy/MM/dd"), Guid.NewGuid().ToString());

//var fileStream = new MemoryStream(Encoding.UTF8.GetBytes("Test"));

//await sharePointOnlineClient.CreateAsync(fileLocation, fileStream);

//var permissions = await sharePointOnlineClient.GetPermissionsAsync(fileLocation);

//foreach (var permission in permissions)
//{
//    Console.WriteLine($"Permission: {permission.Id}, Roles: {string.Join(", ", permission.Roles)}, GrantedTo: {permission.GrantedTo?.User?.DisplayName}");
//}

// Keep calling delta with the returned link until it stops returning new items,
// i.e. we've caught up to the current state (Microsoft Graph will only send further
// webhook notifications once we've consumed all pending changes this way).
string deltaLink = null;

while (true)
{
    var (changedItems, newDeltaLink) = await sharePointOnlineClient.GetDeltaAsync(deltaLink);

    foreach (var change in changedItems)
    {
        Console.WriteLine($"{change.EventType} {change.ItemType}: {change.Item.Name} ({change.Item.Id}), ETag: {change.Item.ETag}, SharedChanged: {change.SharedChanged}");
    }

    if (changedItems.Count == 0)
    {
        // Caught up - no more pending changes. Persist newDeltaLink so the next
        // webhook notification can resume from here.
        deltaLink = newDeltaLink;
        break;
    }

    deltaLink = newDeltaLink;
}

Console.WriteLine($"DeltaLink: {deltaLink}");