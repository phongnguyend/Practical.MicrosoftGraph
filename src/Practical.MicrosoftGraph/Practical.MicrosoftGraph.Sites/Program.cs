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

var fileLocation = UriPath.Combine(DateTime.Now.ToString("yyyy/MM/dd"), Guid.NewGuid().ToString());

var fileStream = new MemoryStream(Encoding.UTF8.GetBytes("Test"));

await sharePointOnlineClient.CreateAsync(fileLocation, fileStream);

var permissions = await sharePointOnlineClient.GetPermissionsAsync(fileLocation);

foreach (var permission in permissions)
{
    Console.WriteLine($"Permission: {permission.Id}, Roles: {string.Join(", ", permission.Roles)}, GrantedTo: {permission.GrantedTo?.User?.DisplayName}");
}