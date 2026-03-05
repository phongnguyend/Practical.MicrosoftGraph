using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Practical.MicrosoftGraph.Teams;
using System;
using System.IO;
using System.Linq;

var config = new ConfigurationBuilder()
    .AddUserSecrets("473ed7c3-3710-46ab-a7f1-816a98fe18c6")
    .Build();

var scopes = new[] { "https://graph.microsoft.com/.default" };
var tenantId = config["TenantId"];
var clientId = config["ClientId"];
var clientSecret = config["ClientSecret"];

var options = new TokenCredentialOptions
{
    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
};

var clientSecretCredential = new ClientSecretCredential(tenantId, clientId, clientSecret, options);

var graphClient = new GraphServiceClient(clientSecretCredential, scopes);
var teamsManager = new TeamsManager(graphClient);

// Example usage
Console.WriteLine("=== Teams Management ===\n");

// List teams
Console.WriteLine("Listing all teams...");

var teams = await teamsManager.ListTeamsAsync();

Console.WriteLine($"Found {teams.Count} team(s)");

foreach (var team in teams)
{
    if (team.Id == "33c2816e-ca12-4a96-a978-fa7e3086db29")
    {
        continue;
    }

    Console.WriteLine($"- {team.DisplayName} (ID: {team.Id})");

    // List team members
    Console.WriteLine("\nListing team members...");
    var members = await teamsManager.ListTeamMembersAsync(team.Id);
    Console.WriteLine($"Found {members.Count} member(s)");
    foreach (var member in members)
    {
        Console.WriteLine($"- {member.DisplayName} (ID: {member.Id})");
    }
}


var myTeam = teams.FirstOrDefault(x => x.DisplayName.Contains("My New Team"));

if (myTeam == null)
{
    // Create a new team
    Console.WriteLine("\nCreating a new team...");

    var user = await graphClient.Users["phongnguyend@phungnguyenminh.onmicrosoft.com"].GetAsync();

    var newTeam = await teamsManager.CreateTeamAsync("My New Team", "Team created from C#", user.Id);

    return;
}


// List channels in the new team
Console.WriteLine("\n=== Channels Management ===\n");
Console.WriteLine("Listing channels in the team...");
var channels = await teamsManager.ListChannelsAsync(myTeam.Id);
Console.WriteLine($"Found {channels.Count} channel(s)");
foreach (var channel in channels)
{
    Console.WriteLine($"- {channel.DisplayName} (ID: {channel.Id})");
}

// Create a new channel
Console.WriteLine("\nCreating a new channel...");

var newChannel = channels.FirstOrDefault(x => x.DisplayName.Contains("My New Channel")) ?? await teamsManager.CreateChannelAsync(myTeam.Id, "My New Channel", "Channel created from C#");

if (newChannel != null)
{
    Console.WriteLine($"Channel created: {newChannel.DisplayName} (ID: {newChannel.Id})");
    var channelId = newChannel.Id;

    // Get specific channel
    Console.WriteLine("\nGetting channel details...");
    var channelDetails = await teamsManager.GetChannelAsync(myTeam.Id, channelId);
    if (channelDetails != null)
    {
        Console.WriteLine($"Channel: {channelDetails.DisplayName}");
        Console.WriteLine($"Description: {channelDetails.Description}");
    }

    // Update channel
    Console.WriteLine("\nUpdating channel...");
    await teamsManager.UpdateChannelAsync(myTeam.Id, channelId, "My New Channel Updated at " + DateTime.UtcNow.Ticks, "Updated description");
    Console.WriteLine("Channel updated successfully");

    // Channel Files Management
    Console.WriteLine("\n=== Channel Files Management ===\n");
    Console.WriteLine("Listing channel files...");
    var files = await teamsManager.ListChannelFilesAsync(myTeam.Id, channelId);
    Console.WriteLine($"Found {files.Count} file(s)");
    foreach (var file in files)
    {
        Console.WriteLine($"- {file.Name} (ID: {file.Id})");
    }

    // Create a file in the channel
    Console.WriteLine("\nCreating a file in the channel...");
    var fileContent = "Hello from Teams Channel!";
    using (var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(fileContent)))
    {
        var newFile = await teamsManager.CreateChannelFileAsync(myTeam.Id, channelId, "test-file.txt", stream);
        if (newFile != null)
        {
            Console.WriteLine($"File created: {newFile.Name} (ID: {newFile.Id})");
            var fileId = newFile.Id;

            // Get specific file
            Console.WriteLine("\nGetting file details...");
            var fileDetails = await teamsManager.GetChannelFileAsync(myTeam.Id, channelId, newFile.Name);
            if (fileDetails != null)
            {
                Console.WriteLine($"File: {fileDetails.Name}");
                Console.WriteLine($"Size: {fileDetails.Size} bytes");
            }

            // Update file
            Console.WriteLine("\nUpdating file...");
            var updatedFileContent = "Updated content from Teams Channel!";
            using (var updateStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(updatedFileContent)))
            {
                await teamsManager.UpdateChannelFileAsync(myTeam.Id, channelId, fileId, updateStream);
                Console.WriteLine("File updated successfully");
            }

            // Delete file
            Console.WriteLine("\nDeleting file...");
            await teamsManager.DeleteChannelFileAsync(myTeam.Id, channelId, fileId);
            Console.WriteLine("File deleted successfully");
        }
    }

    // Delete channel
    Console.WriteLine("\nDeleting channel...");
    await teamsManager.DeleteChannelAsync(myTeam.Id, channelId);
    Console.WriteLine("Channel deleted successfully");
}

// Update team
Console.WriteLine("\nUpdating team...");
await teamsManager.UpdateTeamAsync(myTeam.Id, "My New Team Updated at " + DateTime.UtcNow.Ticks, "Updated team description");
Console.WriteLine("Team updated successfully");

// Delete team
Console.WriteLine("\nDeleting team...");
await teamsManager.DeleteTeamAsync(myTeam.Id);
Console.WriteLine("Team deleted successfully");

Console.WriteLine("\n=== Operations Complete ===");