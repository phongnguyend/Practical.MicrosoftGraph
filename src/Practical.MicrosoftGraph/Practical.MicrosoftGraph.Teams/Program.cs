using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Practical.MicrosoftGraph.Teams;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

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
    Console.WriteLine($"- {team.DisplayName} (ID: {team.Id})");
}

// Create a new team
Console.WriteLine("\nCreating a new team...");
var newTeam = await teamsManager.CreateTeamAsync("My New Team", "Team created from C#");
if (newTeam != null)
{
    Console.WriteLine($"Team created: {newTeam.DisplayName} (ID: {newTeam.Id})");
    var teamId = newTeam.Id;

    // Wait a moment for the team to be fully provisioned
    await Task.Delay(2000);

    // Team Owners and Members Management
    Console.WriteLine("\n=== Team Owners and Members Management ===\n");

    // List team owners
    Console.WriteLine("Listing team owners...");
    var owners = await teamsManager.ListTeamOwnersAsync(teamId);
    Console.WriteLine($"Found {owners.Count} owner(s)");
    foreach (var owner in owners)
    {
        Console.WriteLine($"- {owner.Id}");
    }

    // List team members
    Console.WriteLine("\nListing team members...");
    var members = await teamsManager.ListTeamMembersAsync(teamId);
    Console.WriteLine($"Found {members.Count} member(s)");
    foreach (var member in members)
    {
        Console.WriteLine($"- {member.DisplayName} (ID: {member.Id})");
    }

    // List team members with roles (owners and members)
    Console.WriteLine("\nListing team members with roles...");
    var membersWithRoles = await teamsManager.ListTeamMembersWithRolesAsync(teamId);
    Console.WriteLine($"Found {membersWithRoles.Count} member(s) total:");
    var teamOwners = membersWithRoles.Where(m => m.Roles.Contains("owner")).ToList();
    var regularMembers = membersWithRoles.Where(m => !m.Roles.Contains("owner")).ToList();

    Console.WriteLine($"\n  Owners ({teamOwners.Count}):");
    foreach (var owner in teamOwners)
    {
        Console.WriteLine($"  - {owner.DisplayName} (ID: {owner.Id}) - Roles: {string.Join(", ", owner.Roles)}");
    }

    Console.WriteLine($"\n  Regular Members ({regularMembers.Count}):");
    foreach (var member in regularMembers)
    {
        Console.WriteLine($"  - {member.DisplayName} (ID: {member.Id}) - Roles: {string.Join(", ", member.Roles)}");
    }

    // Add a team owner (use a real user ID)
    var ownerUserId = "user-id-here"; // Replace with actual user ID
    Console.WriteLine($"\nAdding team owner...");
    var addOwnerSuccess = await teamsManager.AddTeamOwnerAsync(teamId, ownerUserId);
    Console.WriteLine(addOwnerSuccess ? "Owner added successfully" : "Failed to add owner");

    // Add a team member (use a real user ID)
    var memberUserId = "user-id-here"; // Replace with actual user ID
    Console.WriteLine($"\nAdding team member...");
    var addMemberSuccess = await teamsManager.AddTeamMemberAsync(teamId, memberUserId, "member");
    Console.WriteLine(addMemberSuccess ? "Member added successfully" : "Failed to add member");

    // Update team member roles
    if (members.Count > 0)
    {
        var firstMemberId = members.First().Id;
        Console.WriteLine($"\nUpdating member roles...");
        var updateRolesSuccess = await teamsManager.UpdateTeamMemberAsync(teamId, firstMemberId, new List<string> { "owner" });
        Console.WriteLine(updateRolesSuccess ? "Member roles updated successfully" : "Failed to update member roles");

        // Remove team member
        Console.WriteLine($"\nRemoving team member...");
        var removeMemberSuccess = await teamsManager.RemoveTeamMemberAsync(teamId, firstMemberId);
        Console.WriteLine(removeMemberSuccess ? "Member removed successfully" : "Failed to remove member");
    }

    // List channels in the new team
    Console.WriteLine("\n=== Channels Management ===\n");
    Console.WriteLine("Listing channels in the team...");
    var channels = await teamsManager.ListChannelsAsync(teamId);
    Console.WriteLine($"Found {channels.Count} channel(s)");
    foreach (var channel in channels)
    {
        Console.WriteLine($"- {channel.DisplayName} (ID: {channel.Id})");
    }

    // Create a new channel
    Console.WriteLine("\nCreating a new channel...");
    var newChannel = await teamsManager.CreateChannelAsync(teamId, "My New Channel", "Channel created from C#");
    if (newChannel != null)
    {
        Console.WriteLine($"Channel created: {newChannel.DisplayName} (ID: {newChannel.Id})");
        var channelId = newChannel.Id;

        // Get specific channel
        Console.WriteLine("\nGetting channel details...");
        var channelDetails = await teamsManager.GetChannelAsync(teamId, channelId);
        if (channelDetails != null)
        {
            Console.WriteLine($"Channel: {channelDetails.DisplayName}");
            Console.WriteLine($"Description: {channelDetails.Description}");
        }

        // Update channel
        Console.WriteLine("\nUpdating channel...");
        var updateSuccess = await teamsManager.UpdateChannelAsync(teamId, channelId, "Updated Channel Name", "Updated description");
        Console.WriteLine(updateSuccess ? "Channel updated successfully" : "Failed to update channel");

        // Channel Files Management
        Console.WriteLine("\n=== Channel Files Management ===\n");
        Console.WriteLine("Listing channel files...");
        var files = await teamsManager.ListChannelFilesAsync(teamId, channelId);
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
            var newFile = await teamsManager.CreateChannelFileAsync(teamId, channelId, "test-file.txt", stream);
            if (newFile != null)
            {
                Console.WriteLine($"File created: {newFile.Name} (ID: {newFile.Id})");
                var fileId = newFile.Id;

                // Get specific file
                Console.WriteLine("\nGetting file details...");
                var fileDetails = await teamsManager.GetChannelFileAsync(teamId, channelId, fileId);
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
                    var updateFileSuccess = await teamsManager.UpdateChannelFileAsync(teamId, channelId, fileId, updateStream);
                    Console.WriteLine(updateFileSuccess ? "File updated successfully" : "Failed to update file");
                }

                // Delete file
                Console.WriteLine("\nDeleting file...");
                var deleteFileSuccess = await teamsManager.DeleteChannelFileAsync(teamId, channelId, fileId);
                Console.WriteLine(deleteFileSuccess ? "File deleted successfully" : "Failed to delete file");
            }
        }

        // Delete channel
        Console.WriteLine("\nDeleting channel...");
        var deleteSuccess = await teamsManager.DeleteChannelAsync(teamId, channelId);
        Console.WriteLine(deleteSuccess ? "Channel deleted successfully" : "Failed to delete channel");
    }

    // Update team
    Console.WriteLine("\nUpdating team...");
    var updateTeamSuccess = await teamsManager.UpdateTeamAsync(teamId, "Updated Team Name", "Updated team description");
    Console.WriteLine(updateTeamSuccess ? "Team updated successfully" : "Failed to update team");

    // Delete team
    Console.WriteLine("\nDeleting team...");
    var deleteTeamSuccess = await teamsManager.DeleteTeamAsync(teamId);
    Console.WriteLine(deleteTeamSuccess ? "Team deleted successfully" : "Failed to delete team");
}

Console.WriteLine("\n=== Operations Complete ===");