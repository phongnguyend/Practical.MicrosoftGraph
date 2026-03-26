using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Practical.MicrosoftGraph.Teams;

public class TeamsManager
{
    private readonly GraphServiceClient _graphClient;

    public TeamsManager(GraphServiceClient graphClient)
    {
        _graphClient = graphClient;
    }

    public async Task<List<Team>> ListTeamsAsync()
    {
        var teams = await _graphClient.Teams.GetAsync();
        return teams?.Value?.ToList() ?? new List<Team>();
    }

    public async Task<Team?> GetTeamAsync(string teamId)
    {
        var team = await _graphClient.Teams[teamId].GetAsync();
        return team;
    }

    public async Task<Team> GetTeamByNameAsync(string teamName)
    {
        var teams = await _graphClient.Teams.GetAsync(config => config.QueryParameters.Filter = $"displayName eq '{teamName}'");

        return teams?.Value?.SingleOrDefault();
    }

    public async Task<List<Team>> GetTeamsByNameAsync(string teamName)
    {
        var teams = await _graphClient.Teams.GetAsync(config => config.QueryParameters.Filter = $"displayName eq '{teamName}'");

        return teams?.Value?.ToList() ?? new List<Team>();
    }

    public async Task<Team?> CreateTeamAsync(string displayName, string description, string ownerUserId)
    {
        var team = new Team
        {
            DisplayName = displayName,
            Description = description,
            Visibility = TeamVisibilityType.Private,
            AdditionalData = new Dictionary<string, object>
            {
                {
                    "template@odata.bind",
                    "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
                }
            },
            Members = new List<ConversationMember>
            {
                new AadUserConversationMember
                {
                    Roles = new List<string>{ "owner" },
                    AdditionalData = new Dictionary<string, object>
                    {
                        {"user@odata.bind", $"https://graph.microsoft.com/v1.0/users/{ownerUserId}"}
                    }
                }
            }
        };

        var result = await _graphClient.Teams.PostAsync(team);
        return result;
    }

    public async Task UpdateTeamAsync(string teamId, string displayName, string description)
    {
        var team = new Team
        {
            DisplayName = displayName,
            Description = description
        };

        await _graphClient.Teams[teamId].PatchAsync(team);
    }

    public async Task DeleteTeamAsync(string teamId)
    {
        await _graphClient.Groups[teamId].DeleteAsync();
    }

    public async Task AddTeamOwnerAsync(string teamId, string userId)
    {
        var referenceBody = new ReferenceCreate
        {
            OdataId = $"https://graph.microsoft.com/v1.0/directoryObjects/{userId}"
        };

        await _graphClient.Groups[teamId].Owners.Ref.PostAsync(referenceBody);
    }

    public async Task RemoveTeamOwnerAsync(string teamId, string userId)
    {
        await _graphClient.Groups[teamId].Owners[userId].Ref.DeleteAsync();
    }

    public async Task<List<ConversationMember>> ListTeamMembersAsync(string teamId)
    {
        var members = await _graphClient.Teams[teamId].Members.GetAsync();
        return members?.Value?.ToList() ?? new List<ConversationMember>();
    }

    public async Task AddTeamMemberAsync(string teamId, string userId, string roles = "member")
    {
        var conversationMember = new AadUserConversationMember
        {
            OdataType = "#microsoft.graph.aadUserConversationMember",
            Roles = new List<string> { roles },
            UserId = userId
        };

        await _graphClient.Teams[teamId].Members.PostAsync(conversationMember);
    }

    public async Task RemoveTeamMemberAsync(string teamId, string memberId)
    {
        await _graphClient.Teams[teamId].Members[memberId].DeleteAsync();
    }

    public async Task UpdateTeamMemberAsync(string teamId, string memberId, List<string> roles)
    {
        var conversationMember = new AadUserConversationMember
        {
            Roles = roles
        };

        await _graphClient.Teams[teamId].Members[memberId].PatchAsync(conversationMember);
    }

    public async Task<List<Channel>> ListChannelsAsync(string teamId)
    {
        var channels = await _graphClient.Teams[teamId].Channels.GetAsync();
        return channels?.Value?.ToList() ?? new List<Channel>();
    }

    public async Task<Channel?> GetChannelAsync(string teamId, string channelId)
    {
        var channel = await _graphClient.Teams[teamId].Channels[channelId].GetAsync();
        return channel;
    }

    public async Task<Channel?> CreateChannelAsync(string teamId, string displayName, string description, ChannelMembershipType membershipType = ChannelMembershipType.Standard)
    {
        var channel = new Channel
        {
            DisplayName = displayName,
            Description = description,
            MembershipType = membershipType
        };

        var result = await _graphClient.Teams[teamId].Channels.PostAsync(channel);
        return result;
    }

    public async Task UpdateChannelAsync(string teamId, string channelId, string displayName, string description)
    {
        var channel = new Channel
        {
            DisplayName = displayName,
            Description = description
        };

        await _graphClient.Teams[teamId].Channels[channelId].PatchAsync(channel);
    }

    public async Task DeleteChannelAsync(string teamId, string channelId)
    {
        await _graphClient.Teams[teamId].Channels[channelId].DeleteAsync();
    }

    public async Task<List<DriveItem>> ListChannelFilesAsync(string teamId, string channelId)
    {
        var folder = await _graphClient.Teams[teamId].Channels[channelId].FilesFolder.GetAsync();
        var files = await _graphClient.Drives[folder.ParentReference.DriveId].Items[folder.Id].Children.GetAsync();
        return files.Value;
    }

    public async Task<DriveItem?> GetChannelFileAsync(string teamId, string channelId, string fileName)
    {
        var folder = await _graphClient.Teams[teamId].Channels[channelId].FilesFolder.GetAsync();
        var file = await _graphClient.Drives[folder.ParentReference.DriveId].Items[folder.Id].Children[fileName].GetAsync();
        return file;
    }

    public async Task<DriveItem?> CreateChannelFileAsync(string teamId, string channelId, string fileName, Stream fileContent)
    {
        var filesFolder = await _graphClient.Teams[teamId].Channels[channelId].FilesFolder.GetAsync();
        if (filesFolder?.ParentReference?.DriveId == null)
        {
            throw new InvalidOperationException("Could not get channel files folder");
        }

        var driveId = filesFolder.ParentReference.DriveId;
        var folderId = filesFolder.Id;

        // Upload file directly with content using the Put method
        var uploadedFile = await _graphClient.Drives[driveId].Items[folderId].Children[fileName].Content
            .PutAsync(fileContent);

        return uploadedFile;
    }

    public async Task UpdateChannelFileAsync(string teamId, string channelId, string fileId, Stream fileContent)
    {
        var filesFolder = await _graphClient.Teams[teamId].Channels[channelId].FilesFolder.GetAsync();
        if (filesFolder?.ParentReference?.DriveId == null)
        {
            throw new InvalidOperationException("Could not get channel files folder");
        }

        var driveId = filesFolder.ParentReference.DriveId;
        await _graphClient.Drives[driveId].Items[fileId].Content
            .PutAsync(fileContent);
    }

    public async Task DeleteChannelFileAsync(string teamId, string channelId, string fileId)
    {
        var filesFolder = await _graphClient.Teams[teamId].Channels[channelId].FilesFolder.GetAsync();
        if (filesFolder?.ParentReference?.DriveId == null)
        {
            throw new InvalidOperationException("Could not get channel files folder");
        }

        var driveId = filesFolder.ParentReference.DriveId;
        await _graphClient.Drives[driveId].Items[fileId].DeleteAsync();
    }

    public async Task<Site> GetSiteAsync(string tenentName, string siteName)
    {
        var site = await _graphClient.Sites[$"{tenentName}.sharepoint.com:/sites/{siteName}"].GetAsync();
        return site;
    }

    public async Task<Site> GetSiteAsync(string siteId)
    {
        var site = await _graphClient.Sites[siteId].GetAsync();
        return site;
    }
}
