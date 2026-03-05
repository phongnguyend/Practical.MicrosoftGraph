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

    public async Task<Team?> CreateTeamAsync(string displayName, string description)
    {
        var team = new Team
        {
            DisplayName = displayName,
            Description = description
        };

        var result = await _graphClient.Teams.PostAsync(team);
        return result;
    }

    public async Task<bool> UpdateTeamAsync(string teamId, string displayName, string description)
    {
        var team = new Team
        {
            DisplayName = displayName,
            Description = description
        };

        await _graphClient.Teams[teamId].PatchAsync(team);
        return true;
    }

    public async Task<bool> DeleteTeamAsync(string teamId)
    {
        await _graphClient.Teams[teamId].DeleteAsync();
        return true;
    }

    public async Task<List<DirectoryObject>> ListTeamOwnersAsync(string teamId)
    {
        var owners = await _graphClient.Groups[teamId].Owners.GetAsync();
        return owners?.Value?.ToList() ?? new List<DirectoryObject>();
    }

    public async Task<bool> AddTeamOwnerAsync(string teamId, string userId)
    {
        var referenceBody = new ReferenceCreate
        {
            OdataId = $"https://graph.microsoft.com/v1.0/directoryObjects/{userId}"
        };

        await _graphClient.Groups[teamId].Owners.Ref.PostAsync(referenceBody);
        return true;
    }

    public async Task<bool> RemoveTeamOwnerAsync(string teamId, string userId)
    {
        await _graphClient.Groups[teamId].Owners[userId].Ref.DeleteAsync();
        return true;
    }

    public async Task<List<ConversationMember>> ListTeamMembersAsync(string teamId)
    {
        var members = await _graphClient.Teams[teamId].Members.GetAsync();
        return members?.Value?.ToList() ?? new List<ConversationMember>();
    }

    public async Task<bool> AddTeamMemberAsync(string teamId, string userId, string roles = "member")
    {
        var conversationMember = new AadUserConversationMember
        {
            OdataType = "#microsoft.graph.aadUserConversationMember",
            Roles = new List<string> { roles },
            UserId = userId
        };

        await _graphClient.Teams[teamId].Members.PostAsync(conversationMember);
        return true;
    }

    public async Task<bool> RemoveTeamMemberAsync(string teamId, string memberId)
    {
        await _graphClient.Teams[teamId].Members[memberId].DeleteAsync();
        return true;
    }

    public async Task<bool> UpdateTeamMemberAsync(string teamId, string memberId, List<string> roles)
    {
        var conversationMember = new AadUserConversationMember
        {
            Roles = roles
        };

        await _graphClient.Teams[teamId].Members[memberId].PatchAsync(conversationMember);
        return true;
    }

    public async Task<List<(string Id, string DisplayName, List<string> Roles)>> ListTeamMembersWithRolesAsync(string teamId)
    {
        var members = await _graphClient.Teams[teamId].Members.GetAsync();
        var memberList = new List<(string Id, string DisplayName, List<string> Roles)>();

        if (members?.Value != null)
        {
            foreach (var member in members.Value)
            {
                if (member is AadUserConversationMember aadMember)
                {
                    memberList.Add((
                        Id: aadMember.Id,
                        DisplayName: aadMember.DisplayName ?? "Unknown",
                        Roles: aadMember.Roles?.ToList() ?? new List<string>()
                    ));
                }
            }
        }

        return memberList;
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

    public async Task<bool> UpdateChannelAsync(string teamId, string channelId, string displayName, string description)
    {
        var channel = new Channel
        {
            DisplayName = displayName,
            Description = description
        };

        await _graphClient.Teams[teamId].Channels[channelId].PatchAsync(channel);
        return true;
    }

    public async Task<bool> DeleteChannelAsync(string teamId, string channelId)
    {
        await _graphClient.Teams[teamId].Channels[channelId].DeleteAsync();
        return true;
    }

    public async Task<List<DriveItem>> ListChannelFilesAsync(string teamId, string channelId)
    {
        var files = await _graphClient.Teams[teamId].Channels[channelId].FilesFolder.GetAsync();
        return files?.Children?.ToList() ?? new List<DriveItem>();
    }

    public async Task<DriveItem?> GetChannelFileAsync(string teamId, string channelId, string fileId)
    {
        var file = await _graphClient.Drives[teamId].Items[fileId].GetAsync();
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

        var file = await _graphClient.Drives[driveId].Items[folderId].Children
            .PostAsync(new DriveItem
            {
                Name = fileName
            });

        if (file != null)
        {
            await _graphClient.Drives[driveId].Items[file.Id].Content
                .PutAsync(fileContent);
        }

        return file;
    }

    public async Task<bool> UpdateChannelFileAsync(string teamId, string channelId, string fileId, Stream fileContent)
    {
        var filesFolder = await _graphClient.Teams[teamId].Channels[channelId].FilesFolder.GetAsync();
        if (filesFolder?.ParentReference?.DriveId == null)
        {
            throw new InvalidOperationException("Could not get channel files folder");
        }

        var driveId = filesFolder.ParentReference.DriveId;
        await _graphClient.Drives[driveId].Items[fileId].Content
            .PutAsync(fileContent);
        return true;
    }

    public async Task<bool> DeleteChannelFileAsync(string teamId, string channelId, string fileId)
    {
        var filesFolder = await _graphClient.Teams[teamId].Channels[channelId].FilesFolder.GetAsync();
        if (filesFolder?.ParentReference?.DriveId == null)
        {
            throw new InvalidOperationException("Could not get channel files folder");
        }

        var driveId = filesFolder.ParentReference.DriveId;
        await _graphClient.Drives[driveId].Items[fileId].DeleteAsync();
        return true;
    }
}
