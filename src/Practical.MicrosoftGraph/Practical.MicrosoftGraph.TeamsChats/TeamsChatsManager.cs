using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Practical.MicrosoftGraph.TeamsChats;

public class TeamsChatsManager
{
    private readonly GraphServiceClient _graphClient;

    public TeamsChatsManager(GraphServiceClient graphClient)
    {
        _graphClient = graphClient;
    }

    public async Task<Team> GetTeamByNameAsync(string teamName)
    {
        var teams = await _graphClient.Teams.GetAsync(config => config.QueryParameters.Filter = $"displayName eq '{teamName}'");

        return teams?.Value?.SingleOrDefault();
    }

    public async Task<List<Channel>> ListChannelsAsync(string teamId)
    {
        var channels = await _graphClient.Teams[teamId].Channels.GetAsync();
        return channels?.Value?.ToList() ?? new List<Channel>();
    }

    public async Task<List<Chat>> ListChatsAsync()
    {
        var chats = await _graphClient.Chats.GetAsync();
        return chats?.Value?.ToList() ?? new List<Chat>();
    }

    public async Task<Chat?> GetChatAsync(string chatId)
    {
        var chat = await _graphClient.Chats[chatId].GetAsync();
        return chat;
    }

    public async Task<Chat?> CreateGroupChatAsync(string topic, string ownerUserId, List<string> userIds)
    {
        var members = new List<ConversationMember>();

        // Add current user as owner
        members.Add(new AadUserConversationMember
        {
            OdataType = "#microsoft.graph.aadUserConversationMember",
            Roles = new List<string> { "owner" },
            AdditionalData = new Dictionary<string, object>
                {
                    { "user@odata.bind", $"https://graph.microsoft.com/v1.0/users/{ownerUserId}" }
                }
        });

        // Add other users as members
        foreach (var userId in userIds)
        {
            members.Add(new AadUserConversationMember
            {
                OdataType = "#microsoft.graph.aadUserConversationMember",
                Roles = new List<string> { "guest" },
                AdditionalData = new Dictionary<string, object>
                {
                    { "user@odata.bind", $"https://graph.microsoft.com/v1.0/users/{userId}" }
                }
            });
        }

        var chat = new Chat
        {
            ChatType = ChatType.Group,
            Topic = topic,
            Members = members
        };

        var result = await _graphClient.Chats.PostAsync(chat);
        return result;
    }

    public async Task<List<ConversationMember>> ListChatMembersAsync(string chatId)
    {
        var members = await _graphClient.Chats[chatId].Members.GetAsync();
        return members?.Value?.ToList() ?? new List<ConversationMember>();
    }

    public async Task AddChatMemberAsync(string chatId, string userId)
    {
        var conversationMember = new AadUserConversationMember
        {
            OdataType = "#microsoft.graph.aadUserConversationMember",
            Roles = new List<string> { "member" },
            AdditionalData = new Dictionary<string, object>
            {
                { "user@odata.bind", $"https://graph.microsoft.com/v1.0/users/{userId}" }
            }
        };

        await _graphClient.Chats[chatId].Members.PostAsync(conversationMember);
    }

    public async Task RemoveChatMemberAsync(string chatId, string memberId)
    {
        await _graphClient.Chats[chatId].Members[memberId].DeleteAsync();
    }

    public async Task<ChatMessage?> SendMessageToGroupChatAsync(string chatId, string messageText)
    {
        var message = new ChatMessage
        {
            Body = new ItemBody
            {
                ContentType = BodyType.Html,
                Content = messageText
            }
        };

        var result = await _graphClient.Chats[chatId].Messages.PostAsync(message);
        return result;
    }

    public async Task<ChatMessage?> SendMessageToChannelAsync(string teamId, string channelId, string messageText)
    {
        var message = new ChatMessage
        {
            Body = new ItemBody
            {
                ContentType = BodyType.Html,
                Content = messageText
            }
        };

        var result = await _graphClient.Teams[teamId].Channels[channelId].Messages.PostAsync(message);
        return result;
    }

    public async Task<List<ChatMessage>> ListChatMessagesAsync(string chatId)
    {
        var messages = await _graphClient.Chats[chatId].Messages.GetAsync();
        return messages?.Value?.ToList() ?? new List<ChatMessage>();
    }

    public async Task<List<ChatMessage>> ListChannelMessagesAsync(string teamId, string channelId)
    {
        var messages = await _graphClient.Teams[teamId].Channels[channelId].Messages.GetAsync();
        return messages?.Value?.ToList() ?? new List<ChatMessage>();
    }

    public async Task<ChatMessage?> GetChatMessageAsync(string chatId, string messageId)
    {
        var message = await _graphClient.Chats[chatId].Messages[messageId].GetAsync();
        return message;
    }

    public async Task<ChatMessage?> GetChannelMessageAsync(string teamId, string channelId, string messageId)
    {
        var message = await _graphClient.Teams[teamId].Channels[channelId].Messages[messageId].GetAsync();
        return message;
    }
}
