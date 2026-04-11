using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Practical.MicrosoftGraph.TeamsChats;
using System;
using System.Linq;

// Load configuration from user secrets
var configuration = new ConfigurationBuilder()
    .AddUserSecrets<Program>()
    .Build();

var tenantId = configuration["TenantId"];
var clientId = configuration["ClientId"];
var clientSecret = configuration["ClientSecret"];

// Create credential and graph client
var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
var graphClient = new GraphServiceClient(credential, ["https://graph.microsoft.com/.default"]);

var teamsChatsManager = new TeamsChatsManager(graphClient);

var team = await teamsChatsManager.GetTeamByNameAsync("Test");
var channels = await teamsChatsManager.ListChannelsAsync(team.Id);
var channel = channels.FirstOrDefault();


// Example usage:

// 1. Create a group chat
//var groupChat = await teamsChatsManager.CreateGroupChatAsync("My Group Chat", "test@test.onmicrosoft.com", ["test@test.onmicrosoft.com"]);
//Console.WriteLine($"Created group chat: {groupChat?.Id}");

// chatId: 19:5b2c9564-02fa-4187-8e58-f2bba39ba78c_b9879bc5-552a-4e54-a8d8-543c2e641b10@unq.gbl.spaces
// chatId: 19:fd72923a5f234b8fb6514661a4211a6d@thread.v2

var chat = await teamsChatsManager.GetChatAsync("19:fd72923a5f234b8fb6514661a4211a6d@thread.v2");

var messages = await teamsChatsManager.ListChatMessagesAsync(chat.Id);

foreach (var message in messages)
{
    var sender = message.From?.User?.DisplayName ?? "Unknown";
    var sentTime = message.CreatedDateTime?.ToString("g") ?? "Unknown";
    Console.WriteLine($"[{sentTime}] [MessageType: {message.MessageType}] [Sender: {sender}]: {message.Body.Content}");
}

messages = await teamsChatsManager.ListChannelMessagesAsync(team.Id, channel.Id);

foreach (var message in messages)
{
    var sender = message.From?.User?.DisplayName ?? "Unknown";
    var sentTime = message.CreatedDateTime?.ToString("g") ?? "Unknown";
    Console.WriteLine($"[{sentTime}] [MessageType: {message.MessageType}] [Sender: {sender}]: {message.Body.Content}");
}

Console.WriteLine("Teams Chats Manager initialized successfully!");
