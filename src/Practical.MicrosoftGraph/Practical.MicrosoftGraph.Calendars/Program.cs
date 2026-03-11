using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Practical.MicrosoftGraph.Calendars;
using System;
using System.Collections.Generic;

var config = new ConfigurationBuilder()
//.AddJsonFile("appsettings.json")
.AddUserSecrets("473ed7c3-3710-46ab-a7f1-816a98fe18c6")
.Build();

// The client credentials flow requires that you request the
// /.default scope, and preconfigure your permissions on the
// app registration in Azure. An administrator must grant consent
// to those permissions beforehand.
var scopes = new[] { "https://graph.microsoft.com/.default" };

// Multi-tenant apps can use "common",
// single-tenant apps must use the tenant ID from the Azure portal
var tenantId = config["TenantId"];

// Values from app registration
var clientId = config["ClientId"];
var clientSecret = config["ClientSecret"];

// using Azure.Identity;
var options = new TokenCredentialOptions
{
    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
};

// https://docs.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
var clientSecretCredential = new ClientSecretCredential(
    tenantId, clientId, clientSecret, options);

var graphClient = new GraphServiceClient(clientSecretCredential, scopes);

var userName = config["Calendars:UserName"];
var attendee1 = config["Calendars:Attendee1"];
var attendee2 = config["Calendars:Attendee2"];
var attendee3 = config["Calendars:Attendee3"];

var userManager = new UserManager(graphClient);

//var users = await userManager.GetUsersAsync();
//foreach (var user in users)
//{
//    Console.WriteLine($"User: {user.DisplayName} ({user.Id}) ({user.Mail}) ({user.UserPrincipalName})");
//}

//var myUser = await userManager.GetUserAsync(userName);

var events = await userManager.GetEventsAsync(userName);

var start = new DateTimeTimeZone
{
    DateTime = DateTime.Now.ToString("o"),
    TimeZone = "Eastern Standard Time"   // Windows timezone name
};

var end = new DateTimeTimeZone
{
    DateTime = DateTime.Now.AddHours(1).ToString("o"),
    TimeZone = "Eastern Standard Time"
};

var @event = await userManager.CreateEventAsync(userName, "Book an Appointment Demo " + DateTime.Now.ToString("yyyyMMdd_HHmmss"), "Does noon work for you?", start, end, new List<Attendee>
                {
                    new Attendee
                    {
                        EmailAddress = new EmailAddress {Address = attendee1 },
                        Type = AttendeeType.Required,
                    },
                    new Attendee
                    {
                        EmailAddress = new EmailAddress {Address = attendee2 },
                        Type = AttendeeType.Required,
                    },
                    new Attendee
                    {
                        EmailAddress = new EmailAddress {Address = attendee3 },
                        Type = AttendeeType.Optional,
                    },
                }, isOnlineMeeting: true);

//events = await userManager.GetEventsAsync(userName);

//await userManager.DeleteEventAsync(userName, @event.Id);

//events = await userManager.GetEventsAsync(userName);

//var onlineMeetings = await graphClient.Users[userId].OnlineMeetings.PostAsync(new OnlineMeeting
//{
//    Subject = "Book an Appointment Demo"
//});

var startOfWeek = DateTime.Now.AddDays(-1);
var endOfWeek = startOfWeek.AddDays(7);

var searchEvents = await userManager.SearchEventsAsync(userName, startOfWeek, endOfWeek);

foreach (var searchEvent in searchEvents)
{
    Console.WriteLine($"Event: {searchEvent.Subject} ({searchEvent.Start.DateTime} - {searchEvent.End.DateTime})");
}
