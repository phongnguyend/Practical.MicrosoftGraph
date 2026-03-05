using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.Models;
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

var userId = config["Calendars:UserId"];
var userName = config["Calendars:UserName"];
var domain = config["Domain"];
var attendee1 = config["Calendars:Attendee1"];
var attendee2 = config["Calendars:Attendee2"];
var attendee3 = config["Calendars:Attendee3"];

//var users = await graphClient.Users.Request().GetAsync();

var @event = await graphClient.Users[userName].Events.PostAsync(new Event
{
    Subject = "Book an Appointment Demo",
    Body = new ItemBody
    {
        ContentType = BodyType.Html,
        Content = "Does noon work for you?"
    },
    IsDraft = false,
    IsOnlineMeeting = true,
    Attendees = new List<Attendee>
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
                },
});

await graphClient.Users[userName].Events[@event.Id].DeleteAsync();

var onlineMeetings = await graphClient.Users[userId].OnlineMeetings.PostAsync(new OnlineMeeting
{
    Subject = "Book an Appointment Demo"
});

var allEvents = await graphClient.Users[userName].Events.GetAsync();

var startOfWeek = DateTime.Now;
var endOfWeek = startOfWeek.AddDays(7);

var events = await graphClient.Users[userName].CalendarView.GetAsync((requestConfiguration) =>
{
    requestConfiguration.QueryParameters.StartDateTime = startOfWeek.ToString("o");
    requestConfiguration.QueryParameters.EndDateTime = endOfWeek.ToString("o");
    requestConfiguration.QueryParameters.Select = new[] { "subject", "organizer", "start", "end" };
    requestConfiguration.QueryParameters.Orderby = new[] { "start/DateTime" };
});
