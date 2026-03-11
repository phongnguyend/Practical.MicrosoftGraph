using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Practical.MicrosoftGraph.Calendars;

public class UserManager
{
    private readonly GraphServiceClient _graphClient;

    public UserManager(GraphServiceClient graphClient)
    {
        _graphClient = graphClient;
    }

    public async Task<List<User>> GetUsersAsync()
    {
        var users = await _graphClient.Users.GetAsync();
        return users?.Value?.ToList() ?? new List<User>();
    }

    public async Task<User?> GetUserAsync(string userIdOrName)
    {
        var user = await _graphClient.Users[userIdOrName].GetAsync();
        return user;
    }

    public async Task<List<Event>> GetEventsAsync(string userIdOrName)
    {
        var events = await _graphClient.Users[userIdOrName].Events.GetAsync();
        return events?.Value?.ToList() ?? new List<Event>();
    }

    public async Task<List<Event>> SearchEventsAsync(string userIdOrName, DateTime start, DateTime end)
    {
        var events = await _graphClient.Users[userIdOrName].CalendarView.GetAsync((requestConfiguration) =>
        {
            requestConfiguration.QueryParameters.StartDateTime = start.ToString("o");
            requestConfiguration.QueryParameters.EndDateTime = end.ToString("o");
            requestConfiguration.QueryParameters.Select = ["subject", "organizer", "start", "end"];
            requestConfiguration.QueryParameters.Orderby = ["start/DateTime"];
        });

        return events?.Value?.ToList() ?? new List<Event>();
    }

    public async Task<Event?> CreateEventAsync(string userIdOrName, string subject, string bodyContent, DateTimeTimeZone start, DateTimeTimeZone end, List<Attendee>? attendees = null, bool isOnlineMeeting = false)
    {
        var @event = new Event
        {
            Subject = subject,
            Body = new ItemBody
            {
                ContentType = BodyType.Html,
                Content = bodyContent
            },
            Start = start,
            End = end,
            IsOnlineMeeting = isOnlineMeeting,
            Attendees = attendees ?? new List<Attendee>()
        };

        var result = await _graphClient.Users[userIdOrName].Events.PostAsync(@event);
        return result;
    }

    public async Task DeleteEventAsync(string userIdOrName, string eventId)
    {
        await _graphClient.Users[userIdOrName].Events[eventId].DeleteAsync();
    }
}
