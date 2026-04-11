using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Practical.MicrosoftGraph.Mails;

public class EmailManager
{
    private readonly GraphServiceClient _graphClient;

    public EmailManager(GraphServiceClient graphClient)
    {
        _graphClient = graphClient;
    }

    public async Task<List<Message>> ListMessagesAsync(string userIdOrName, int top = 10)
    {
        var messages = await _graphClient.Users[userIdOrName].Messages.GetAsync(requestConfiguration =>
        {
            requestConfiguration.QueryParameters.Top = top;
        });
        return messages?.Value?.ToList() ?? new List<Message>();
    }

    public async Task<List<Message>> ListDraftMessagesAsync(string userIdOrName, int top = 10)
    {
        var messages = await _graphClient.Users[userIdOrName].Messages.GetAsync(requestConfiguration =>
        {
            requestConfiguration.QueryParameters.Filter = "isDraft eq true";
            requestConfiguration.QueryParameters.Top = top;
        });
        return messages?.Value?.ToList() ?? new List<Message>();
    }

    public async Task<List<Message>> ListSentMessagesAsync(string userIdOrName, int top = 10)
    {
        var messages = await _graphClient.Users[userIdOrName].MailFolders["sentItems"].Messages.GetAsync(requestConfiguration =>
        {
            requestConfiguration.QueryParameters.Top = top;
        });
        return messages?.Value?.ToList() ?? new List<Message>();
    }

    public async Task<Message?> GetMessageAsync(string userIdOrName, string messageId)
    {
        var message = await _graphClient.Users[userIdOrName].Messages[messageId].GetAsync();
        return message;
    }

    public async Task<List<Message>> SearchMessagesAsync(string userIdOrName, string searchQuery, int top = 10)
    {
        var messages = await _graphClient.Users[userIdOrName].Messages.GetAsync(requestConfiguration =>
        {
            requestConfiguration.QueryParameters.Search = $"\"subject:{searchQuery}\"";
            requestConfiguration.QueryParameters.Top = top;
        });
        return messages?.Value?.ToList() ?? new List<Message>();
    }

    public async Task<Message?> CreateDraftEmailAsync(string userIdOrName, string subject, string bodyContent, List<string> toRecipients, List<string>? ccRecipients = null, List<string>? bccRecipients = null, List<FileAttachment>? attachments = null)
    {
        var toAddresses = toRecipients.Select(x => new Recipient
        {
            EmailAddress = new EmailAddress { Address = x }
        }).ToList();

        var ccAddresses = ccRecipients?.Select(x => new Recipient
        {
            EmailAddress = new EmailAddress { Address = x }
        }).ToList() ?? new List<Recipient>();

        var bccAddresses = bccRecipients?.Select(x => new Recipient
        {
            EmailAddress = new EmailAddress { Address = x }
        }).ToList() ?? new List<Recipient>();

        var message = new Message
        {
            Subject = subject,
            Body = new ItemBody
            {
                ContentType = BodyType.Html,
                Content = bodyContent
            },
            ToRecipients = toAddresses,
            CcRecipients = ccAddresses,
            BccRecipients = bccAddresses
        };

        if (attachments != null && attachments.Count > 0)
        {
            message.Attachments = new List<Attachment>(attachments);
        }

        var draftMessage = await _graphClient.Users[userIdOrName].Messages.PostAsync(message);
        return draftMessage;
    }

    public async Task<Message?> SendEmailAsync(string userIdOrName, string subject, string bodyContent, List<string> toRecipients, List<string>? ccRecipients = null, List<string>? bccRecipients = null, List<FileAttachment>? attachments = null)
    {
        var toAddresses = toRecipients.Select(x => new Recipient
        {
            EmailAddress = new EmailAddress { Address = x }
        }).ToList();

        var ccAddresses = ccRecipients?.Select(x => new Recipient
        {
            EmailAddress = new EmailAddress { Address = x }
        }).ToList() ?? new List<Recipient>();

        var bccAddresses = bccRecipients?.Select(x => new Recipient
        {
            EmailAddress = new EmailAddress { Address = x }
        }).ToList() ?? new List<Recipient>();

        var message = new Message
        {
            Subject = subject,
            Body = new ItemBody
            {
                ContentType = BodyType.Html,
                Content = bodyContent
            },
            ToRecipients = toAddresses,
            CcRecipients = ccAddresses,
            BccRecipients = bccAddresses
        };

        if (attachments != null && attachments.Count > 0)
        {
            message.Attachments = new List<Attachment>(attachments);
        }

        await _graphClient.Users[userIdOrName].SendMail.PostAsync(new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody
        {
            Message = message
        });

        return message;
    }

    public async Task<Message?> SendDraftEmailAsync(string userIdOrName, string messageId)
    {
        await _graphClient.Users[userIdOrName].Messages[messageId].Send.PostAsync();
        return await GetMessageAsync(userIdOrName, messageId);
    }

    public async Task UpdateDraftEmailAsync(string userIdOrName, string messageId, string subject, string bodyContent, List<string> toRecipients, List<string>? ccRecipients = null, List<string>? bccRecipients = null)
    {
        var toAddresses = toRecipients.Select(x => new Recipient
        {
            EmailAddress = new EmailAddress { Address = x }
        }).ToList();

        var ccAddresses = ccRecipients?.Select(x => new Recipient
        {
            EmailAddress = new EmailAddress { Address = x }
        }).ToList() ?? new List<Recipient>();

        var bccAddresses = bccRecipients?.Select(x => new Recipient
        {
            EmailAddress = new EmailAddress { Address = x }
        }).ToList() ?? new List<Recipient>();

        var message = new Message
        {
            Subject = subject,
            Body = new ItemBody
            {
                ContentType = BodyType.Html,
                Content = bodyContent
            },
            ToRecipients = toAddresses,
            CcRecipients = ccAddresses,
            BccRecipients = bccAddresses
        };

        await _graphClient.Users[userIdOrName].Messages[messageId].PatchAsync(message);
    }

    public async Task DeleteEmailAsync(string userIdOrName, string messageId)
    {
        await _graphClient.Users[userIdOrName].Messages[messageId].DeleteAsync();
    }

    public async Task<FileAttachment?> AddAttachmentAsync(string userIdOrName, string messageId, string fileName, byte[] fileContent, string contentType = "application/octet-stream")
    {
        var attachment = new FileAttachment
        {
            OdataType = "#microsoft.graph.fileAttachment",
            Name = fileName,
            ContentBytes = fileContent,
            ContentType = contentType
        };

        var result = await _graphClient.Users[userIdOrName].Messages[messageId].Attachments.PostAsync(attachment);
        return result as FileAttachment;
    }

    public async Task<List<Attachment>> ListAttachmentsAsync(string userIdOrName, string messageId)
    {
        var attachments = await _graphClient.Users[userIdOrName].Messages[messageId].Attachments.GetAsync();
        return attachments?.Value?.ToList() ?? new List<Attachment>();
    }

    public async Task<Attachment?> GetAttachmentAsync(string userIdOrName, string messageId, string attachmentId)
    {
        var attachment = await _graphClient.Users[userIdOrName].Messages[messageId].Attachments[attachmentId].GetAsync();
        return attachment;
    }

    public async Task DeleteAttachmentAsync(string userIdOrName, string messageId, string attachmentId)
    {
        await _graphClient.Users[userIdOrName].Messages[messageId].Attachments[attachmentId].DeleteAsync();
    }
}
