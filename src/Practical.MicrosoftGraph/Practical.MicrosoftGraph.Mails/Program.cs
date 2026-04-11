using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Practical.MicrosoftGraph.Mails;
using System;
using System.Collections.Generic;

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
var emailManager = new EmailManager(graphClient);

var userName = "xxx@xxx.onmicrosoft.com";
var recipientEmail = "xxx@xxx.onmicrosoft.com";

// List existing messages
Console.WriteLine("Listing recent messages...");
var messages = await emailManager.ListMessagesAsync(userName, top: 10);
Console.WriteLine($"Found {messages.Count} message(s)");
foreach (var message in messages)
{
    Console.WriteLine($"- Subject: {message.Subject}, From: {message.From?.EmailAddress.Address}");
}

// List draft messages
Console.WriteLine("\nListing draft messages...");
var draftMessages = await emailManager.ListDraftMessagesAsync(userName, top: 10);
Console.WriteLine($"Found {draftMessages.Count} draft message(s)");
foreach (var draft in draftMessages)
{
    Console.WriteLine($"- Subject: {draft.Subject}, IsDraft: {draft.IsDraft}");
}

// List sent messages
Console.WriteLine("\nListing sent messages...");
var sentMessages = await emailManager.ListSentMessagesAsync(userName, top: 10);
Console.WriteLine($"Found {sentMessages.Count} sent message(s)");
foreach (var sent in sentMessages)
{
    Console.WriteLine($"- Subject: {sent.Subject}, From: {sent.From?.EmailAddress.Address}");
}

// Create a draft email
Console.WriteLine("\nCreating a draft email...");
var draftEmail = await emailManager.CreateDraftEmailAsync(
    userName,
    subject: "Test Draft Email " + DateTime.Now.ToString("yyyyMMdd_HHmmss"),
    bodyContent: "This is a draft email created from C#.",
    toRecipients: new List<string> { recipientEmail }
);

if (draftEmail != null)
{
    Console.WriteLine($"Draft email created with ID: {draftEmail.Id}");
    var draftMessageId = draftEmail.Id;

    // Add attachment to the draft email
    Console.WriteLine("\nAdding attachment to the draft email...");
    var fileContent = System.Text.Encoding.UTF8.GetBytes("This is a test attachment content.");
    var attachment = await emailManager.AddAttachmentAsync(userName, draftMessageId, "test-attachment.txt", fileContent, "text/plain");
    if (attachment != null)
    {
        Console.WriteLine($"Attachment added: {attachment.Name}");
    }

    // List attachments
    Console.WriteLine("\nListing attachments in the draft email...");
    var attachments = await emailManager.ListAttachmentsAsync(userName, draftMessageId);
    Console.WriteLine($"Found {attachments.Count} attachment(s)");
    foreach (var att in attachments)
    {
        Console.WriteLine($"- {att.Name}");
    }

    // Update draft email
    Console.WriteLine("\nUpdating the draft email...");
    await emailManager.UpdateDraftEmailAsync(
        userName,
        draftMessageId,
        subject: "Updated Draft Email " + DateTime.Now.ToString("yyyyMMdd_HHmmss"),
        bodyContent: "This draft email has been updated.",
        toRecipients: new List<string> { recipientEmail }
    );
    Console.WriteLine("Draft email updated successfully");

    // Send the draft email
    Console.WriteLine("\nSending the draft email...");
    await emailManager.SendDraftEmailAsync(userName, draftMessageId);
    Console.WriteLine("Draft email sent successfully");
}

// Send a new email with attachments directly
Console.WriteLine("\nSending an email with attachment...");
var attachmentContent = System.Text.Encoding.UTF8.GetBytes("Sample attachment for new email");
var fileAttachment = new Microsoft.Graph.Models.FileAttachment
{
    OdataType = "#microsoft.graph.fileAttachment",
    Name = "sample.txt",
    ContentBytes = attachmentContent,
    ContentType = "text/plain"
};

var sentEmail = await emailManager.SendEmailAsync(
    userName,
    subject: "Test Email with Attachment " + DateTime.Now.ToString("yyyyMMdd_HHmmss"),
    bodyContent: "This email includes an attachment.",
    toRecipients: new List<string> { recipientEmail },
    attachments: new List<Microsoft.Graph.Models.FileAttachment> { fileAttachment }
);

if (sentEmail != null)
{
    Console.WriteLine("Email with attachment sent successfully");
}

// Search for messages
Console.WriteLine("\nSearching for messages...");
var searchResults = await emailManager.SearchMessagesAsync(userName, "Test", top: 5);
Console.WriteLine($"Found {searchResults.Count} message(s) matching 'Test'");
foreach (var result in searchResults)
{
    Console.WriteLine($"- {result.Subject}");
}

Console.WriteLine("\n=== Operations Complete ===");
