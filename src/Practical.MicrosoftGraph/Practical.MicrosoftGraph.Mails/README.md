## Email Management

### Permissions:
| API / Permissions name                 | Type        | Description                                                          |
|----------------------------------------|-------------|----------------------------------------------------------------------|
| Mail.ReadWrite                         | Application | Read and write to mailboxes                                          |
| Mail.Send                              | Application | Send mail                                                             |
| User.ReadBasic.All                     | Application | Read all users' basic profiles                                       |

### EmailManager Overview:
The `EmailManager` class provides methods to manage emails, including creating drafts, sending emails with attachments, and managing email attachments:

#### Message Management Methods:
- `ListMessagesAsync(string userIdOrName, int top = 10)` - List recent messages from a user's mailbox
- `ListDraftMessagesAsync(string userIdOrName, int top = 10)` - List draft messages
- `ListSentMessagesAsync(string userIdOrName, int top = 10)` - List sent messages
- `GetMessageAsync(string userIdOrName, string messageId)` - Get a specific message by ID
- `SearchMessagesAsync(string userIdOrName, string searchQuery, int top = 10)` - Search for messages by subject

#### Draft Email Methods:
- `CreateDraftEmailAsync(string userIdOrName, string subject, string bodyContent, List<string> toRecipients, List<string>? ccRecipients = null, List<string>? bccRecipients = null, List<FileAttachment>? attachments = null)` - Create a draft email
- `UpdateDraftEmailAsync(string userIdOrName, string messageId, string subject, string bodyContent, List<string> toRecipients, List<string>? ccRecipients = null, List<string>? bccRecipients = null)` - Update an existing draft email
- `SendDraftEmailAsync(string userIdOrName, string messageId)` - Send a draft email

#### Send Email Methods:
- `SendEmailAsync(string userIdOrName, string subject, string bodyContent, List<string> toRecipients, List<string>? ccRecipients = null, List<string>? bccRecipients = null, List<FileAttachment>? attachments = null)` - Send an email directly (bypassing draft)

#### Delete Email Methods:
- `DeleteEmailAsync(string userIdOrName, string messageId)` - Delete an email message

#### Attachment Management Methods:
- `AddAttachmentAsync(string userIdOrName, string messageId, string fileName, byte[] fileContent, string contentType = "application/octet-stream")` - Add an attachment to a message
- `ListAttachmentsAsync(string userIdOrName, string messageId)` - List all attachments in a message
- `GetAttachmentAsync(string userIdOrName, string messageId, string attachmentId)` - Get a specific attachment
- `DeleteAttachmentAsync(string userIdOrName, string messageId, string attachmentId)` - Delete an attachment from a message

### Typical Usage Examples:

#### Send Email with Attachment
```csharp
var emailManager = new EmailManager(graphClient);
var attachmentContent = System.Text.Encoding.UTF8.GetBytes("File content here");
var fileAttachment = new FileAttachment
{
    OdataType = "#microsoft.graph.fileAttachment",
    Name = "file.txt",
    ContentBytes = attachmentContent,
    ContentType = "text/plain"
};

await emailManager.SendEmailAsync(
    "user@example.com",
    subject: "Hello",
    bodyContent: "Email body",
    toRecipients: new List<string> { "recipient@example.com" },
    attachments: new List<FileAttachment> { fileAttachment }
);
```

#### Create and Send Draft Email
```csharp
// Create a draft
var draft = await emailManager.CreateDraftEmailAsync(
    "user@example.com",
    subject: "Draft Email",
    bodyContent: "This is a draft",
    toRecipients: new List<string> { "recipient@example.com" }
);

// Add attachment
await emailManager.AddAttachmentAsync(
    "user@example.com",
    draft.Id,
    "attachment.txt",
    System.Text.Encoding.UTF8.GetBytes("Content")
);

// Send the draft
await emailManager.SendDraftEmailAsync("user@example.com", draft.Id);
