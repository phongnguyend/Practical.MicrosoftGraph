### Permissions:
| API / Permissions name                 | Type        | Description                                                          |
|----------------------------------------|-------------|----------------------------------------------------------------------|
| Channel.ReadBasic.All                  | Application | Read the names and descriptions of all channels                      |
| ChannelMessage.Read.All                | Application | Read all channel messages                                            |
| Chat.Create                            | Application | Create chats and add members                                         |
| Chat.ReadWrite                         | Application | Read and write to user's chats                                       |
| Chat.ReadWrite.All                     | Application | Read and write all chats                                             |
| ChatMessage.Send                       | Application | Send chat messages                                                   |
| User.ReadBasic.All                     | Application | Read all users' basic profiles                                       |

### TeamsChatsManager Overview:
The `TeamsChatsManager` class provides methods to manage Teams chats and send messages through various channels:

#### Group Chat Methods:
- `ListChatsAsync()` - List all chats for the current user
- `GetChatAsync(string chatId)` - Get a specific chat by ID
- `CreateGroupChatAsync(string topic, List<string> userIds)` - Create a new group chat and add members
- `ListChatMembersAsync(string chatId)` - List members in a chat
- `AddChatMemberAsync(string chatId, string userId)` - Add a user to an existing chat
- `RemoveChatMemberAsync(string chatId, string memberId)` - Remove a member from a chat

#### Messaging Methods:
- `SendMessageToGroupChatAsync(string chatId, string messageText)` - Send a message to a group chat
- `SendMessageToChannelAsync(string teamId, string channelId, string messageText)` - Send a message to a team channel
- `SendDirectMessageToUserAsync(string userId, string messageText)` - Send a direct message to a specific user
- `ListChatMessagesAsync(string chatId)` - List messages in a chat
- `ListChannelMessagesAsync(string teamId, string channelId)` - List messages in a channel
- `GetChatMessageAsync(string chatId, string messageId)` - Get a specific chat message
- `GetChannelMessageAsync(string teamId, string channelId, string messageId)` - Get a specific channel message
- `UpdateChatMessageAsync(string chatId, string messageId, string messageText)` - Update a chat message
- `UpdateChannelMessageAsync(string teamId, string channelId, string messageId, string messageText)` - Update a channel message
- `DeleteChatMessageAsync(string chatId, string messageId)` - Delete a chat message
- `DeleteChannelMessageAsync(string teamId, string channelId, string messageId)` - Delete a channel message

#### Direct Message Methods:
- `GetOrCreateDirectChatWithUserAsync(string userId)` - Get or create a 1-on-1 chat with a user
