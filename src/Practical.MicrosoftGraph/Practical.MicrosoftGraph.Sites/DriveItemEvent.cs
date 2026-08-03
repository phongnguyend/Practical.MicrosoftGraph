using Microsoft.Graph.Models;

namespace Practical.MicrosoftGraph.Sites;

public enum DriveItemEventType
{
    Created,
    Updated,
    Deleted,
}

public enum DriveItemType
{
    File,
    Folder,
}

public class DriveItemEvent
{
    public DriveItemEvent(DriveItem item, DriveItemEventType eventType)
    {
        Item = item;
        EventType = eventType;
    }

    public DriveItem Item { get; }

    public DriveItemEventType EventType { get; }

    public DriveItemType ItemType => Item.Folder != null ? DriveItemType.Folder : DriveItemType.File;

    public string IdempotencyKey => $"{Item.Id}:{Item.ETag}";

    public static DriveItemEvent Create(DriveItem item)
    {
        return new DriveItemEvent(item, GetEventType(item));
    }

    private static DriveItemEventType GetEventType(DriveItem item)
    {
        if (item.Deleted != null)
        {
            return DriveItemEventType.Deleted;
        }

        // Graph delta does not explicitly flag created vs updated items, so we infer it by
        // comparing the creation and last modification timestamps of the item.
        if (item.CreatedDateTime.HasValue && item.LastModifiedDateTime.HasValue
            && item.CreatedDateTime.Value == item.LastModifiedDateTime.Value)
        {
            return DriveItemEventType.Created;
        }

        return DriveItemEventType.Updated;
    }
}
