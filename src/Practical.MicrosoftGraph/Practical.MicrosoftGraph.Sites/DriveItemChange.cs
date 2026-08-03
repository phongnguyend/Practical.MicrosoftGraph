using Microsoft.Graph.Models;

namespace Practical.MicrosoftGraph.Sites;

public enum DriveItemChangeType
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

public class DriveItemChange
{
    public DriveItemChange(DriveItem item, DriveItemChangeType changeType)
    {
        Item = item;
        ChangeType = changeType;
    }

    public DriveItem Item { get; }

    public DriveItemChangeType ChangeType { get; }

    public DriveItemType ItemType => Item.Folder != null ? DriveItemType.Folder : DriveItemType.File;

    public string IdempotencyKey => $"{Item.Id}:{Item.ETag}";

    public static DriveItemChange Create(DriveItem item)
    {
        return new DriveItemChange(item, GetChangeType(item));
    }

    private static DriveItemChangeType GetChangeType(DriveItem item)
    {
        if (item.Deleted != null)
        {
            return DriveItemChangeType.Deleted;
        }

        // Graph delta does not explicitly flag created vs updated items, so we infer it by
        // comparing the creation and last modification timestamps of the item.
        if (item.CreatedDateTime.HasValue && item.LastModifiedDateTime.HasValue
            && item.CreatedDateTime.Value == item.LastModifiedDateTime.Value)
        {
            return DriveItemChangeType.Created;
        }

        return DriveItemChangeType.Updated;
    }
}
