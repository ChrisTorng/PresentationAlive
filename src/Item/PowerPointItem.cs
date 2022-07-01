namespace PresentationAlive;

public class PowerPointItem : IItem
{
    public PowerPointItem(string displayName, string path)
    {
        this.DisplayName = displayName;
        this.Path = path;
    }

    public ItemType ItemType { get; } = ItemType.PowerPoint;

    public string DisplayName { get; }

    public string Path { get; }

    public override string ToString() =>
        "PowerPoint: " + this.DisplayName;
}