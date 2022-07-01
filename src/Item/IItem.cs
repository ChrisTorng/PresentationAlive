namespace PresentationAlive;

public interface IItem
{
    ItemType ItemType { get; }

    string DisplayName { get; }
}