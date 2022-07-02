namespace PresentationAlive.ItemLib;

public interface IItem
{
    ItemType ItemType { get; }

    string DisplayName { get; }

    void Start();

    void Next();

    event EventHandler Stopped;

    void Close();
}