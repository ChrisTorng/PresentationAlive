namespace PresentationAlive.ItemLib;

public interface IItem
{
    ItemType ItemType { get; }

    string DisplayName { get; }

    static void Open() { }

    void Start();

    void Next();

    event EventHandler Stopped;

    void Stop();

    static void Close() { }
}