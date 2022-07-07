namespace PresentationAlive.ItemLib;

public interface IItem : IDisposable
{
    ItemType ItemType { get; }

    string DisplayName { get; }

    void Open();

    void Start();

    void Next();

    event EventHandler Stopped;

    void Stop();

    void Close();
}