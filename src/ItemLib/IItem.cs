namespace PresentationAlive.ItemLib;

public interface IItem : IDisposable
{
    ItemType ItemType { get; }

    string DisplayName { get; }

    string Path { get; }

    IEnumerable<IItem>? SubItems { get; }

    void Open();

    void Start();

    event EventHandler? Stopped;

    void Stop();

    void Close();
}