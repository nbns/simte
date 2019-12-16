namespace simte.Common
{
    public interface IAttachable<out T>
    {
        T Attach();
    }
}
