namespace simte.Common
{
    public interface IApplyable<out T>
    {
        T Apply();
    }
}
