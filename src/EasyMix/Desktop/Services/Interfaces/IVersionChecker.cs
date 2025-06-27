using Desktop.Models;

namespace Desktop.Services.Interfaces
{
    public interface IVersionChecker
    {
        Task<bool> CheckAndUpdateAsync(string jsonPath);
        Task<bool> CheckAsync(string jsonPath);
        Task<VersionInfo?> ReadLocalAsync(string path);
        Task<VersionInfo?> ReadRemoteAsync(string url);
    }
}
