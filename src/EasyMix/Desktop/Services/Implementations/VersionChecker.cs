using Desktop.Models;
using Desktop.Services.Interfaces;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Net.Http.Json;
using System.Text.Json;

namespace Desktop.Services.Implementations
{
    public class VersionChecker : IVersionChecker
    {
        public async Task<bool> CheckAndUpdateAsync(string jsonPath)
        {
            if (!File.Exists(jsonPath)) return false;

            var local = await ReadLocalAsync(jsonPath);
            if (local is null || string.IsNullOrEmpty(local.VersionUrl)) return false;

            var remote = await ReadRemoteAsync(local.VersionUrl);
            if (remote is null || local.Version == remote.Version) return false;

            Console.WriteLine($"🔄 Có bản mới cho {local.AppName}: {local.Version} → {remote.Version}");

            var http = new HttpClient();
            var zipData = await http.GetByteArrayAsync(remote.ZipUrl);
            var zipFile = Path.Combine(Path.GetDirectoryName(jsonPath)!, remote.File);
            await File.WriteAllBytesAsync(zipFile, zipData);

            ZipFile.ExtractToDirectory(zipFile, Path.GetDirectoryName(jsonPath)!, true);
            File.Delete(zipFile);

            Console.WriteLine($"✅ Đã cập nhật {remote.AppName}!");

            return true;
        }

        public async Task<bool> CheckAsync(string jsonPath)
        {
            if (!File.Exists(jsonPath)) return false;

            var local = await ReadLocalAsync(jsonPath);
            if (local is null || string.IsNullOrEmpty(local.VersionUrl)) return false;

            var remote = await ReadRemoteAsync(local.VersionUrl);
            if (remote is null || local.Version == remote.Version) return false;

            return true;
        }

        public async Task<VersionInfo?> ReadLocalAsync(string path)
        {
            var json = await File.ReadAllTextAsync(path);
            return JsonSerializer.Deserialize<VersionInfo>(json);
        }

        public async Task<VersionInfo?> ReadRemoteAsync(string url)
        {
            var http = new HttpClient();
            return await http.GetFromJsonAsync<VersionInfo>(url);
        }
    }
}
