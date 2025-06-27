using Desktop.Services.Interfaces;
using Octokit;

namespace Desktop.Services.Implementations
{
    public class GitHubService : IGitHubService
    {
        private readonly GitHubClient _client = new GitHubClient(new Octokit.ProductHeaderValue("WpfAutoUpdateApp"));
        private readonly string _owner = "TenChuRepo";
        private readonly string _repo = "TenRepo";
        public async Task<string> GetLatestReleaseTagAsync()
        {
            var releases = await _client.Repository.Release.GetAll(_owner, _repo);
            return releases.FirstOrDefault()?.TagName!;
        }
    }
}
