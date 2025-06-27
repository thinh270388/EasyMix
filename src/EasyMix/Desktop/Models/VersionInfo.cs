namespace Desktop.Models
{
    public record VersionInfo
    (
        string AppName,
        string Version,
        string File,
        string ZipUrl,
        string VersionUrl,
        string Sha,
        string Build,
        string? ChangeLog
    );
}
