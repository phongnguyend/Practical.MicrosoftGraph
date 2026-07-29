using Microsoft.Extensions.Caching.Memory;
using Microsoft.Graph;
using Microsoft.Graph.Drives.Item.Items.Item.Checkin;
using Microsoft.Graph.Models;
using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using UriHelper;

namespace Practical.MicrosoftGraph.Sites;

public class SharePointOnlineClient
{
    private readonly GraphServiceClient _client;
    private readonly SharePointOnlineClientOptions _options;
    private readonly IMemoryCache _memoryCache;

    public SharePointOnlineClient(SharePointOnlineClientOptions options, IMemoryCache memoryCache)
    {
        _client = options.CreateGraphServiceClient();
        _options = options;
        _memoryCache = memoryCache;
    }

    private string GetRelativePath(string fileLocation)
    {
        return UriPath.Combine(_options.Path, fileLocation);
    }

    private async Task<Site> GetSiteAsync(CancellationToken cancellationToken = default)
    {
        var cacheKey = $"SharePointSite_{_options.SiteHostname}_{_options.SitePath}";

        var cacheOptions = new MemoryCacheEntryOptions
        {
            AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(10),
            SlidingExpiration = TimeSpan.FromMinutes(5)
        };

        return await _memoryCache.GetOrSetAsync(cacheKey, async () =>
        {
            var site = await _client.Sites[$"{_options.SiteHostname}:{_options.SitePath}"]
                .GetAsync(cancellationToken: cancellationToken);
            return site;
        }, cacheOptions);
    }

    private async Task<Drive> GetDocumentLibraryAsync(CancellationToken cancellationToken = default)
    {
        var cacheKey = $"SharePointDrive_{_options.SiteHostname}_{_options.SitePath}_{_options.DocumentLibraryName}";

        var cacheOptions = new MemoryCacheEntryOptions
        {
            AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(10),
            SlidingExpiration = TimeSpan.FromMinutes(5)
        };

        return await _memoryCache.GetOrSetAsync(cacheKey, async () =>
        {
            var site = await GetSiteAsync(cancellationToken);

            var drives = await _client.Sites[site.Id].Drives
                .GetAsync(cancellationToken: cancellationToken);

            var drive = drives.Value.FirstOrDefault(d =>
                d.Name.Equals(_options.DocumentLibraryName, StringComparison.OrdinalIgnoreCase));

            if (drive == null)
            {
                throw new InvalidOperationException($"Document library '{_options.DocumentLibraryName}' not found");
            }

            return drive;
        }, cacheOptions);
    }

    public async Task CreateAsync(string fileLocation, Stream stream, CancellationToken cancellationToken = default)
    {
        var drive = await GetDocumentLibraryAsync(cancellationToken);
        var relativePath = GetRelativePath(fileLocation);

        // Upload the file directly - SharePoint will create directories as needed
        await _client.Drives[drive.Id].Root
            .ItemWithPath(relativePath)
            .Content
            .PutAsync(stream, cancellationToken: cancellationToken);

        // Auto checkin the file if required
        if (_options.CheckinRequired)
        {
            await _client.Drives[drive.Id].Root
                .ItemWithPath(relativePath)
                .Checkin.PostAsync(new CheckinPostRequestBody
                {
                    Comment = "Checkin",
                }, cancellationToken: cancellationToken);
        }
    }

    public async Task DeleteAsync(string fileLocation, CancellationToken cancellationToken = default)
    {
        var drive = await GetDocumentLibraryAsync(cancellationToken);
        var relativePath = GetRelativePath(fileLocation);

        await _client.Drives[drive.Id].Root
            .ItemWithPath(relativePath)
            .DeleteAsync(cancellationToken: cancellationToken);
    }

    public async Task<byte[]> ReadAsync(string fileLocation, CancellationToken cancellationToken = default)
    {
        using var stream = new MemoryStream();
        await DownloadAsync(fileLocation, stream, cancellationToken);
        return stream.ToArray();
    }

    public async Task DownloadAsync(string fileLocation, string path, CancellationToken cancellationToken = default)
    {
        using var fileStream = File.Create(path);
        await DownloadAsync(fileLocation, fileStream, cancellationToken);
    }

    public async Task DownloadAsync(string fileLocation, Stream stream, CancellationToken cancellationToken = default)
    {
        var drive = await GetDocumentLibraryAsync(cancellationToken);
        var relativePath = GetRelativePath(fileLocation);

        var contentStream = await _client.Drives[drive.Id].Root
            .ItemWithPath(relativePath)
            .Content
            .GetAsync(cancellationToken: cancellationToken);

        await contentStream.CopyToAsync(stream, cancellationToken);
    }

    public async Task<System.Collections.Generic.List<Permission>> GetPermissionsAsync(string fileLocation, CancellationToken cancellationToken = default)
    {
        var drive = await GetDocumentLibraryAsync(cancellationToken);
        var relativePath = GetRelativePath(fileLocation);

        var permissions = await _client.Drives[drive.Id].Root
            .ItemWithPath(relativePath)
            .Permissions
            .GetAsync(cancellationToken: cancellationToken);

        return permissions.Value;
    }

    public async Task ArchiveAsync(string fileLocation, CancellationToken cancellationToken = default)
    {
        // SharePoint Online doesn't have a direct archive concept like S3
        // This could be implemented by moving files to an "Archive" folder
        // For now, we'll just implement it as a no-op
        await Task.CompletedTask;
    }

    public async Task UnArchiveAsync(string fileLocation, CancellationToken cancellationToken = default)
    {
        // SharePoint Online doesn't have a direct archive concept like S3
        // This would move files back from an "Archive" folder
        // For now, we'll just implement it as a no-op
        await Task.CompletedTask;
    }

    public void Dispose()
    {
        _client?.Dispose();
    }
}