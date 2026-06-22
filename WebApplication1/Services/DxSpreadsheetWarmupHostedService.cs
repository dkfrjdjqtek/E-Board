using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Hosting.Server.Features;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Net.Security;
using System.Threading;
using System.Threading.Tasks;

namespace WebApplication1.Services
{
    // 2026.06.15 Added: DevExpress Spreadsheet cold-start 비용을 클라이언트가 아니라 서버 프로세스 시작 후 1회만 부담하도록 처리한다.
    public sealed class DxSpreadsheetWarmupHostedService : IHostedService
    {
        private readonly IHostApplicationLifetime _lifetime;
        private readonly IServer _server;
        private readonly IConfiguration _configuration;
        private readonly ILogger<DxSpreadsheetWarmupHostedService> _log;
        private CancellationTokenSource? _cts;
        private int _started;

        public DxSpreadsheetWarmupHostedService(
            IHostApplicationLifetime lifetime,
            IServer server,
            IConfiguration configuration,
            ILogger<DxSpreadsheetWarmupHostedService> log)
        {
            _lifetime = lifetime;
            _server = server;
            _configuration = configuration;
            _log = log;
        }

        public Task StartAsync(CancellationToken cancellationToken)
        {
            _cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);

            _lifetime.ApplicationStarted.Register(() =>
            {
                _ = Task.Run(() => RunWarmupOnceAsync(_cts.Token));
            });

            return Task.CompletedTask;
        }

        public Task StopAsync(CancellationToken cancellationToken)
        {
            try
            {
                _cts?.Cancel();
                _cts?.Dispose();
            }
            catch
            {
            }

            return Task.CompletedTask;
        }

        private async Task RunWarmupOnceAsync(CancellationToken cancellationToken)
        {
            if (Interlocked.Exchange(ref _started, 1) == 1)
                return;

            var totalSw = Stopwatch.StartNew();

            try
            {
                var delayMs = GetIntSetting("DxSpreadsheetWarmup:DelayMilliseconds", 3000);
                if (delayMs > 0)
                    await Task.Delay(delayMs, cancellationToken);

                var baseUrl = ResolveBaseUrl();
                if (string.IsNullOrWhiteSpace(baseUrl))
                {
                    _log.LogWarning("DX-WARMUP skipped reason=BaseUrlNotResolved");
                    return;
                }

                using var handler = new HttpClientHandler
                {
                    ServerCertificateCustomValidationCallback = (message, certificate, chain, errors) =>
                    {
                        if (errors == SslPolicyErrors.None)
                            return true;

                        var host = message?.RequestUri?.Host ?? string.Empty;
                        return string.Equals(host, "localhost", StringComparison.OrdinalIgnoreCase)
                            || string.Equals(host, "127.0.0.1", StringComparison.OrdinalIgnoreCase)
                            || string.Equals(host, "::1", StringComparison.OrdinalIgnoreCase);
                    }
                };

                using var client = new HttpClient(handler)
                {
                    Timeout = TimeSpan.FromSeconds(GetIntSetting("DxSpreadsheetWarmup:TimeoutSeconds", 90))
                };

                await CallWarmupUrlAsync(
                    client,
                    "DX-SPREADSHEET-WARMUP",
                    CombineUrl(baseUrl, "/Doc/DetailDXWarmup?source=hosted"),
                    cancellationToken);

                await CallWarmupUrlAsync(
                    client,
                    "DX-COMPOSE-WARMUP",
                    CombineUrl(baseUrl, "/Doc/ComposeDXWarmup?source=hosted"),
                    cancellationToken);

                totalSw.Stop();

                _log.LogInformation(
                    "DX-WARMUP completed elapsedMs={ElapsedMs}",
                    totalSw.Elapsed.TotalMilliseconds);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                totalSw.Stop();
                _log.LogInformation("DX-WARMUP canceled elapsedMs={ElapsedMs}", totalSw.Elapsed.TotalMilliseconds);
            }
            catch (Exception ex)
            {
                totalSw.Stop();
                _log.LogWarning(ex, "DX-WARMUP failed elapsedMs={ElapsedMs}", totalSw.Elapsed.TotalMilliseconds);
            }
        }

        // 2026.06.15 Added: warm-up URL 호출 공통 처리를 추가하여 DetailDX 와 ComposeDX 예열 로그를 분리한다.
        private async Task CallWarmupUrlAsync(
            HttpClient client,
            string logCode,
            string url,
            CancellationToken cancellationToken)
        {
            var sw = Stopwatch.StartNew();

            if (string.IsNullOrWhiteSpace(url))
            {
                _log.LogWarning("{LogCode} skipped reason=UrlEmpty", logCode);
                return;
            }

            _log.LogInformation("{LogCode} start url={Url}", logCode, url);

            using var response = await client.GetAsync(
                url,
                HttpCompletionOption.ResponseHeadersRead,
                cancellationToken);

            sw.Stop();

            if (!response.IsSuccessStatusCode)
            {
                _log.LogWarning(
                    "{LogCode} failed statusCode={StatusCode} elapsedMs={ElapsedMs} url={Url}",
                    logCode,
                    (int)response.StatusCode,
                    sw.Elapsed.TotalMilliseconds,
                    url);

                return;
            }

            _log.LogInformation(
                "{LogCode} completed statusCode={StatusCode} elapsedMs={ElapsedMs} url={Url}",
                logCode,
                (int)response.StatusCode,
                sw.Elapsed.TotalMilliseconds,
                url);
        }

        private string ResolveBaseUrl()
        {
            var configured = (_configuration["DxSpreadsheetWarmup:BaseUrl"] ?? string.Empty).Trim();
            if (!string.IsNullOrWhiteSpace(configured))
                return configured.TrimEnd('/');

            var addresses = _server.Features.Get<IServerAddressesFeature>()?.Addresses;
            if (addresses == null || addresses.Count == 0)
                return string.Empty;

            var normalized = addresses
                .Select(NormalizeLoopbackAddress)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            var http = normalized.FirstOrDefault(x => x.StartsWith("http://", StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrWhiteSpace(http))
                return http;

            return normalized.FirstOrDefault() ?? string.Empty;
        }

        private static string NormalizeLoopbackAddress(string? address)
        {
            var raw = (address ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(raw))
                return string.Empty;

            if (!Uri.TryCreate(raw, UriKind.Absolute, out var uri))
                return raw.TrimEnd('/');

            var host = uri.Host;
            if (string.Equals(host, "*", StringComparison.OrdinalIgnoreCase)
                || string.Equals(host, "+", StringComparison.OrdinalIgnoreCase)
                || string.Equals(host, "0.0.0.0", StringComparison.OrdinalIgnoreCase)
                || string.Equals(host, "::", StringComparison.OrdinalIgnoreCase)
                || string.Equals(host, "[::]", StringComparison.OrdinalIgnoreCase))
            {
                host = "127.0.0.1";
            }

            var builder = new UriBuilder(uri)
            {
                Host = host
            };

            return builder.Uri.GetLeftPart(UriPartial.Authority).TrimEnd('/');
        }

        private int GetIntSetting(string key, int defaultValue)
        {
            var raw = (_configuration[key] ?? string.Empty).Trim();
            return int.TryParse(raw, out var value) && value >= 0 ? value : defaultValue;
        }

        private static string CombineUrl(string baseUrl, string pathAndQuery)
        {
            var left = (baseUrl ?? string.Empty).TrimEnd('/');
            var right = (pathAndQuery ?? string.Empty).TrimStart('/');
            return string.IsNullOrWhiteSpace(left) ? string.Empty : left + "/" + right;
        }
    }
}
