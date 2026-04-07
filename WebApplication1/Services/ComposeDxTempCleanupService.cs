using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace WebApplication1.Services
{
    /// <summary>
    /// App_Data/DocDxCompose 폴더의 임시 파일을 7일 주기로 정리하는 백그라운드 서비스
    /// </summary>
    public class ComposeDxTempCleanupService : BackgroundService
    {
        private readonly IWebHostEnvironment _env;
        private readonly ILogger<ComposeDxTempCleanupService> _log;

        // 7일보다 오래된 파일 삭제
        private static readonly TimeSpan FileMaxAge = TimeSpan.FromDays(7);
        // 24시간마다 실행
        private static readonly TimeSpan CheckInterval = TimeSpan.FromHours(24);

        public ComposeDxTempCleanupService(
            IWebHostEnvironment env,
            ILogger<ComposeDxTempCleanupService> log)
        {
            _env = env;
            _log = log;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            _log.LogInformation("ComposeDxTempCleanupService started.");

            // 앱 시작 직후 한 번 실행
            await CleanupAsync(stoppingToken);

            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    await Task.Delay(CheckInterval, stoppingToken);
                    await CleanupAsync(stoppingToken);
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    _log.LogError(ex, "ComposeDxTempCleanupService 실행 중 오류 발생");
                }
            }

            _log.LogInformation("ComposeDxTempCleanupService stopped.");
        }

        private Task CleanupAsync(CancellationToken cancellationToken)
        {
            var dir = Path.Combine(_env.ContentRootPath, "App_Data", "DocDxCompose");

            if (!Directory.Exists(dir))
            {
                _log.LogDebug("ComposeDxTempCleanup: 폴더 없음, 스킵 ({dir})", dir);
                return Task.CompletedTask;
            }

            var cutoff = DateTime.Now - FileMaxAge;
            var deleted = 0;
            var failed = 0;
            var files = Directory.GetFiles(dir, "*.xlsx");

            foreach (var file in files)
            {
                if (cancellationToken.IsCancellationRequested) break;

                try
                {
                    var created = File.GetCreationTime(file);
                    if (created < cutoff)
                    {
                        File.Delete(file);
                        deleted++;
                        _log.LogDebug("ComposeDxTempCleanup: 삭제 {file}", file);
                    }
                }
                catch (Exception ex)
                {
                    failed++;
                    _log.LogWarning(ex, "ComposeDxTempCleanup: 삭제 실패 {file}", file);
                }
            }

            if (deleted > 0 || failed > 0)
                _log.LogInformation(
                    "ComposeDxTempCleanup 완료 — 삭제: {deleted}, 실패: {failed}, 대상폴더: {dir}",
                    deleted, failed, dir);

            return Task.CompletedTask;
        }
    }
}
