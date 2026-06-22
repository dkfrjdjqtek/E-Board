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
    /// App_Data DocDXCompose 및 DocDXStamp 폴더의 임시 파일을 7일 주기로 정리하는 백그라운드 서비스
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

        // 2026.06.22 Changed: 정리 메서드 이름을 DocDXCompose와 DocDXStamp 기준으로 통일 Contents 폴더명과 함수명 규칙을 동일하게 맞춤
        private Task CleanupAsync(CancellationToken cancellationToken)
        {
            CleanupDocDXComposeFiles(cancellationToken);
            CleanupDocDXStampFiles(cancellationToken);

            return Task.CompletedTask;
        }

        // 2026.06.22 Added: DocDXCompose 임시 파일을 7일 기준으로 정리 Contents 기존 App_Data DocDxCompose 폴더 정리 동작은 유지
        private void CleanupDocDXComposeFiles(CancellationToken cancellationToken)
        {
            var dir = Path.Combine(_env.ContentRootPath, "App_Data", "DocDxCompose");

            if (!Directory.Exists(dir))
            {
                _log.LogDebug("DocDXComposeCleanup: 폴더 없음, 스킵 ({dir})", dir);
                return;
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
                        _log.LogDebug("DocDXComposeCleanup: 삭제 {file}", file);
                    }
                }
                catch (Exception ex)
                {
                    failed++;
                    _log.LogWarning(ex, "DocDXComposeCleanup: 삭제 실패 {file}", file);
                }
            }

            if (deleted > 0 || failed > 0)
                _log.LogInformation(
                    "DocDXComposeCleanup 완료 삭제 {deleted} 실패 {failed} 대상폴더 {dir}",
                    deleted, failed, dir);
        }

        // 2026.06.22 Added: DocDXStamp 작업 파일을 7일 기준으로 정리 Contents App_Data DocDXStamp 하위 xlsx 파일만 대상으로 한다
        private void CleanupDocDXStampFiles(CancellationToken cancellationToken)
        {
            var dir = Path.Combine(_env.ContentRootPath, "App_Data", "DocDXStamp");

            if (!Directory.Exists(dir))
            {
                _log.LogDebug("DocDXStampCleanup: 폴더 없음, 스킵 ({dir})", dir);
                return;
            }

            var cutoff = DateTime.Now - FileMaxAge;
            var deleted = 0;
            var failed = 0;
            string[] files;

            try
            {
                files = Directory.GetFiles(dir, "*.xlsx", SearchOption.AllDirectories);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "DocDXStampCleanup: 파일 검색 실패 {dir}", dir);
                return;
            }

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
                        _log.LogDebug("DocDXStampCleanup: 삭제 {file}", file);
                    }
                }
                catch (Exception ex)
                {
                    failed++;
                    _log.LogWarning(ex, "DocDXStampCleanup: 삭제 실패 {file}", file);
                }
            }

            if (deleted > 0 || failed > 0)
                _log.LogInformation(
                    "DocDXStampCleanup 완료 삭제 {deleted} 실패 {failed} 대상폴더 {dir}",
                    deleted, failed, dir);
        }
    }
}