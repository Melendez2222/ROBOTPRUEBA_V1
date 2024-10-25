using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Toolkit.Uwp.Notifications;
using ROBOTPRUEBA_V1.CONFIG;

namespace ROBOTPRUEBA_V1
{
    public class Program
    {
        public static void Main(string[] args)
        {
            
            //CreateHostBuilder(args, logFilePath).Build().Run();
            if (Environment.UserInteractive)
            {
               
                var worker = new Worker(GlobalSettings.logFilePath);
                worker.RunAsConsoleApp().GetAwaiter().GetResult();
            }
            else
            {
                
                CreateHostBuilder(args, GlobalSettings.logFilePath).Build().Run();
            }
        }

        public static IHostBuilder CreateHostBuilder(string[] args, string logFilePath) =>
            Host.CreateDefaultBuilder(args)
                .UseWindowsService()
                .ConfigureServices((hostContext, services) =>
                {
                    services.AddSingleton<Worker>(provider => new Worker(logFilePath));
                    services.AddHostedService<Worker>(provider => provider.GetRequiredService<Worker>());
                });
    }

    public class Worker : BackgroundService
    {
        private readonly ServicioSunat _ServiceSunat;
        private readonly string _logFilePath;
        private readonly CancellationTokenSource _stoppingCts = new CancellationTokenSource();

        public Worker(string logFilePath)
        {
            _ServiceSunat = new ServicioSunat();
            _logFilePath = logFilePath;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            try
            {
                await _ServiceSunat.Run(stoppingToken);
                CreateLogFile("El servicio ha terminado correctamente.");
                _stoppingCts.Cancel();
            }
            catch (Exception ex)
            {
                CreateLogFile($"Error: {ex.Message}");
            }
        }
        public async Task RunAsConsoleApp()
        {
            await ExecuteAsync(new CancellationToken());
        }
        public override async Task StopAsync(CancellationToken cancellationToken)
        {
            CreateLogFile("El servicio se está deteniendo.");
            _stoppingCts.Cancel();
            await base.StopAsync(cancellationToken);
        }

        private void CreateLogFile(string message)
        {
            string logMessage = $"{DateTime.Now}: {message}{Environment.NewLine}";
            File.AppendAllText(_logFilePath, logMessage);
        }
    }
}