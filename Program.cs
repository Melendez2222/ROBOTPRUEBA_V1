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
using ROBOTPRUEBA_V1.Service;

namespace ROBOTPRUEBA_V1
{
    public class Program
    {
        public static void Main(string[] args)
        {
            
                CreateHostBuilder(args, GlobalSettings.logFilePath).Build().Run();
            
        }

		public static IHostBuilder CreateHostBuilder(string[] args, string logFilePath) =>
	Host.CreateDefaultBuilder(args)
		.UseWindowsService()
		.ConfigureServices((hostContext, services) =>
		{
			services.AddSingleton<Worker>(provider =>
				new Worker(logFilePath, provider.GetRequiredService<IHostApplicationLifetime>()));
			services.AddHostedService<Worker>(provider => provider.GetRequiredService<Worker>());
		});
	}

    public class Worker : BackgroundService
    {
        private readonly ServicioSunat _ServiceSunat;
        private readonly string _logFilePath;
        private readonly CancellationTokenSource _stoppingCts = new CancellationTokenSource();
		private readonly IHostApplicationLifetime _hostApplicationLifetime;
		private readonly Serviceemail serviceEmail;
		public Worker(string logFilePath, IHostApplicationLifetime hostApplicationLifetime)
        {
            _ServiceSunat = new ServicioSunat();
            _logFilePath = logFilePath;
			_hostApplicationLifetime = hostApplicationLifetime;
			serviceEmail = new Serviceemail();
		}

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            try
            {
				
					await _ServiceSunat.Run(stoppingToken);
					var procesos = Process.GetProcessesByName("geckodriver");
					foreach (var proceso in procesos)
					{
						proceso.Kill(); // Termina el proceso
					}
					CreateLogFile("El servicio ha terminado correctamente.");
				
			}
            catch (Exception ex)
            {
				var procesos = Process.GetProcessesByName("geckodriver");
				foreach (var proceso in procesos)
				{
					proceso.Kill(); // Termina el proceso
				}
				CreateLogFile($"Error al correr el servicio: {ex.Message} - StackTrace: {ex.StackTrace}");
				await serviceEmail.SendErrorMessageAsync($"Error al correr el servicio: {ex.Message} - StackTrace: {ex.StackTrace}");

			}
			finally
			{
				// Detiene el host del servicio después de que todo haya terminado.
				_hostApplicationLifetime.StopApplication(); // O utiliza un método más específico para terminar el host.
			}
		}
        public async Task RunAsConsoleApp()
        {
            await ExecuteAsync(new CancellationToken());
        }
        public override async Task StopAsync(CancellationToken cancellationToken)
        {
			var procesos = Process.GetProcessesByName("geckodriver");
			foreach (var proceso in procesos)
			{
				proceso.Kill(); // Termina el proceso
			}
			CreateLogFile("El servicio se está deteniendo.");
            
            _stoppingCts.Cancel();
            await base.StopAsync(cancellationToken);
        }

        private void CreateLogFile(string message)
        {
			if (string.IsNullOrEmpty(GlobalSettings.logFilePath)) { throw new ArgumentNullException(nameof(_logFilePath), "La ruta del archivo de log no puede ser nula o vacía."); }
			string logMessage = $"{DateTime.Now}: {message}{Environment.NewLine}"; 
            File.AppendAllText(GlobalSettings.logFilePath, logMessage);
		}
    }
}