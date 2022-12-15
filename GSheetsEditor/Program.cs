using GSheetsEditor.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace GSheetsEditor
{
    internal class Program
    {
        public IConfigurationRoot Configuration { get; private set; }

        private IServiceCollection _services = new ServiceCollection();
        private ServiceProvider _serviceProvider;

        static void Main(string[] args) => new Program().MainAsync(args).GetAwaiter().GetResult();

        public async Task MainAsync(string[] args)
        {
            var builder = new ConfigurationBuilder().SetBasePath(AppContext.BaseDirectory).AddJsonFile("config.json");
            Configuration = builder.Build();
            _services.AddSingleton<TelegramBotCommandService>().AddSingleton(Configuration);
            _serviceProvider = _services.BuildServiceProvider();
            _serviceProvider.GetRequiredService<TelegramBotCommandService>();

            await Task.Delay(-1);
        }

        public void ConfigureServices()
        {
           
        }
    }
}