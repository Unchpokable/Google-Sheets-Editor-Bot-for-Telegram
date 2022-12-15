using Microsoft.Extensions.Configuration;


namespace GSheetsEditor.Services
{
    internal class ServiceBase
    {
        protected readonly IServiceProvider _serviceProvider;
        protected readonly IConfigurationRoot _configuration;

        public ServiceBase(IServiceProvider serviceProvider, IConfigurationRoot configuration)
        {
            _serviceProvider = serviceProvider;
            _configuration = configuration;
        }
    }
}
