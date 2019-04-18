using Microsoft.Extensions.Configuration;
using System;
using System.IO;
using System.Linq;

namespace JnJ.EAS.TicketTracker.Core
{    
    public static class ConfigurationHelper
    {
        private static IConfigurationRoot configuration;

        public static void LoadConfig()
        {
            var builder = new ConfigurationBuilder()
                        .SetBasePath(Directory.GetCurrentDirectory())
                        .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);

            configuration = builder.Build();
        }

        public static string GetAppSettingValue(string key)
        {
            return configuration.GetSection("AppSettings").GetChildren()
                                .Where(x => x.Key == key)
                                .Select(x => x.Value)
                                .FirstOrDefault();
        }
    }
}
