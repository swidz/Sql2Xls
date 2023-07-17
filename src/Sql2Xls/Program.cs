using CommandLine;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Serilog;

namespace Sql2Xls;

internal class Program
{
    static void Main(string[] args)
    {
      
        
        var result = Parser.Default.ParseArguments<Sql2XlsOptions>(args);
        result.WithParsed<Sql2XlsOptions>(context =>
        {
            IConfiguration configuration = new ConfigurationBuilder()
                                           .AddEnvironmentVariables()
                                           .Build();

            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Is((Serilog.Events.LogEventLevel)context.LogLevel)
                .WriteTo.Console(
                    restrictedToMinimumLevel: (Serilog.Events.LogEventLevel)context.LogLevel,
                    outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] <{ThreadId}> {Message:lj}{NewLine}{Exception}"
                    )
                .WriteTo.File(
                    context.LogFullPath,
                    //rollingInterval: RollingInterval.Day,
                    restrictedToMinimumLevel: (Serilog.Events.LogEventLevel)context.LogLevel,
                    outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] <{ThreadId}> {Message:lj}{NewLine}{Exception}"
                    )
                .CreateLogger();

            var serviceCollection = new ServiceCollection();
            ConfigureServices(serviceCollection, configuration);

            var serviceProvider = serviceCollection.BuildServiceProvider();


            var sql2xls = serviceProvider.GetService<Sql2XlsService>();
            sql2xls.Run();

            Log.CloseAndFlush();
        });

    }

    private static void ConfigureServices(IServiceCollection services, IConfiguration configuration)
    {
        services.AddLogging(configure =>
        {
            configure.AddSerilog();
        });


        //TODO

    }
}
