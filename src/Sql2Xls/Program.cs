using CommandLine;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Serilog;
using Sql2Xls.Interfaces;
using Sql2Xls.Sql;
using HelpText = CommandLine.Text.HelpText;

namespace Sql2Xls;

internal class Program
{
    static void Main(string[] args)
    {

        //var parser = Parser.Default;
        var parser = new CommandLine.Parser(with => with.HelpWriter = null);
        var result = parser.ParseArguments<Sql2XlsOptions>(args);

        result
            .WithParsed<Sql2XlsOptions>(options =>
            {
                IConfiguration configuration = new ConfigurationBuilder()
                                               .AddEnvironmentVariables()
                                               .Build();

                Log.Logger = new LoggerConfiguration()
                    .MinimumLevel.Is((Serilog.Events.LogEventLevel)options.LogLevel)
                    .Enrich.WithThreadId()
                    .WriteTo.Console(
                        restrictedToMinimumLevel: (Serilog.Events.LogEventLevel)options.LogLevel,
                        outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] <{ThreadId}> {Message:lj}{NewLine}{Exception}"
                        )
                    .WriteTo.File(
                        options.LogFullPath,
                        //rollingInterval: RollingInterval.Day,
                        restrictedToMinimumLevel: (Serilog.Events.LogEventLevel)options.LogLevel,
                        outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] <{ThreadId}> {Message:lj}{NewLine}{Exception}"
                        )
                    .CreateLogger();


                var serviceCollection = new ServiceCollection();
                ConfigureServices(serviceCollection, configuration);

                var serviceProvider = serviceCollection.BuildServiceProvider();


                var sql2xls = serviceProvider.GetService<ISql2XlsService>();
                sql2xls.Run(options);

                Log.CloseAndFlush();
            })
            .WithNotParsed(errs => DisplayHelp(result, errs));
    }

    static void DisplayHelp<T>(ParserResult<T> result, IEnumerable<Error> errs)
    {
        HelpText helpText = null;
        if (errs.IsVersion())  //check if error is version request
        {
            helpText = HelpText.AutoBuild(result);
        }
        else
        {
            helpText = HelpText.AutoBuild(result, h =>
            {
                h.AdditionalNewLineAfterOption = false;
                //h.Heading = "Sql2Xls 1.0.0";
                //h.Copyright = "Copyright (c) 2023 Sebastian Widz";
                return HelpText.DefaultParsingErrorsHandler(result, h);
            }, e => e);
        }

        Console.WriteLine(helpText);
    }


    private static void ConfigureServices(IServiceCollection services, IConfiguration configuration)
    {
        services.AddLogging(configure =>
        {
            configure.AddSerilog();
        });

        services.AddSingleton(typeof(ISql2XlsService), typeof(Sql2XlsService));
        services.AddSingleton(typeof(ISqlDataService), typeof(SqlDataService));

    }
}
