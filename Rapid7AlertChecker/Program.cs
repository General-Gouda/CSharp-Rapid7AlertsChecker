using System;
using Topshelf;
using NLog;

namespace Rapid7AlertChecker
{
    class Program
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        static void Main()
        {
            HostFactory.Run(x => {
                x.UseNLog();
                x.StartAutomaticallyDelayed();
                x.RunAsPrompt();
                x.SetServiceName("Rapid7Alerts");
                x.SetDisplayName("Rapid7 Alerts to Slack Service");

                x.Service<R7AlertsService>(svcHost =>
                {
                    svcHost.ConstructUsing(svc => new R7AlertsService());
                    svcHost.WhenStarted(svc =>
                    {
                        Logger.Info("Service start...");
                        svc.Start();
                        Logger.Info("Service start complete...");
                    });
                    svcHost.WhenStopped((svc, hostControl) =>
                    {
                        Logger.Info("Service stopping...");
                        hostControl.RequestAdditionalTime(TimeSpan.FromSeconds(60));
                        Logger.Info("Service stopping...");
                        return svc.Stop();
                    });
                    svcHost.WhenShutdown((svc, hostControl) =>
                    {
                        Logger.Info("Service shutdown...");
                        hostControl.RequestAdditionalTime(TimeSpan.FromSeconds(60));
                        svc.Stop();
                        Logger.Info("Service shutdown complete.");
                    });
                });
            });
        }
    }
}
