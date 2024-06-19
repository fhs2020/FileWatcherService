using System.ServiceProcess;

namespace FileWatcherService
{
    public static class Program
    {
        static void Main()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new FileWatcher()
            };
            ServiceBase.Run(ServicesToRun);
        }
    }
}
