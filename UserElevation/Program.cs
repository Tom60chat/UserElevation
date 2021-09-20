using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;

namespace UserElevation
{
    class Program
    {
        static void Main(string[] args)
        {
            TaskSchedulerClient taskScheduler = new TaskSchedulerClient();

            if (taskScheduler.TryGetTask(out var task))
            {

                if (IsAdministrator())
                {
                    Process.Start("cmd.exe");
                }
                else
                {
                    //taskScheduler.CheckAndRepairTask(ref task);
                    task.Run(0);
                }
            }
            else
            {
                if (IsAdministrator())
                    taskScheduler.CreatTask();
                else
                {
                    Console.WriteLine("You need to execute this app one time in admin");

                    Console.WriteLine("Press any key to exit");
                    Console.ReadLine();
                }
            }
        }

        // https://stackoverflow.com/questions/3600322/check-if-the-current-user-is-administrator
        public static bool IsAdministrator()
        {
            using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
            {
                WindowsPrincipal principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
        }
    }
}
