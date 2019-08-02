using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DevicePreparation
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {

                DeviceCfg device = new DeviceCfg(args);
                if(device != null)
                {
                    if(device.HasCommand())
                    { 
                        bool resetport = device.DeviceInit();
                        if(resetport)
                        {
                            Console.WriteLine($"\r\nmain  : reset port requested.");
                            Task responseTask = Task.Run(() => 
                            { 
                                bool result = device.ResetPort();
                                Console.WriteLine($"main  : reset port completed with result={result}");
                            });
                            Task newTask = responseTask.ContinueWith(t => Console.WriteLine("main  : waiting for port reset to complete..."));
                            newTask.Wait();
                        }
                    }
                    else
                    {
                        Console.WriteLine("NO MATCHING COMMAND TO RUN: USE /HELP FOR COMMANDS.");
                    }

                    if(Debugger.IsAttached)
                    { 
                        Console.Write("\r\nPRESS ANY KEY TO EXIT...");
                        Console.ReadKey();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"main: exception={ex.Message}");
            }
        }
    }
}
