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

                DeviceCfg config = new DeviceCfg(args);
                if(config != null)
                {
                    if(config.HasCommand())
                    { 
                        config.DeviceInit();
                    }
                    else
                    {
                        Console.WriteLine("NO MATCHING COMMAND TO RUN: USE /HELP FOR COMMANDS.");
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
