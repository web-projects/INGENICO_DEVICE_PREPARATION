using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DevicePreparation.Helpers
{
    public static class Utility
    {
        private static StringBuilder stdOutput;
        private static StringBuilder stdError;
        private static Process process;

        private const string IngenicoUSBDriver280 = "IngenicoUSBDrivers_2.80_setup";
        private const string IngenicoUSBDriver315 = "IngenicoUSBDrivers_3.15_setup";
        private const string IngenicoUSBPorts280 = "/PORT='109' /PID='0062:109' /PORT='110' /PID='0060:110' /PORT='111' /PID='0061:111'";
        private const string IngenicoUSBPorts315 = "/PORT='35' /PID='0062:35' /PORT='35' /PID='0060:35' /PORT='35' /PID='0061:35'";
        private const string IngencioUSBDestination = "'C:\\Program Files\\Ingenico\\IngenicoUSBDrivers'";

        public static string UninstallDrivers(int version)
        {
            string result = string.Empty;
            switch(version)
            {
                case (int)IngenicoDriverVersions.INGENICO280:
                {
                    string executable = $"./Assets/{IngenicoUSBDriver280}.exe";
                    if (File.Exists(executable))
                    {
                        var fInfo = new FileInfo(executable);
                        executable = fInfo.FullName;
                        result = RunExternalExeElevated(fInfo.DirectoryName, executable, "/UNINSTALL");
                    }
                    break;
                }
                case (int)IngenicoDriverVersions.INGENICO315:
                {
                    string executable = $"./Assets/{IngenicoUSBDriver315}.exe";
                    if (File.Exists(executable))
                    {
                        var fInfo = new FileInfo(executable);
                        executable = fInfo.FullName;
                        result = RunExternalExeElevated(fInfo.DirectoryName, executable, "/UNINSTALL");
                    }
                    break;
                }
            }

            return result;
        }

        public static string InstallDrivers(int version)
        {
            string result = string.Empty;
            switch(version)
            {
                case (int)IngenicoDriverVersions.INGENICO280:
                {
                    string IngencioUSBArguments = $"/S PORTS={IngenicoUSBPorts280} /INST_PATH={IngencioUSBDestination}";
                    string executable = $"./Assets/{IngenicoUSBDriver280}.exe";
                    if (File.Exists(executable))
                    {
                        var fInfo = new FileInfo(executable);
                        executable = fInfo.FullName;
                        result = RunExternalExeElevated(fInfo.DirectoryName, executable, IngencioUSBArguments);
                    }
                    break;
                }
                case (int)IngenicoDriverVersions.INGENICO315:
                {
                    string IngencioUSBArguments = $"/S PORTS={IngenicoUSBPorts315} /INST_PATH={IngencioUSBDestination}";
                    string executable = $"./Assets/{IngenicoUSBDriver315}.exe";
                    if (File.Exists(executable))
                    {
                        var fInfo = new FileInfo(executable);
                        executable = fInfo.FullName;
                        result = RunExternalExeElevated(fInfo.DirectoryName, executable, IngencioUSBArguments);
                    }
                    break;
                }
            }
            return result;
        }

        public static string RunExternalExeElevated(string directory, string filename, string arguments = null, string[] environmentVariables = null)
        {
            Console.WriteLine($"{directory}--{filename} {arguments ?? ""}");

            var output = Path.GetTempFileName();
            var process = Process.Start(new ProcessStartInfo
            {
                FileName  = filename,
                Arguments = arguments,
                Verb      = "runas",
                UseShellExecute = true
            });
            process.WaitForExit();

            string response = File.ReadAllText(output);
            File.Delete(output);

            if (process.ExitCode == 0)
            {
                Console.WriteLine("Process exited normally.");
                return response;
            }
            else if (process.ExitCode < 0)
            {
                Console.WriteLine($"Process terminated abnormally {process.StartInfo.FileName}.");
                return stdOutput.ToString();  //Exit cleanly, don't throw, which breaks current execution pattern.
            }
            else
            {
                var message = new StringBuilder(256);
                if (stdOutput.Length != 0)
                {
                    message.AppendLine("Std output:");
                    message.AppendLine(stdOutput.ToString());
                }

                if (!string.IsNullOrEmpty(stdError.ToString()))
                {
                    message.AppendLine(stdError.ToString());
                }

                throw new Exception($"Exception: finished with exit code = {process.ExitCode}: {message}");
            }
        }
    }

    public enum IngenicoDriverVersions
    {
        INGENICO280,
        INGENICO315
    }
}
