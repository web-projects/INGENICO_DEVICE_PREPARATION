using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO.Ports;
using System.Linq;
using System.Management;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using DevicePreparation.Helpers;
using DevicePreparation.Interop.DeviceHelper;
using HidLibrary;

namespace DevicePreparation
{
    class DeviceCfg
    {
        [DllImport("MSPorts.dll", SetLastError=true)]
        static extern int ComDBOpen(out IntPtr hComDB);
        [DllImport("MSPorts.dll")]
        public static extern long ComDBClose(UInt32 PHCOMDB);
        [DllImport("MSPorts.dll")]
        public static extern long ComDBReleasePort(UInt32 PHCOMDB,int ComNumber);
        [DllImport("MSPorts.dll")]
        public static extern long ComDBClaimPort(IntPtr hComDB, int  ComNumber, bool ForceClaim, out bool Force);

        // Rescan for hardware changes
        [DllImport("CfgMgr32.dll", SetLastError=true)]
        public static extern int CM_Locate_DevNodeA(ref int pdnDevInst, string pDeviceID, int ulFlags);

        [DllImport("CfgMgr32.dll", SetLastError=true)]
        public static extern int CM_Reenumerate_DevNode(int dnDevInst, int ulFlags);

        public const int CM_LOCATE_DEVNODE_NORMAL = 0x00000000;
        public const int CM_REENUMERATE_NORMAL = 0x00000000;
        public const int CR_SUCCESS = 0x00000000;

        /********************************************************************************************************/
        // ATTRIBUTES
        /********************************************************************************************************/
        #region -- attributes --

        const string INGNAR = "0b00"; //Do NOT make this uppercase
        const string IDTECH = "0acd";
        const string IdTechString = "idtech";

        List<int> comPorts = new List<int>(new int [] { 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 109, 110, 111, 112, 113 });
        int targetPort;

        // states
        bool hascommand;
        bool setdefaults;
        bool setinuse;
        bool helponly;
        bool install280;
        bool install315;
        bool uninstall280;
        bool uninstall315;

        #endregion

        public DeviceCfg(string[] args)
        {
            hascommand = true;
            foreach(var option in args)
            {

                switch(option.ToLower())
                {
                    case "/?":
                    case "/h":
                    case "-h":
                    case "/help":
                    case "-help":
                    { 
                        helponly = true;
                        string fullName = Assembly.GetEntryAssembly().Location;
                        Console.WriteLine($"{System.IO.Path.GetFileNameWithoutExtension(fullName).ToUpper()} [/SETINUSE] [/SETDEFAULTS] | [/INSTALL280] [/INSTALL315] | [/UNINSTALL280] [/UNINSTALL315]");
                        break;
                    }
                    case "/setdefaults":
                    case "-setdefaults":
                    { 
                        setdefaults = true;
                        break;
                    }
                    case "/setinuse":
                    case "-setinuse":
                    { 
                        setinuse = true;
                        break;
                    }
                    case "/install280":
                    case "-install280":
                    { 
                        install280 = true;
                        break;
                    }
                    case "/install315":
                    case "-install315":
                    { 
                        install315 = true;
                        break;
                    }
                    case "/uninstall280":
                    case "-uninstall280":
                    { 
                        uninstall280 = true;
                        break;
                    }
                    case "/uninstall315":
                    case "-uninstall315":
                    { 
                        uninstall315 = true;
                        break;
                    }
                    default:
                    {
                        hascommand = false;
                        break;
                    }
                }
            }
        }

        public bool HasCommand()
        {
            return hascommand;
        }

        public void DeviceInit()
        {
            //DeviceHelper.SetDeviceEnabled(null, null, false);
            DeviceHelper.Test();

            if(!helponly)
            { 
                // Uninstallation: ignore everything else
                if(uninstall280)
                {
                    Utility.UninstallDrivers((int)IngenicoDriverVersions.INGENICO315);
                    return;
                }
                else if(uninstall315)
                {
                    Utility.UninstallDrivers((int)IngenicoDriverVersions.INGENICO315);
                    return;
                }

                // Set TC ports-in-use
                if(setinuse)
                {
                    SetPortsInUse();
                }

                // Set TC default COM Ports
                if(setdefaults)
                {
                    SetDefaultCommPorts();
                }

                string description = string.Empty;
                string deviceID = string.Empty;

                if (FindIngenicoDevice(ref description, ref deviceID))
                {
                    Console.WriteLine("");
                    Console.WriteLine("device identification ----------------------------------------------------------------");
                    Console.WriteLine($"DESCRIPTION                  : {description}");
                    Console.WriteLine($"DEVICE ID                    : {deviceID}");

                    using (var searcher = new ManagementObjectSearcher("SELECT * FROM WIN32_SerialPort"))
                    {
                        try
                        {
                            string[] portNames = SerialPort.GetPortNames();
                            var ports = searcher.Get().Cast<ManagementBaseObject>().ToList();
                            if(ports.Count > 0)
                            { 
                                Console.Write($"PORT NAMES                   : ");
                                foreach(var port in ports)
                                { 
                                    Console.Write($"{port},");
                                }
                                Console.WriteLine("");
                            }
                            var tList = (from n in portNames
                                join p in ports on n equals p["DeviceID"].ToString()
                                select n + " - " + p["Caption"]).ToList();
                            if(tList.Count > 0)
                            { 
                                Console.WriteLine($"DEVICE PORT                  : {tList.FirstOrDefault()}");
                            }
                            else
                            { 
                                Console.WriteLine($"DEVICE PORT                  : NONE FOUND");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"EXCEPTION                    : {ex.Message}");
                        }
                    }

                    string result = ReportCommPorts()?.TrimStart(new char [] { 'C', 'O', 'M' }) ?? "0";
                    targetPort = Convert.ToInt32(result);
                    //if(targetPort > 0)
                    //{
                    //    comPorts.Add(targetPort);
                    //}
                    //ClearUpCommPorts();
                }
                else
                {
                    Console.WriteLine($"device: no Ingenico device found!");
                }

                // IngenicoUSBDrivers installation
                if(install315)
                {
                    Utility.InstallDrivers((int)IngenicoDriverVersions.INGENICO315);
                }
                else if(install280)
                {
                    Utility.InstallDrivers((int)IngenicoDriverVersions.INGENICO280);
                }
            }
        }

        public bool FindIngenicoDevice(ref string description, ref string deviceID)
        {
            List<USBDeviceInfo> devices = GetUSBDevices();
            if (devices.Count == 1)
            {
                BoolStringDuple output = output = new BoolStringDuple(true, devices[0].Description.ToLower().Replace("ingenico ", ""));
                deviceID = devices[0].DeviceID;
                description = devices[0].Description;

                return true;
            }
            return false;
        }

        public static List<USBDeviceInfo> GetUSBDevices()
        {
            List<USBDeviceInfo> devices = new List<USBDeviceInfo>();
            ManagementObjectCollection collection;
            try
            {
                using (var searcher = new ManagementObjectSearcher(@"Select * From Win32_PnPEntity"))
                {
                    collection = searcher.Get();
                }
                foreach (var device in collection)
                {
                    var deviceID = ((string)device.GetPropertyValue("DeviceID") ?? "").ToLower();
                    if (string.IsNullOrWhiteSpace(deviceID))
                    { 
                        continue;
                    }
                    Debug.WriteLine($"device: {deviceID}");
                    if (deviceID.Contains("usb\\") && (deviceID.Contains($"vid_{INGNAR}") || deviceID.Contains($"vid_{IDTECH}")))
                    {
                        devices.Add(new USBDeviceInfo(
                            (string)device.GetPropertyValue("DeviceID"),
                            (string)device.GetPropertyValue("PNPDeviceID"),
                            (deviceID.Contains($"vid_{IDTECH}") ? DeviceCfg.IdTechString : (string)device.GetPropertyValue("Description"))
                        ));
                    }
                }
                collection.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return devices;
        }

        private void SetPortsInUse()
        {
            int state = ComDBOpen(out IntPtr PHCOMDB);
            if(PHCOMDB != null && state == (int)ERROR_STATUS.ERROR_SUCCESS)
            {
                for(int port = 1; port < 250; port++)
                { 
                    long result = ComDBClaimPort(PHCOMDB, port, true, out bool forced);
                    Console.WriteLine($"device: COM{port} forced in-use with status={result}.");
                }
            }
        }

        private void SetDefaultCommPorts()
        {
            try
            {
                // Clear in-use COM Port
                int state = ComDBOpen(out IntPtr PHCOMDB);
                if(PHCOMDB != null && state == (int)ERROR_STATUS.ERROR_SUCCESS)
                { 
                    foreach(var port in comPorts)
                    { 
                        long status = ComDBReleasePort((UInt32)PHCOMDB, port);
                        Console.WriteLine($"device: COM{port} released with status={status}.");
                    }
                    long dsfdf1 = ComDBClose((UInt32)PHCOMDB);
                }

                RescanForHardwareChanges();

                string description = string.Empty;
                string deviceID = string.Empty;
                if (FindIngenicoDevice(ref description, ref deviceID))
                {
                    Console.WriteLine("device capabilities ----------------------------------------------------------------");
                    Console.WriteLine($"DESCRIPTION                  : {description}");
                    Console.WriteLine($"DEVICE ID                    : {deviceID}");

                    using (var searcher = new ManagementObjectSearcher("SELECT * FROM WIN32_SerialPort"))
                    {
                        string[] portNames = SerialPort.GetPortNames();
                        var ports = searcher.Get().Cast<ManagementBaseObject>().ToList();
                        var tList = (from n in portNames
                            join p in ports on n equals p["DeviceID"].ToString()
                            select n + " - " + p["Caption"]).ToList();
                        Console.WriteLine($"DEVICE PORT                  : {tList.FirstOrDefault()}");
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private string ReportCommPorts()
        {
            string port = "0";

            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\WMI", "SELECT * FROM MSSerial_PortName");

                foreach (ManagementObject queryObj in searcher.Get())
                {
                    //If the serial port's instance name contains USB 
                    //it must be a USB to serial device
                    if (queryObj["InstanceName"].ToString().Contains("USB"))
                    {
                        Console.WriteLine("--------------------------------------------------------------------------------------");
                        Console.WriteLine("MSSerial_PortName instance");
                        Console.WriteLine("--------------------------------------------------------------------------------------");
                        Console.WriteLine("InstanceName: {0}", queryObj["InstanceName"]);
                        Console.WriteLine(queryObj["PortName"] + " is a USB to SERIAL adapter/converter");
                        port = queryObj["PortName"].ToString();
                        //SerialPort p = new SerialPort(port);
                        //p.PortName = "COM11";
                        //return port;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"device: ReportCommPorts() exception={ex.Message}");
            }

            return port;
        }

        private void RescanForHardwareChanges()
        {
            int pdnDevInst = 0;

            if (CM_Locate_DevNodeA(ref pdnDevInst, null, CM_LOCATE_DEVNODE_NORMAL) != CR_SUCCESS)
            { 
                Console.WriteLine("device: failed to locate hardware devices");
            }
            if (CM_Reenumerate_DevNode(pdnDevInst, CM_REENUMERATE_NORMAL) != CR_SUCCESS)
            { 
                Console.WriteLine("Failed to reenumerate hardware devices");
            }
            else
            {
                Console.WriteLine("device: reenumerated hardware devices");
            }
        }
    }

    enum ERROR_STATUS
    {
        ERROR_SUCCESS = 0,
        ERROR_ACCESS_DENIED = 5
    };
}
