using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO.Ports;
using System.Linq;
using System.ComponentModel;
using System.Management;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using DevicePreparation.Helpers;
using DevicePreparation.Interop.DeviceHelper;
using HidLibrary;
using Microsoft.Win32;
using System.Text.RegularExpressions;

namespace DevicePreparation
{
    class DeviceCfg
    {
        [DllImport("MSPorts.dll", SetLastError=true)]
        static extern int ComDBOpen(out IntPtr hComDB);
        [DllImport("MSPorts.dll")]
        public static extern long ComDBClose(UInt32 PHCOMDB);
        [DllImport("MSPorts.dll", SetLastError=true)]
        public static extern long ComDBGetCurrentPortUsage(IntPtr PHComDB, [In, Out] byte[] Buffer, Int32 BufferSize, Int32 ReportType, out Int32 MaxPortsReported);
        [DllImport("MSPorts.dll", SetLastError=true)]
        public static extern long ComDBReleasePort(UInt32 PHCOMDB,int ComNumber);
        [DllImport("MSPorts.dll", SetLastError=true)]
        public static extern long ComDBClaimPort(IntPtr hComDB, int  ComNumber, bool ForceClaim, out bool Force);

        // Rescan for hardware changes
        [DllImport("CfgMgr32.dll", SetLastError=true)]
        public static extern int CM_Locate_DevNodeA(ref int pdnDevInst, string pDeviceID, int ulFlags);

        [DllImport("CfgMgr32.dll", SetLastError=true)]
        public static extern int CM_Reenumerate_DevNode(int dnDevInst, int ulFlags);

        public const int CM_LOCATE_DEVNODE_NORMAL = 0x00000000;
        public const int CM_REENUMERATE_NORMAL = 0x00000000;
        public const int CR_SUCCESS = 0x00000000;

         [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int RegisterWindowMessage(string lpString);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern bool SendNotifyMessage(int hWnd, int Msg, int wParam, int lParam);

        private const int HWND_BROADCAST = 0xffff;
        private const int WM_WININICHANGE = 0x001a;
        private const int WM_SETTINGCHANGE = 0x001a;

        /********************************************************************************************************/
        // ATTRIBUTES
        /********************************************************************************************************/
        #region -- attributes --

        const string INGNAR = "0b00"; //Do NOT make this uppercase
        const string INGNSR = "079b";
        const string IDTECH = "0acd";
        const string IdTechString = "idtech";

        List<int> comPorts = new List<int>(new int [] { 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 109, 110, 111, 112, 113 });
        int targetPort;
        int devicePort;
        string instanceId;

        // states
        bool hascommand;
        bool findingenico;
        bool setdefaults;
        bool getinfo;
        bool setinuse;
        bool setport;
        bool resetport;
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
                        Console.WriteLine($"{System.IO.Path.GetFileNameWithoutExtension(fullName).ToUpper()} [/INFO] [/SETINUSE] [/SETDEFAULTS] [/RESETPORT]| [/INSTALL280] [/INSTALL315] | [/UNINSTALL280] [/UNINSTALL315]");
                        break;
                    }
                    case "/info":
                    case "-info":
                    { 
                        getinfo = true;
                        break;
                    }
                    case "/findingenico":
                    case "-findingenico":
                    { 
                        findingenico = true;
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
                    case "/setport":
                    case "-setport":
                    { 
                        setport = true;
                        break;
                    }
                    case "/resetport":
                    case "-resetport":
                    { 
                        resetport = true;
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

        public bool DeviceInit()
        {
            //DeviceHelper.SetDeviceEnabled(null, null, false);
            DeviceHelper.Test();

            if(!helponly)
            { 
                // Uninstallation: ignore everything else
                if(uninstall280)
                {
                    Utility.UninstallDrivers((int)IngenicoDriverVersions.INGENICO315);
                    return false;
                }
                else if(uninstall315)
                {
                    Utility.UninstallDrivers((int)IngenicoDriverVersions.INGENICO315);
                    return false;
                }

                // Disable USB Suspend
                if(!getinfo)
                {
                    SelectiveUSBSuspendDisable();
                }

                // Set TC ports-in-use
                if(setinuse)
                {
                    SetPortsInUse();
                }

                // Set TC default COM Ports
                if(setdefaults)
                {
                    string instanceId = string.Empty;
                    List<string> result = ReportUSBCommPorts(ref instanceId);
                    foreach(var port in result)
                    { 
                        int targetPort = Convert.ToInt32(port?.TrimStart(new char [] { 'C', 'O', 'M' }) ?? "0");
                        if (targetPort > 0)
                        {
                            if (comPorts.IndexOf(targetPort) != -1)
                            { 
                                comPorts.Remove(targetPort);
                            }
                        }
                    }

                    devicePort = SetDefaultCommPorts();
                }

                string description = string.Empty;
                string deviceID = string.Empty;

                if (FindIngenicoDevice(ref description, ref deviceID))
                {
                    Console.WriteLine("");
                    Console.WriteLine("device identification ----------------------------------------------------------------");
                    Console.WriteLine($"DESCRIPTION                  : {description}");
                    Console.WriteLine($"DEVICE ID                    : {deviceID}");

                    string compare1 = "";
                    Match usbPattern = Regex.Match(deviceID,  @"^USB\\", RegexOptions.IgnoreCase);
                    if(usbPattern.Success)
                    {
                        compare1 = Regex.Replace(deviceID, @"^USB\\", "", RegexOptions.IgnoreCase);
                    }
                    else
                    {
                        usbPattern = Regex.Match(deviceID,  @"^USBVCOM\\", RegexOptions.IgnoreCase);
                        if(usbPattern.Success)
                        {
                            compare1 = Regex.Replace(deviceID, @"^USBVCOM\\", "", RegexOptions.IgnoreCase);
                        }
                    }
                    string compare2 = "";
                    if (!string.IsNullOrEmpty(compare1))
                    {
                        compare2 = compare1.Replace('\\', '#');
                        string usbcommport = GetDeviceSerialComm(compare1, compare2, "INST_0");
                        Console.WriteLine($"DEVICE PORT                  : {usbcommport}");
                        int targetPort = Convert.ToInt32(usbcommport?.TrimStart(new char[] { 'C', 'O', 'M' }) ?? "0");
                        if (targetPort > 0)
                        {
                            if (targetPort == 35)
                            {
                                Console.WriteLine($"\r\nIngenico device currently installed on COM{targetPort}\r\n");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"\r\nDevice  ID '{deviceID}' is currently installed on '{usbcommport}'\r\n");
                        }
                    }

                    if(setport)
                    {
                        SetDeviceSerialComm(compare1, compare2, "COM35");
                    }

                    List<string> usbCommPorts = ReportUSBCommPorts(ref instanceId);

                    Console.WriteLine("");
                    List<string> result = ReportSerialCommPorts();

                    foreach(var port in result)
                    { 
                        targetPort = Convert.ToInt32(port?.TrimStart(new char [] { 'C', 'O', 'M' }) ?? "0");
                        if(targetPort > 0)
                        {
                            Console.WriteLine($"SERIAL DEVICE FOUND ON         : COM{targetPort}");
                        //    comPorts.Add(targetPort);
                        }
                        //ClearUpCommPorts();
                    }
                }
                else
                {
                    Console.WriteLine($"device: no Ingenico device found!");
                    resetport = false;
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

            return resetport ? (!string.IsNullOrEmpty(instanceId)) : false;
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

        string GetUSBComPorts(string usbDeviceName)
        {
            string port = string.Empty;
            using(var searcher = new ManagementObjectSearcher(@"Select * From Win32_USBHub"))
            { 
                foreach (var device in searcher.Get())
                {                
                    string deviceId = device["DeviceID"].ToString();
                    //port = device["Caption"].ToString();
                    port = device["Name"].ToString();
                    if (deviceId == usbDeviceName)
                    { 
                        Console.WriteLine($"Port for \"{usbDeviceName}\" is \"{port}\"");
                        Debug.WriteLine($"Port for \"{usbDeviceName}\" is \"{port}\"");
                        break;
                    }
                }
            }
            return port;
        }

        string GetDeviceSerialComm(string compare1, string compare2, string compare3)
        {
            string port = null;
            try
            {
                using (var hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
                using (var rk = hklm.OpenSubKey(@"HARDWARE\DEVICEMAP\SERIALCOMM"))
                {
                    foreach (var skName in rk.GetValueNames())
                    {
                        int offset = skName.IndexOf("VID_0B00");
                        if(offset == -1)
                        { 
                            offset = skName.IndexOf("VID_079B");
                        }
                        if(offset != -1)
                        { 
                            if(skName.IndexOf(compare1, offset, StringComparison.InvariantCultureIgnoreCase) != -1 ||
                               skName.IndexOf(compare2, offset, StringComparison.InvariantCultureIgnoreCase) != -1 ||
                               skName.IndexOf(compare3, offset, StringComparison.InvariantCultureIgnoreCase) != -1)
                            {
                                port =  rk.GetValue(skName).ToString();
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"SearchIngenicoDriversInstalled - Execution Error: {ex.Message}");
            }

            return port;
        }

        void SetTerminalComPort()
        {
            //Computer\HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Enum\usbVcom\VID_0B00&PID_0061\80242705\Device Parameters
            /*try
            {
                using (var hklm = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64))
                using (var rk = hklm.OpenSubKey(@"Software\Ingenico\Ingenico Form Builder\Terminal\Serial", true))
                {
                    foreach (var skName in rk.GetValueNames())
                    {
                        int offset = skName.IndexOf("VID_0B00");
                        if(offset != -1)
                        { 
                            if(skName.IndexOf(compare1, offset, StringComparison.InvariantCultureIgnoreCase) != -1 ||
                               skName.IndexOf(compare2, offset, StringComparison.InvariantCultureIgnoreCase) != -1)
                            {
                                string port =  rk.GetValue(skName).ToString();
                                if(!string.Equals(port, comPort, StringComparison.InvariantCultureIgnoreCase))
                                {
                                    rk.SetValue(skName, comPort);
                                    SendNotifyMessage(HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0);
                                    Task responseTask = Task.Run(() => 
                                    { 
                                        RescanForHardwareChanges();
                                        Console.WriteLine($"device: hardware scan complete.");
                                    });
                                    Task newTask = responseTask.ContinueWith(t => Console.WriteLine("device: waiting for hardware scan to complete..."));
                                    newTask.Wait();
                                }
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"SearchIngenicoDriversInstalled - Execution Error: {ex.Message}");
            }*/
        }
        string SetDeviceSerialComm(string compare1, string compare2, string comPort)
        {
            string port = string.Empty;
            try
            {
                using (var hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
                using (var rk = hklm.OpenSubKey(@"HARDWARE\DEVICEMAP\SERIALCOMM", true))
                {
                    foreach (var skName in rk.GetValueNames())
                    {
                        int offset = skName.IndexOf("VID_0B00");
                        if(offset != -1)
                        { 
                            if(skName.IndexOf(compare1, offset, StringComparison.InvariantCultureIgnoreCase) != -1 ||
                               skName.IndexOf(compare2, offset, StringComparison.InvariantCultureIgnoreCase) != -1)
                            {
                                port =  rk.GetValue(skName).ToString();
                                if(!string.Equals(port, comPort, StringComparison.InvariantCultureIgnoreCase))
                                {
                                    rk.SetValue(skName, comPort);
                                    SendNotifyMessage(HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0);
                                    Task responseTask = Task.Run(() => 
                                    { 
                                        RescanForHardwareChanges();
                                        Console.WriteLine($"device: hardware scan complete.");
                                    });
                                    Task newTask = responseTask.ContinueWith(t => Console.WriteLine("device: waiting for hardware scan to complete..."));
                                    newTask.Wait();
                                }
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"SearchIngenicoDriversInstalled - Execution Error: {ex.Message}");
            }

            return port;
        }

        public static List<USBDeviceInfo> GetUSBDevices()
        {
            List<USBDeviceInfo> devices = new List<USBDeviceInfo>();
            try
            {
                using (var searcher = new ManagementObjectSearcher(@"Select * From Win32_PnPEntity"))
                {
                    foreach (var device in searcher.Get())
                    {
                        var deviceID = ((string)device.GetPropertyValue("DeviceID") ?? "").ToLower();
                        if (string.IsNullOrWhiteSpace(deviceID))
                        { 
                            continue;
                        }
                        Debug.WriteLine($"device: {deviceID}");
                        if (deviceID.Contains("usb\\") && (deviceID.Contains($"vid_{INGNAR}") || deviceID.Contains($"vid_{INGNSR}") || deviceID.Contains($"vid_{IDTECH}")))
                        {
                            devices.Add(new USBDeviceInfo(
                                (string)device.GetPropertyValue("DeviceID"),
                                (string)device.GetPropertyValue("PNPDeviceID"),
                                (deviceID.Contains($"vid_{IDTECH}") ? DeviceCfg.IdTechString : (string)device.GetPropertyValue("Description"))
                            ));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return devices;
        }

        public bool ResetPort()
        {
            Console.WriteLine($"device: reset device with instanceId: {instanceId}");
            bool result = PortHelper.TryResetPortByInstanceId(instanceId);

            // Search for new devices attaching
            RescanForHardwareChanges();

            return result;
        }

        private void SetPortsInUse()
        {
            int state = ComDBOpen(out IntPtr PHCOMDB);
            if(PHCOMDB != null)
            {
                if(state == (int)ERROR_STATUS.ERROR_SUCCESS)
                { 
                    for(int port = 1; port < 250; port++)
                    { 
                        long result = ComDBClaimPort(PHCOMDB, port, true, out bool forced);
                        result &= 0x00000000000FFFFF;
                        Console.WriteLine($"device: COM{port} forced in-use with status={result}.");
                    }
                }
                else
                {
                    Console.WriteLine($"ClearUpCommPorts: failed with state={state}");
                }
            }
            else
            {
                Console.WriteLine($"ClearUpCommPorts: failed with state={state}");
            }
        }

        private byte [] GetCurrentCommPorts(IntPtr PHComDB)
        {
            byte [] buffer = null;
            try
            {
                int bufferSize = 0;
                Int32 maxportsReported = 0;
                long status = ComDBGetCurrentPortUsage(PHComDB, buffer, bufferSize, (Int32)REPORT_BYTES.CDB_REPORT_BYTES, out maxportsReported);
                status &= 0x00000000000FFFFF;
                if(status == 0 && Marshal.GetLastWin32Error() == 0)
                { 
                    if(maxportsReported > 0)
                    { 
                        bufferSize = maxportsReported;
                        //IntPtr buffer = new IntPtr(bufferSize);
                        //byte [] buffer = new byte[bufferSize];
                        //IntPtr ptr = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(byte)) * buffer.Length);
                        //Marshal.Copy(buffer, 0, ptr, buffer.Length);
                        //Marshal.FreeHGlobal(ptr);
                        //IntPtr ptr = Marshal.AllocCoTaskMem(bufferSize);
                        //status = ComDBGetCurrentPortUsage(PHComDB, ref ptr, bufferSize, (Int32)REPORT_BYTES.CDB_REPORT_BITS, ref maxportsReported);
                        buffer = new byte[bufferSize];
                        status = ComDBGetCurrentPortUsage(PHComDB, buffer, bufferSize, (Int32)REPORT_BYTES.CDB_REPORT_BYTES, out maxportsReported);
                        status &= 0x00000000000FFFFF;
                        if(status == 0 && Marshal.GetLastWin32Error() == 0)
                        { 
                            Console.WriteLine($"device: GetCurrentCommPorts() buffer size={maxportsReported}.");
                        }
                    }
                    else
                    {
                        buffer = new byte[1];
                    }
                }
                else
                {
                    //ERROR_ADAP_HDW_ERR = 57L
                    string value = string.Format("{0:X}", status);
                    Console.WriteLine($"device: GetCurrentCommPorts() failed with status=0x{value}.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"EXCEPTION IN GetCurrentCommPorts(): {ex.Message}");
            }

            return buffer;
        }

        private int SetDefaultCommPorts()
        {
            int ingenicoPort = 0;
            try
            {
                // Clear in-use COM Port
                int state = ComDBOpen(out IntPtr PHCOMDB);
                if(PHCOMDB != null && state == (int)ERROR_STATUS.ERROR_SUCCESS)
                {
                    byte [] buffer = GetCurrentCommPorts(PHCOMDB);
                    string comPortList = string.Empty;
                    foreach(var val in comPorts)
                    {
                        if(buffer?.Length > 0)
                        { 
                            if(buffer.Length >= val -1)
                            { 
                                if(buffer[val - 1] == 1)
                                { 
                                    comPortList += $"{val},";
                                }
                            }
                        }
                        else
                        { 
                            comPortList += $"{val},";
                        }
                    }
                    comPortList = comPortList.TrimEnd(',');
                    if(comPortList.Length > 0)
                    { 
                        Console.WriteLine($"device: releasing COM ports=[{comPortList}]");
                        foreach(var port in comPorts)
                        { 
                            long status = ComDBReleasePort((UInt32)PHCOMDB, port);
                            status &= 0x00000000000FFFFF;
                            if(status != 0 || Marshal.GetLastWin32Error() != 0)
                            { 
                                //string value = string.Format("{0:X}", new Win32Exception(Marshal.GetLastWin32Error()).Message);
                                string value = string.Format("{0:X}", status);
                                Console.WriteLine($"device: COM{port} release failed with status=0x{value}.");
                            }
                        }
                        Console.WriteLine($"device: releasing COM ports completed.");
                    }
                    else
                    { 
                        Console.WriteLine($"no COM ports to release.");
                    }

                    long dsfdf1 = ComDBClose((UInt32)PHCOMDB);
                }

                // Search for new devices attaching
                RescanForHardwareChanges();

                string description = string.Empty;
                string deviceID = string.Empty;
                if (FindIngenicoDevice(ref description, ref deviceID))
                {
                    Console.WriteLine("\ndevice capabilities ----------------------------------------------------------------");
                    Console.WriteLine($"DESCRIPTION                  : {description}");
                    Console.WriteLine($"DEVICE ID                    : {deviceID}");

                    using (var searcher = new ManagementObjectSearcher("SELECT * FROM WIN32_SerialPort"))
                    {
                        string[] portNames = SerialPort.GetPortNames();
                        var ports = searcher.Get().Cast<ManagementBaseObject>().ToList();
                        var tList = (from n in portNames
                            join p in ports on n equals p["DeviceID"].ToString()
                            select n + " - " + p["Caption"]).ToList();

                        foreach(var device in tList)
                        { 
                            if(device.StartsWith("COM"))
                            {
                                string [] tokens = device.Split(' ');
                                ingenicoPort = Convert.ToInt32(tokens[0].TrimStart(new char [] { 'C', 'O', 'M' }));
                            }
                            Debug.WriteLine($"DEVICE PORT                  : {device}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return ingenicoPort;
        }

        private List<string> ReportUSBCommPorts(ref string instanceId)
        {
            List<string> ports = new List<string>();

            try
            {
                Console.WriteLine("--------------------------------------------------------------------------------------");
                Console.WriteLine("SEARCH FOR WIN32_SerialPort BY INSTANCE");
                Console.WriteLine("--------------------------------------------------------------------------------------");

                using (var searcher = new ManagementObjectSearcher("SELECT * FROM WIN32_SerialPort"))
                {
                    try
                    {
                        string[] portNames = SerialPort.GetPortNames();
                        var portsFound = searcher.Get().Cast<ManagementBaseObject>().ToList();
                        if(portsFound.Count > 0)
                        { 
                            Console.Write($"PORT NAMES                   : ");
                            string portValues = string.Empty;
                            foreach(var port in portsFound)
                            { 
                                portValues += $"{port},";
                            }
                            portValues = portValues.TrimEnd(',');
                            Console.WriteLine($"{portValues}");
                        
                            var tList = (from n in portNames
                                join p in portsFound on n equals p["DeviceID"].ToString()
                                select n + " - " + p["Caption"]).ToList();
                            if(tList.Count > 0)
                            { 
                                string portName = tList.FirstOrDefault();
                                Console.WriteLine($"DEVICE PORT                  : {portName}");
                                foreach (ManagementObject port in searcher.Get())
                                {
                                    string [] portCOM = portName.Split(' ');
                                    if (port["DeviceID"].ToString().Equals(portCOM[0]))
                                    { 
                                        instanceId = port["PNPDeviceID"].ToString();
                                        ports.Add(portCOM[0]);
                                    }
                                }
                            }
                            else
                            { 
                                Console.WriteLine($"DEVICE PORT                  : NONE FOUND");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"NONE FOUND");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"EXCEPTION                    : {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"device: ReportCommPorts() exception={ex.Message}");
            }

            return ports;
        }

        private List<string> ReportSerialCommPorts()
        {
            List<string> ports = new List<string>();

            try
            {
                using (var searcher = new ManagementObjectSearcher("root\\WMI", "SELECT * FROM MSSerial_PortName"))
                { 
                    Console.WriteLine("--------------------------------------------------------------------------------------");
                    Console.WriteLine("SEARCH FOR MSSerial_PortName BY INSTANCE");
                    Console.WriteLine("--------------------------------------------------------------------------------------");

                    if(searcher != null && searcher.Get() != null)
                    { 
                        foreach (ManagementObject queryObj in searcher.Get())
                        {
                            Console.WriteLine("GOT SEARCH");
                            //If the serial port's instance name contains USB 
                            //it must be a USB to serial device
                            if (queryObj["InstanceName"]?.ToString().Contains("USB") ?? false)
                            {
                                string instanceName = $"{queryObj["InstanceName"]}";
                                if(instanceName.StartsWith("USB\\VID_0B00"))
                                {
                                    instanceId = instanceName;
                                }
                                Console.WriteLine($"INSTANCE NAME                  : {instanceName}");
                                Console.WriteLine($"USB to SERIAL adapter/converter: {queryObj["PortName"]}");
                                ports.Add(queryObj["PortName"].ToString());
                                //SerialPort p = new SerialPort(port);
                                //p.PortName = "COM11";
                                //return port;
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("NONE FOUND.");
                    }
                }
            }
            catch (Exception ex)
            {
                if(ex.Message.IndexOf("Not supported", StringComparison.InvariantCultureIgnoreCase) != -1)
                {
                    Console.WriteLine("NONE FOUND.");
                }
                else
                { 
                    Console.WriteLine($"device: ReportCommPorts() exception='{ex.Message}'");
                }
            }

            return ports;
        }

        private void RescanForHardwareChanges()
        {
            Console.WriteLine("device: rescanning for hardware changes...");

            int pdnDevInst = 0;

            if (CM_Locate_DevNodeA(ref pdnDevInst, null, CM_LOCATE_DEVNODE_NORMAL) != CR_SUCCESS)
            { 
                Console.WriteLine("device: failed to locate hardware devices");
            }
            if (CM_Reenumerate_DevNode(pdnDevInst, CM_REENUMERATE_NORMAL) != CR_SUCCESS)
            { 
                Console.WriteLine("device: failed to reenumerate hardware devices");
            }
            else
            {
                Console.WriteLine("device: reenumerated hardware devices");
            }
        }

        private static void SelectiveUSBSuspendDisable()
        {
            // on battery: disabled
            string powerCmd = "powercfg";
            string powerCmdParams = "/SETDCVALUEINDEX SCHEME_CURRENT 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 0";
            string processExitCode = Utility.RunExternalExeElevated("C:\\Windows\\System32", powerCmd, powerCmdParams);

            // plugged in: disable
            powerCmdParams = "/SETACVALUEINDEX SCHEME_CURRENT 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 0";
            processExitCode = Utility.RunExternalExeElevated("C:\\Windows\\System32", powerCmd, powerCmdParams);
        }
    }

    enum ERROR_STATUS
    {
        ERROR_SUCCESS = 0,
        ERROR_ACCESS_DENIED = 5
    };

    enum REPORT_BYTES
    {
        CDB_REPORT_BITS  = 0,
        CDB_REPORT_BYTES = 1
    };
}
