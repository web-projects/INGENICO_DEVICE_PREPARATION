using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HidLibrary;

namespace DevicePreparation
{
    class DeviceInformation
    {
    }

    public class USBDeviceInfo
    {
        public USBDeviceInfo(string deviceID, string pnpDeviceID, string description)
        {
            this.DeviceID = deviceID;
            this.PnpDeviceID = pnpDeviceID;
            this.Description = description;
        }
        public string DeviceID { get; private set; }
        public string PnpDeviceID { get; private set; }
        public string Description { get; private set; }
    }
    public struct BoolStringDuple
    {
        public bool Item1 { get; set; }
        public string Item2 { get; set; }
        public BoolStringDuple(bool item1, string item2)
        {
            Item1 = item1;
            Item2 = item2;
        }
    }
}
