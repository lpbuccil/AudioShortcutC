using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AudioShortcutC.Data
{
    public class AudioDevice
    {
        public String Name { set; get; }
        public String DeviceId { set; get; }
        public String ControllerInfo { set; get; }

        public AudioDevice(String DeviceId, String Name, String ControllerInfo)
        {
            this.DeviceId = DeviceId;
            this.Name = Name;
            this.ControllerInfo = ControllerInfo;
        }
    }
}
