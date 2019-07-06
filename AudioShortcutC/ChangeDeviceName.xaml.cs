using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.Win32;
using System;
using AudioShortcutC.Data;
using System.Security.AccessControl;
using System.Security.Principal;

namespace AudioShortcutC
{
    /// <summary>
    /// Interaction logic for ChangeDeviceName.xaml
    /// </summary>
    public partial class ChangeDeviceName : Window
    {
        AudioDevice selectedAudioDevice = null;
        public ChangeDeviceName(AudioDevice selectedDevice)
        {
            selectedAudioDevice = selectedDevice;
            InitializeComponent();
        }

        private void ButtonChangeNameSubmit_Click(object sender, RoutedEventArgs e)
        {

            var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            if( principal.IsInRole(WindowsBuiltInRole.Administrator) == false)
            {
                MessageBox.Show("not admin");
            }

            try
            {
                String path = $@"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render\{selectedAudioDevice.DeviceId}\Properties\";
               
                    
                   using (RegistryKey regKey = Registry.LocalMachine.OpenSubKey(path, RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryRights.SetValue))
                    {
                        regKey.SetValue("{a45c254e-df1c-4efd-8020-67d146a850e0},2", textBoxDeviceName.Text);
                    }
                
                
            }catch(Exception ex)
            {
                Console.WriteLine(ex);
                Console.WriteLine(ex.StackTrace);
                throw;
            }
            Close();
        }
    }
}
