using AudioShortcutC.Data;
using Microsoft.Win32;
using System;
using System.IO;
using System.Collections.Generic;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using System.Net;
using System.IO.Compression;
using IWshRuntimeLibrary;
using File = System.IO.File;
using System.Drawing;
using System.Windows.Media;
using System.Windows.Interop;

namespace AudioShortcutC
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 
    
    public partial class MainWindow : Window
    {

        String NirCMDWebsite = "http://www.nirsoft.net/utils/nircmd.zip";
        String DirectoryPath;

        public MainWindow()
        {
            InitializeComponent();
            PopulateDevices(ListBoxAudioDevices);
            RadioButtonSpeaker.IsChecked = true;
        }


        private void PopulateDevices(ListBox listBox)
        {

            var AudioDeviceList = new List<AudioDevice>();
            try
            {
                using (var hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
                {
                    using (var myKey =
                        hklm.OpenSubKey(
                            "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\MMDevices\\Audio\\Render"))
                    {
                        if (myKey == null)
                        {
                            Console.WriteLine("Could not open Registry");
                            return;
                        }

                        foreach (var reg in myKey.GetSubKeyNames())
                        {
                            if ((int)(myKey.OpenSubKey(reg).GetValue("DeviceState")) == 1)
                            {
                                var curAudioDevice = new AudioDevice(reg, null, null);
                                myKey.OpenSubKey(reg).OpenSubKey("Properties")
                                    .GetValue("{a45c254e-df1c-4efd-8020-67d146a850e0}");
                                curAudioDevice.Name = myKey.OpenSubKey(reg).OpenSubKey("Properties")
                                    .GetValue("{a45c254e-df1c-4efd-8020-67d146a850e0},2").ToString();
                                curAudioDevice.ControllerInfo = myKey.OpenSubKey(reg)
                                    .OpenSubKey("Properties")
                                    .GetValue("{b3f8fa53-0004-438e-9003-51a46e139bfc},6").ToString();
                                Console.WriteLine(curAudioDevice.Name);
                                AudioDeviceList.Add(curAudioDevice);
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                Console.WriteLine(e.StackTrace);
                throw;
            }

            List<AudioDevice> DistinctList = AudioDeviceList
            .GroupBy(a => a.Name)
            .Select(g => g.First())
            .ToList();

            if (AudioDeviceList.Count() != DistinctList.Count())
            {
                LabelStatus.Text = "More than one device share the same name. If you intend to use either one, please double click on one of the devices to change its name";
            }
            else {
                LabelStatus.Text = "";
            }


            listBox.ItemsSource = AudioDeviceList;
            listBox.DisplayMemberPath = "Name";
            listBox.SelectedValuePath = "DeviceId";

        }

        private void ButtonSubmit_Click(object sender, RoutedEventArgs e)
        {

            DirectoryPath = null;

            using (var dialog = new CommonOpenFileDialog())
            {
                dialog.IsFolderPicker = true;


                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    DirectoryPath = dialog.FileName;
                }
                else
                {
                    return;
                }
            }

            Boolean filesExist = true;
            if (!File.Exists(DirectoryPath + "\\nircmd.exe")) filesExist = false;
            if (!File.Exists(DirectoryPath + "\\nircmdc.exe")) filesExist = false;
            if (!File.Exists(DirectoryPath + "\\NirCmd.chm")) filesExist = false;

            if (!filesExist)
            {
                using (var client = new WebClient())
                {
                    client.DownloadFile(NirCMDWebsite, DirectoryPath + "\\nircmd.zip");
                }

                //Remove files before extracting
                if (File.Exists(DirectoryPath + "\\nircmd.exe"))
                { 
                    File.Delete(DirectoryPath + "\\nircmd.exe");
                }
                if (File.Exists(DirectoryPath + "\\nircmdc.exe"))
                {
                    File.Delete(DirectoryPath + "\\nircmdc.exe");
                }
                if (File.Exists(DirectoryPath + "\\NirCmd.chm"))
                {
                    File.Delete(DirectoryPath + "\\NirCmd.chm");
                }
                ZipFile.ExtractToDirectory(DirectoryPath + "\\nircmd.zip", DirectoryPath);

                if (File.Exists(DirectoryPath + "\\nircmd.zip"))
                {
                    File.Delete(DirectoryPath + "\\nircmd.zip");
                }
            }

            buttonCreate.IsEnabled = true;



            
        }

        private void ButtonCreate_Click(object sender, RoutedEventArgs e)
        {
            


            LabelStatus.Text = "";
            if (ListBoxAudioDevices.SelectedIndex == -1)
            {
                LabelStatus.Text = "Please select an Audio device";
                return;
            }

            AudioDevice SelectedDevice = (AudioDevice)ListBoxAudioDevices.SelectedItem;
            int SameDevicesNames = 0;

            foreach (Object temp in ListBoxAudioDevices.Items)
            {
                if (((AudioDevice)temp).Name.Equals(SelectedDevice.Name))
                {
                    SameDevicesNames++;
                }
            }

            if(SameDevicesNames > 1)
            {
                ChangeDeviceName changeDeviceName = new ChangeDeviceName(SelectedDevice);
                changeDeviceName.labelControllerInformation.Content = SelectedDevice.ControllerInfo;
                changeDeviceName.textBoxDeviceName.Text = SelectedDevice.Name;
                changeDeviceName.ShowDialog();
                PopulateDevices(ListBoxAudioDevices);
                return;
            }


            String BatchFileName = SelectedDevice.Name.Replace(" ", "").Replace("-", "");

            LabelStatus.Text = BatchFileName;

            if (File.Exists($"{DirectoryPath}\\{BatchFileName}.bat"))
            {
                File.Delete($"{DirectoryPath}\\{BatchFileName}.bat");
            }

            using (StreamWriter sw = System.IO.File.CreateText($"{DirectoryPath}\\{BatchFileName}.bat"))
            {
                sw.WriteLine("@ECHO OFF");
                sw.WriteLine($"{DirectoryPath}\\NIRCMDC setdefaultsounddevice \"{SelectedDevice.Name}\" 1");
                sw.WriteLine($"{DirectoryPath}\\NIRCMDC setdefaultsounddevice \"{SelectedDevice.Name}\" 2");

            }

            object shDesktop = (object)"Desktop";
            WshShell wsh = new WshShell();
            string shortcutAddress = (string)wsh.SpecialFolders.Item(ref shDesktop) + $@"\{BatchFileName}.lnk";
            IWshShortcut shortcut = (IWshShortcut)wsh.CreateShortcut(shortcutAddress);
            shortcut.Description = $"Shortcut for {SelectedDevice.Name}";

            shortcut.TargetPath = Environment.GetFolderPath(Environment.SpecialFolder.Windows) + @"\system32\cmd.exe";
            shortcut.Arguments = $"/c \"{DirectoryPath}\\{BatchFileName}.bat\"";
            shortcut.WorkingDirectory = wsh.ExpandEnvironmentStrings(DirectoryPath);
            if (RadioButtonHeadset.IsChecked == true)
            {
                shortcut.IconLocation = @"%systemroot%\system32\ddores.dll,89";
            }
            else
            {
                shortcut.IconLocation = @"%systemroot%\system32\ddores.dll,88";
            }

            shortcut.Save();




        }

        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
            Bitmap temp = new Bitmap(Properties.Resources.ico30111);
            ImageSpeakers.Source = Convert(temp);
            temp = new Bitmap(Properties.Resources.ico30121);
            ImageHeadset.Source = Convert(temp);

        }


        public BitmapImage Convert(Bitmap src)
        {
            MemoryStream ms = new MemoryStream();
            ((System.Drawing.Bitmap)src).Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
            BitmapImage image = new BitmapImage();
            image.BeginInit();
            ms.Seek(0, SeekOrigin.Begin);
            image.StreamSource = ms;
            image.EndInit();
            return image;
        }

        private void ListBoxAudioDevices_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            AudioDevice SelectedDevice = (AudioDevice)ListBoxAudioDevices.SelectedItem;
            if(SelectedDevice != null)
            {
                ChangeDeviceName changeDeviceName = new ChangeDeviceName(SelectedDevice);
                changeDeviceName.labelControllerInformation.Content = SelectedDevice.ControllerInfo;
                changeDeviceName.textBoxDeviceName.Text = SelectedDevice.Name;
                changeDeviceName.ShowDialog();

                PopulateDevices(ListBoxAudioDevices);
            }
            
        }

        private void ImageHeadset_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            RadioButtonHeadset_Copy2.IsChecked = true;
        }

        private void ImageSpeakers_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            RadioButtonSpeaker.IsChecked = true;
        }
    }
}
