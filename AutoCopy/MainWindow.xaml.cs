using IWshRuntimeLibrary;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Runtime.CompilerServices;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace AutoCopy
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    [Serializable()]
    public partial class MainWindow : Window
    {
        private DispatcherTimer timerDelay;
        public MainWindow()
        {
            InitializeComponent();
            //Khai báo cho timer
            timerDelay = new DispatcherTimer();
            timerDelay.Tick += new EventHandler(Timer_Tick);
//
            this.ResizeMode = ResizeMode.CanMinimize;
            this.Title = "Auto Copy - KTSX Trà Vinh";
            //anr xuong taskbar            
            NotifyIcon notifyIcon = new NotifyIcon();
            notifyIcon.Icon = new System.Drawing.Icon("files.ico");
            ContextMenuStrip menu = new ContextMenuStrip();
            ToolStripMenuItem menuItem = new ToolStripMenuItem();
            menuItem.Text = "Exit";
            menuItem.Click += ExitMenuItem_Click;
            menu.Items.Add(menuItem);

            notifyIcon.ContextMenuStrip = menu;
            //notifyIcon.ContextMenuStrip.Items.Add("Exit");
            //load setting
            LoadSetting();
            //tu chay copy
            if (txt_delay.Text != String.Empty)
            {
                delayCopy(int.Parse(txt_delay.Text.ToString()));
            }
            
            //
            notifyIcon.Text = "Auto Copy";
            notifyIcon.Visible = true;
            notifyIcon.DoubleClick += NotifyIcon_DoubleClick;
            this.WindowState = WindowState.Minimized;
            if (WindowState == WindowState.Minimized) this.Hide();

        }

        private void ExitMenuItem_Click(object? sender, EventArgs e)
        {
            System.Environment.Exit(1);
        }
        //delay copy          
        private void delayCopy(int delay)
        {
            try
            {
                if (txt_delay.Text != string.Empty)
                {
                    if (txt_delay.Text == "0")
                    {
                        lbl_timer.Content = "Pause Copy";
                        timerDelay.Stop();
                    }
                    else
                    {


                        updateTimer(delay);
                        timerDelay.Interval = TimeSpan.FromSeconds(delay);
                        timerDelay.Start();
                    }

                }
                else
                {
                    lbl_timer.Content = "Pause Copy";
                    timerDelay.Stop();
                }
            }
            catch { }
        }
        private void Timer_Tick(object? sender, EventArgs e)
        {
            CopyTo(txt_copyfrom.Text, txt_copyto.Text);
            updateTimer(int.Parse(txt_delay.Text.ToString()));
        }
        //update time
        int timez;
        System.Windows.Forms.Timer timerUpdate = new System.Windows.Forms.Timer();
        private void updateTimer(int times)
        {
            timerUpdate.Tick += new EventHandler(updateTick);
            timerUpdate.Interval = 1000;
            timez = times;
            //Thread.Sleep(100);
            timerUpdate.Enabled = true;

            lbl_timer.Content = "Next copy: " + timez;
        }
        private void updateTick(object? sender, EventArgs e)
        {
            if (timez == 0)
            {
                timez = int.Parse(txt_delay.Text.ToString());
                timerUpdate.Enabled = false;
            }
            else
                timez--;
            lbl_timer.Content = "Next copy: " + timez;

        }
        //hide in minized
        protected override void OnStateChanged(EventArgs e)
        {
            if (WindowState == WindowState.Minimized) this.Hide();
            base.OnStateChanged(e);
        }
        protected override void OnClosing(CancelEventArgs e)
        {
            e.Cancel = true;
            this.Hide();
            base.OnClosing(e);
        }
        //double click to open
        private void NotifyIcon_DoubleClick(object? sender, EventArgs e)
        {
            this.Show();
            this.WindowState = WindowState.Normal;
        }

        //

        //
        void menu1Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void btn_copyfrom_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.ShowDialog();
            txt_copyfrom.Text = dialog.SelectedPath;
        }

        private void btn_copyto_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog dialog1 = new FolderBrowserDialog();
            dialog1.ShowDialog();
            txt_copyto.Text = dialog1.SelectedPath;

        }
        //dang ky auto start
        private static RegistryKey registryKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                registryKey.SetValue("AutoCopy", System.Windows.Forms.Application.ExecutablePath);
                //System.IO.File.Copy(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\AutoCopy.lnk", Environment.GetFolderPath(Environment.SpecialFolder.Startup) + "AutoCopy.lnk");
                CheckExistShortcut();
            }
            catch { }
        }


        private void CheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            registryKey.DeleteValue("AutoCopy", false);
        }
        //create shortcut        
        public static void CreateShortcut(string shortcutname, string shortcutPath, string targetFileLocation)
        {
            string shortcutLocation = System.IO.Path.Combine(shortcutPath, shortcutname + ".lnk");
            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutLocation);

            shortcut.Description = "AutoCopy KTSX";

            shortcut.IconLocation = System.Windows.Forms.Application.StartupPath + @"\files.ico";
            shortcut.TargetPath = targetFileLocation;
            shortcut.WorkingDirectory = System.Windows.Forms.Application.StartupPath;
            shortcut.Save();
        }
        public void CheckExistShortcut()
        {
            if (System.IO.File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Startup) + @"\AutoCopy.lnk") == false)
            {
                CreateShortcut("ExportExcel", Environment.GetFolderPath(Environment.SpecialFolder.Startup), Assembly.GetExecutingAssembly().Location.Substring(0, Assembly.GetExecutingAssembly().Location.Length - 12) + "AutoCopy.exe");
            }
        }
        //Lưu cài đặt lại

        private void SaveSetting(UserSetting userSetting)
        {
            using (var stream = System.IO.File.OpenWrite("data.bin"))
            {
                var formatter = new BinaryFormatter();
                formatter.Serialize(stream, userSetting);
                stream.Close();
            }
        }
        //Load cài đặt lên
        private void LoadSetting()
        {
            UserSetting userS = new UserSetting();
            using (var stream = System.IO.File.OpenRead("data.bin"))
            {
                var formatter = new BinaryFormatter();
                userS = (UserSetting)formatter.Deserialize(stream);
                stream.Close();
            }
            txt_copyfrom.Text = userS.copyFrom;
            txt_copyto.Text = userS.copyTo;
            txt_delay.Text = userS.delayCopy.ToString();
            if (userS.startWithWindow == "true")
            {
                chk_startwithwindow.IsChecked = true;
            }
            else
                chk_startwithwindow.IsChecked = false;
        }
        [Serializable]
        class UserSetting
        {
            public string? copyFrom { get; set; }
            public string? copyTo { get; set; }
            public int delayCopy { get; set; }
            public string? startWithWindow { get; set; }

        }

        private void Window_Closed(object sender, CancelEventArgs e)
        {

        }

        private void txt_delay_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
        }

        //copy file
        private void CopyTo(string copyfrom, string copyto)
        {
            if (txt_copyfrom.Text != string.Empty && txt_copyto.Text != string.Empty)
            {
                string[] filelst = new string[] { };

                filelst = Directory.GetFiles(copyfrom);
                for (int i = 0; i < filelst.Length; i++)
                {
                    CreateFolder(filelst[i]);
                }
                UpdateLog();
            }
            else
                lbl_timer.Content = "Select folder for copy procress";

        }
        //Tao thu muc chua file
        private void CreateFolder(string lstfile)
        {
            try
            {
                string getfilename = System.IO.Path.GetFileNameWithoutExtension(lstfile);
                if (getfilename.Split('_').Length == 2)
                {
                    string path = txt_copyto.Text + "\\" + getfilename.Substring(getfilename.Count() - 8, 4) + "\\Tháng " + getfilename.Substring(getfilename.Count() - 4, 2);
                    Directory.CreateDirectory(path);
                    if (lstfile.Contains("Error_"))
                    {

                        Directory.CreateDirectory(path + "\\Error");
                        System.IO.File.Copy(lstfile, path + "\\Error\\" + lstfile.Split(@"\")[lstfile.Split(@"\").Count() - 1], true);
                    }
                    else
                    {
                        Directory.CreateDirectory(path + "\\Testing");
                        System.IO.File.Copy(lstfile, path + "\\Testing\\" + lstfile.Split(@"\")[lstfile.Split(@"\").Count() - 1], true);
                    }
                }

            }
            catch
            {
                //System.Windows.Forms.MessageBox.Show("Error appears !!"); 
            }
        }
        //Check file is opened
        protected virtual bool IsFileLocked(FileInfo file)
        {
            try
            {
                using (FileStream stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    stream.Close();
                }
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }

            //file is not locked
            return false;
        }
        //

        private void btn_saveSetting_Click(object sender, RoutedEventArgs e)
        {
            if (txt_copyfrom.Text != string.Empty || txt_copyto.Text != string.Empty || txt_delay.Text != string.Empty)
            {

                UserSetting setting = new UserSetting();
                setting.copyFrom = txt_copyfrom.Text.ToString();
                setting.copyTo = txt_copyto.Text.ToString();
                setting.delayCopy = int.Parse(txt_delay.Text.ToString());
                if (chk_startwithwindow.IsChecked == true)
                {
                    setting.startWithWindow = "true";
                }
                else
                    setting.startWithWindow = "false";

                SaveSetting(setting);
                System.Windows.Forms.MessageBox.Show("Save success!!!");
                delayCopy(setting.delayCopy);
            }
            else
                System.Windows.Forms.MessageBox.Show("Please choose folder before save it!!!");

        }

        private void txt_delay_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void btn_logview_Click(object sender, RoutedEventArgs e)
        {
            Process.Start("notepad.exe", Assembly.GetExecutingAssembly().Location.Substring(0, Assembly.GetExecutingAssembly().Location.Length - 12) + "log.txt");
        }
        //
        public void UpdateLog()
        {
            //var file  = System.IO.File.OpenWrite(Assembly.GetExecutingAssembly().Location.Substring(0, Assembly.GetExecutingAssembly().Location.Length - 12) + "log.txt");
            System.IO.File.AppendAllText(Assembly.GetExecutingAssembly().Location.Substring(0, Assembly.GetExecutingAssembly().Location.Length - 12) + "log.txt", DateTime.Now.ToString("F") + Environment.NewLine);

        }
    }
}
