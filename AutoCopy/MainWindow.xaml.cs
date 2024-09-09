using IWshRuntimeLibrary;
using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization.Formatters.Binary;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Threading;
using Application = System.Windows.Forms.Application;
using File = System.IO.File;

namespace AutoCopy
{
    [Serializable]
    public partial class MainWindow : Window
    {
        private DispatcherTimer timerDelay;
        private static RegistryKey registryKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
        private int timez;
        private System.Windows.Forms.Timer timerUpdate = new System.Windows.Forms.Timer();

        public MainWindow()
        {
            InitializeComponent();
            InitializeTimer();
            InitializeNotifyIcon();
            LoadSetting();

            if (!string.IsNullOrEmpty(txt_delay.Text))
            {
                delayCopy(int.Parse(txt_delay.Text));
            }
        }

        private void InitializeTimer()
        {
            timerDelay = new DispatcherTimer();
            timerDelay.Tick += Timer_Tick;
            this.ResizeMode = ResizeMode.CanMinimize;
            this.Title = "Auto Copy - KTSX Trà Vinh";
        }

        private void InitializeNotifyIcon()
        {
            NotifyIcon notifyIcon = new NotifyIcon
            {
                Icon = new System.Drawing.Icon("files.ico"),
                Visible = true,
                Text = "Auto Copy"
            };
            notifyIcon.DoubleClick += NotifyIcon_DoubleClick;

            ContextMenuStrip menu = new ContextMenuStrip();
            ToolStripMenuItem menuItem = new ToolStripMenuItem { Text = "Exit" };
            menuItem.Click += ExitMenuItem_Click;
            menu.Items.Add(menuItem);

            notifyIcon.ContextMenuStrip = menu;
            this.WindowState = WindowState.Minimized;
            if (WindowState == WindowState.Minimized) this.Hide();
        }

        private void ExitMenuItem_Click(object? sender, EventArgs e)
        {
            Environment.Exit(1);
        }

        private async void Timer_Tick(object? sender, EventArgs e)
        {
            await CopyToAsync(txt_copyfrom.Text, txt_copyto.Text);
            updateTimer(int.Parse(txt_delay.Text));
        }

        private void delayCopy(int delay)
        {
            if (delay == 0)
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

        private void updateTimer(int times)
        {
            timerUpdate.Tick += updateTick;
            timerUpdate.Interval = 1000;
            timez = times;
            timerUpdate.Enabled = true;
            lbl_timer.Content = "Next copy: " + timez;
        }

        private void updateTick(object? sender, EventArgs e)
        {
            if (timez == 0)
            {
                timez = int.Parse(txt_delay.Text);
                timerUpdate.Enabled = false;
            }
            else
            {
                timez--;
            }
            lbl_timer.Content = "Next copy: " + timez;
        }

        protected override void OnStateChanged(EventArgs e)
        {
            if (WindowState == WindowState.Minimized) this.Hide();
            base.OnStateChanged(e);
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
            this.Hide();
            base.OnClosing(e);
        }

        private void NotifyIcon_DoubleClick(object? sender, EventArgs e)
        {
            this.Show();
            this.WindowState = WindowState.Normal;
        }

        private async Task CopyToAsync(string copyfrom, string copyto)
        {
            if (!string.IsNullOrEmpty(copyfrom) && !string.IsNullOrEmpty(copyto))
            {
                string[] filelst = Directory.GetFiles(copyfrom);

                await Task.Run(() =>
                {
                    foreach (string file in filelst)
                    {
                        CreateFolder(file);
                    }
                });

                Dispatcher.Invoke(() => lbl_timer.Content = "Copy completed!");
                UpdateLog();
            }
            else
            {
                lbl_timer.Content = "Select folder for copy process";
            }
        }

        private void CreateFolder(string lstfile)
        {
            try
            {
                string getfilename = Path.GetFileNameWithoutExtension(lstfile);
                if (getfilename.Split('_').Length == 2)
                {
                    string path = Path.Combine(txt_copyto.Text, getfilename[^8..^4], "Tháng " + getfilename[^4..^2]);
                    Directory.CreateDirectory(path);
                    string subfolder = lstfile.Contains("Error_") ? "Error" : "Testing";
                    Directory.CreateDirectory(Path.Combine(path, subfolder));
                    System.IO.File.Copy(lstfile, Path.Combine(path, subfolder, Path.GetFileName(lstfile)), true);
                }
            }
            catch (Exception ex)
            {
                // Log exception if necessary
            }
        }

        private void btn_copyfrom_Click(object sender, RoutedEventArgs e)
        {
            using FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txt_copyfrom.Text = dialog.SelectedPath;
            }
        }

        private void btn_copyto_Click(object sender, RoutedEventArgs e)
        {
            using FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txt_copyto.Text = dialog.SelectedPath;
            }
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                registryKey.SetValue("AutoCopy", System.Windows.Forms.Application.ExecutablePath);
                CheckExistShortcut();
            }
            catch (Exception ex)
            {
                // Log exception if necessary
            }
        }

        private void CheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            registryKey.DeleteValue("AutoCopy", false);
        }

        public static void CreateShortcut(string shortcutname, string shortcutPath, string targetFileLocation)
        {
            string shortcutLocation = Path.Combine(shortcutPath, shortcutname + ".lnk");
            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutLocation);

            shortcut.Description = "AutoCopy KTSX";
            shortcut.IconLocation = Path.Combine(Application.StartupPath, "files.ico");
            shortcut.TargetPath = targetFileLocation;
            shortcut.WorkingDirectory = Application.StartupPath;
            shortcut.Save();
        }

        public void CheckExistShortcut()
        {
            string shortcutPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Startup), "AutoCopy.lnk");
            if (!File.Exists(shortcutPath))
            {
                CreateShortcut("AutoCopy", Environment.GetFolderPath(Environment.SpecialFolder.Startup), Assembly.GetExecutingAssembly().Location);
            }
        }

        private void SaveSetting(UserSetting userSetting)
        {
            using FileStream stream = File.OpenWrite("data.bin");
            BinaryFormatter formatter = new BinaryFormatter();
            formatter.Serialize(stream, userSetting);
        }

        private void LoadSetting()
        {
            if (File.Exists("data.bin"))
            {
                using FileStream stream = File.OpenRead("data.bin");
                BinaryFormatter formatter = new BinaryFormatter();
                UserSetting userS = (UserSetting)formatter.Deserialize(stream);

                txt_copyfrom.Text = userS.copyFrom;
                txt_copyto.Text = userS.copyTo;
                txt_delay.Text = userS.delayCopy.ToString();
                chk_startwithwindow.IsChecked = userS.startWithWindow == "true";
            }
        }

        private void btn_saveSetting_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(txt_copyfrom.Text) && !string.IsNullOrEmpty(txt_copyto.Text) && !string.IsNullOrEmpty(txt_delay.Text))
            {
                UserSetting setting = new UserSetting
                {
                    copyFrom = txt_copyfrom.Text,
                    copyTo = txt_copyto.Text,
                    delayCopy = int.Parse(txt_delay.Text),
                    startWithWindow = chk_startwithwindow.IsChecked == true ? "true" : "false"
                };

                SaveSetting(setting);
                System.Windows.Forms.MessageBox.Show("Save successful!");
                delayCopy(setting.delayCopy);
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Please choose folders and set delay before saving!");
            }
        }

        public void UpdateLog()
        {
            string logPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "log.txt");
            File.AppendAllText(logPath, DateTime.Now.ToString("F") + Environment.NewLine);
        }

        [Serializable]
        class UserSetting
        {
            public string? copyFrom { get; set; }
            public string? copyTo { get; set; }
            public int delayCopy { get; set; }
            public string? startWithWindow { get; set; }
        }
    }
}