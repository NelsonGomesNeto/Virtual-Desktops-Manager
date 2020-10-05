using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using System.Windows.Navigation;
using System.Windows.Shapes;
using WindowsDesktop;

namespace VirtualDesktopsManager
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private KeyHandler[] ghks = new KeyHandler[10];
        private VirtualDesktop[] virtualDesktops;
        private IDictionary<IntPtr, string> windows;
        private IDictionary<Guid, ICollection<IntPtr>> virtualDesktopsWindows = new Dictionary<Guid, ICollection<IntPtr>>();
        private bool activated = false;

        public MainWindow()
        {
            InitializeComponent();
            InitializeComObjects();

            IntPtr handle = new WindowInteropHelper(this).Handle;

            ComponentDispatcher.ThreadPreprocessMessage += MessageHandler;

            //EnableButton.BackColor = Color.Red;
            for (int i = 0; i < 10; i++)
                ghks[i] = new KeyHandler(Constants.WIN + Constants.CTRL + Constants.ALT, Constants.D0 + i, handle);
            AddAllKeyHandlers();

            RefreshVirtualDesktops();
        }

        protected override void OnStateChanged(EventArgs e)
        {
            if (WindowState == WindowState.Minimized)
                this.Hide();
            base.OnStateChanged(e);
        }

        private static async void InitializeComObjects()
        {
            try
            {
                await VirtualDesktopProvider.Default.Initialize(TaskScheduler.FromCurrentSynchronizationContext());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Failed to initialize.");
            }

            VirtualDesktop.CurrentChanged += (sender, args) => Debug.WriteLine($"Desktop changed: {args.NewDesktop.Id}");
        }

        private void RefreshVirtualDesktops()
        {
            windows = OpenWindowsGetter.GetOpenWindows();
            virtualDesktops = VirtualDesktop.GetDesktops();
            virtualDesktopsWindows.Clear();

            foreach (KeyValuePair<IntPtr, string> window in windows)
            {
                VirtualDesktop virtualDesktop = VirtualDesktop.FromHwnd(window.Key);
                if (virtualDesktop != null)
                {
                    Guid virtualDesktopGuid = virtualDesktop.Id;
                    if (!virtualDesktopsWindows.ContainsKey(virtualDesktopGuid))
                        virtualDesktopsWindows.Add(virtualDesktopGuid, new List<IntPtr>());
                    virtualDesktopsWindows[virtualDesktopGuid].Add(window.Key);
                }
            }

            for (int i = 0; i < virtualDesktops.Length; i++)
            {
                VirtualDesktop virtualDesktop = virtualDesktops[i];
                string windowsNames = "";
                foreach (IntPtr handles in virtualDesktopsWindows[virtualDesktop.Id])
                {
                    if (!string.IsNullOrEmpty(windowsNames))
                        windowsNames += "\n";
                    windowsNames += windows[handles];
                }
                TextBlock item = new TextBlock { Text = virtualDesktopsWindows[virtualDesktop.Id].Count.ToString(), ToolTip = windowsNames, TextAlignment = TextAlignment.Center, FontSize = 48 };
                WindowsGrid.Children.Add(item);
                if (i / 3 >= WindowsGrid.RowDefinitions.Count)
                    WindowsGrid.RowDefinitions.Add(new RowDefinition());
                Grid.SetColumn(item, i % 3);
                Grid.SetRow(item, i / 3);
            }
        }

        private void MessageHandler(ref MSG m, ref bool handled)
        {
            if (m.message == Constants.WM_HOTKEY_MSG_ID)
            {
                int key = (((int)m.lParam >> 16) & 0xFFFF);
                int modifier = (int)m.lParam & 0xFFFF;
                HandleHotkey(key, modifier);
            }
        }

        private void HandleHotkey(int key, int modifier)
        {
            if (key == Constants.D0)
                virtualDesktops.Last().Switch();
            else if (key - Constants.D1 < virtualDesktops.Length)
                virtualDesktops[key - Constants.D1].Switch();
        }

        private void AddAllKeyHandlers()
        {
            for (int i = 0; i < 10; i++)
                ghks[i].Register();
            EnableButton.Background = new SolidColorBrush(Colors.Green);
            EnableButton.Content = "Disable";
            activated = true;
        }

        private void RemoveAllKeyHandlers()
        {
            bool unregisterFailed = false;
            for (int i = 0; i < ghks.Length; i++)
                unregisterFailed |= !ghks[i].Unregiser();

            if (unregisterFailed)
                MessageBox.Show("Hotkey failed to unregister!");

            EnableButton.Background = new SolidColorBrush(Colors.Red);
            EnableButton.Content = "Enable";
            activated = false;
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            RemoveAllKeyHandlers();
        }

        private void EnableButton_Click(object sender, RoutedEventArgs e)
        {
            if (activated)
                RemoveAllKeyHandlers();
            else
                AddAllKeyHandlers();
        }
    }
}
