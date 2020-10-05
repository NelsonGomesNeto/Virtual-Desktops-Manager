using System;
using System.Collections.Generic;
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

            VirtualDesktop.CurrentChanged += (sender, args) => System.Diagnostics.Debug.WriteLine($"Desktop changed: {args.NewDesktop.Id}");
        }

        private void RefreshVirtualDesktops()
        {
            virtualDesktops = VirtualDesktop.GetDesktops();
            //virtualDesktopCountLabel.Text = virtualDesktops.Length.ToString();
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
