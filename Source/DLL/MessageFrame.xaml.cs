using System.Windows;
using System.Windows.Threading;

namespace DLL
{
    public partial class MessageFrame : Window
    {
        int height = 50;
        int width = 200;

        public MessageFrame(int indice_frame, string messaggio = "")
        {
            InitializeComponent();

            this.DataContext = this;
            this.Height = height;
            this.Width = width;
            this.Topmost = true;
            ShowInTaskbar = false;
            WindowStyle = WindowStyle.None;
            AllowsTransparency = true;
            riga.Content = messaggio;

            switch (indice_frame)
            {
                case 1:
                    this.Left = SystemParameters.FullPrimaryScreenWidth - width;
                    this.Top = SystemParameters.FullPrimaryScreenHeight - height;
                    break;

                case 2:
                    this.Left = SystemParameters.FullPrimaryScreenWidth - width;
                    this.Top = SystemParameters.FullPrimaryScreenHeight - (height * 2) - 5;
                    break;

                case 3:
                    this.Left = SystemParameters.FullPrimaryScreenWidth - width;
                    this.Top = SystemParameters.FullPrimaryScreenHeight - (height * 3) - 10;
                    break;

                case 4:
                    this.Left = SystemParameters.FullPrimaryScreenWidth - width;
                    this.Top = SystemParameters.FullPrimaryScreenHeight - (height * 4) - 15;
                    break;
            }
            this.Show();

        }
        public void scrivi(string cosa)
        {
            this.riga.Content = cosa;
            this.Dispatcher.Invoke(() => this.riga.Content = cosa,
                DispatcherPriority.Background);
            //this.UpdateLayout();
        }
        public void chiudi()
        {
            this.Close();
        }
    }
}
