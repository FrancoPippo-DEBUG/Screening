using System.Diagnostics;
using System.Windows;

namespace DLL
{
    /// <summary>
    /// Logica di interazione per Win.xaml
    /// </summary>
    public partial class Win : Window
    {
        public string browser = "CR";
        public Win(string titolo = "IcoWin", string browser = "CR")
        {
            AppContext.SetSwitch("Switch.System.Windows.Controls.Text.UseAdornerForTextboxSelectionRendering", false); // CAMBIA COLORE SELEZIONE TEXTBOX

            InitializeComponent();
            this.Height = 200;
            this.Width = 400;
            this.WindowStyle = WindowStyle.None;
            this.Visibility = Visibility.Collapsed;
            this.Opacity = 0;
            this.AllowsTransparency = true;
            this.ShowInTaskbar = true;
            this.Title = titolo;
            this.browser = browser;
            this.Closed += this.Chiudi;
            this.Show();
        }
        public void Chiudi(object sender = null, EventArgs e = null)
        {
            try
            {
                Process[] proc;
                switch (this.browser)
                {
                    case "CR":
                        proc = Process.GetProcessesByName("chromedriver");
                        foreach (Process p in proc)
                        {
                            p.Kill();
                        }
                        break;

                    case "ED":
                        proc = Process.GetProcessesByName("geckodriver");
                        foreach (Process p in proc)
                        {
                            p.Kill();
                        }
                        break;

                    case "FF":
                        proc = Process.GetProcessesByName("msedgedriver");
                        foreach (Process p in proc)
                        {
                            p.Kill();
                        }
                        break;
                }

                proc = Process.GetProcessesByName("excel");
                foreach (Process p in proc)
                {
                    p.Kill();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Non sono riuscito a terminare correttamente il processo");
            }
            Environment.Exit(0);
        }
    }
}
