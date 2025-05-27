using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace DLL
{
    public partial class InputBox : Window
    {
        public dynamic result;
        public bool ret = false;
        private List<UIElement> elements = [];
        /// <summary>
        /// Finestra di Input
        /// </summary>
        /// <param name="TextBox">TextBox che si vuole nell'InputBox</param>
        /// <param name="titolo">Titolo della finestra</param>
        public InputBox(string TextBox, string titolo = "InputBox")
        {
            InitializeComponent();

            this.Topmost = true;
            this.Title = titolo;

            TextBlock tBlock = new TextBlock
            {
                Name = "tBlock",
                Text = TextBox,
                Margin = new Thickness(5),
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(tBlock, 0);
            container.Children.Add(tBlock);

            TextBox tBox = new()
            {
                Name = "tBox",
                HorizontalContentAlignment = HorizontalAlignment.Stretch,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                Margin = new Thickness(5),
            };
            tBox.KeyUp += Enter;
            Grid.SetColumn(tBox, 1);
            container.Children.Add(tBox);

            elements.Add(tBox);

            this.Loaded += InputBox_Loaded;
            this.ShowDialog();
            this.Focusable = true;
            this.Focus();
        }
        /// <summary>
        /// Finestra di Input
        /// </summary>
        /// <param name="lista_textbox">Lista dei TextBox che si vogliono nell'InputBox</param>
        /// <param name="lista_combobox">
        /// Dizionario dei ComboBox che si vogliono nell'InputBox di tipo Dictionary<string, List<string>
        /// dove string è il label della ComboBox e la lista è il suo contenuto</param>
        /// <param name="titolo">
        /// Titolo della finestra
        /// </param>
        public InputBox(List<string> lista_textbox = null, Dictionary<string, List<string>> lista_combobox = null, string titolo = "InputBox")
        {
            InitializeComponent();
            this.Topmost = true;
            this.Title = titolo;
            int indice_riga = 0;

            if (lista_combobox != null)
            {
                foreach (string text in lista_combobox.Keys)
                {
                    this.container.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Auto) });

                    Label label = new Label() { Content = text };
                    ComboBox combo = new ComboBox() { Name = text.Replace(" ", "_").ToLower(), Margin = new Thickness(5) };
                    foreach (string value in lista_combobox[text])
                    {
                        combo.Items.Add(new ComboBoxItem() { Content = value });
                    }

                    Grid.SetRow(label, indice_riga);
                    Grid.SetRow(combo, indice_riga);

                    Grid.SetColumn(label, 0);
                    Grid.SetColumn(combo, 1);

                    this.container.Children.Add(label);
                    this.container.Children.Add(combo);

                    this.elements.Add(combo);

                    indice_riga += 1;
                }
            }

            if (lista_textbox != null)
            {
                foreach (string text in lista_textbox)
                {
                    this.container.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Auto) });

                    Label label = new() { Content = text };
                    TextBox textBox = new() { Name = text.Replace(" ", "_").ToLower(), Margin = new Thickness(5) };

                    textBox.KeyUp += Enter;

                    Grid.SetRow(label, indice_riga);
                    Grid.SetRow(textBox, indice_riga);

                    Grid.SetColumn(label, 0);
                    Grid.SetColumn(textBox, 1);

                    this.container.Children.Add(label);
                    this.container.Children.Add(textBox);

                    this.elements.Add(textBox);
                    indice_riga += 1;
                }
            }

            this.Loaded += InputBox_Loaded;
            this.ShowDialog();
            this.Focusable = true;
            this.Focus();
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Button btn = (Button)sender;

            if (btn.Content.ToString() == "Ok")
            {
                this.ret = true;

                Dictionary<string, string> tmp = [];
                foreach (UIElement elem in elements)
                {
                    switch (elem.GetType())
                    {
                        case Type tipo when tipo == typeof(ComboBox):
                            ComboBox combo = elem as ComboBox;
                            tmp.Add(combo.Name, (combo.SelectedItem as ComboBoxItem).Content.ToString());
                            break;

                        case Type tipo when tipo == typeof(TextBox):
                            TextBox textBox = elem as TextBox;
                            tmp.Add(textBox.Name, textBox.Text);
                            break;
                    }
                }

                if (tmp.Count > 1)
                {
                    result = tmp;
                }
                else
                {
                    result = tmp.Values.ToList()[0];
                }

                this.Close();
            }
            else
            {
                this.Close();
            }
        }
        private void Enter(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                this.OK.Focus();
                this.ret = true;

                Dictionary<string, string> tmp = [];
                foreach (UIElement elem in elements)
                {
                    switch (elem.GetType())
                    {
                        case Type tipo when tipo == typeof(ComboBox):
                            ComboBox combo = elem as ComboBox;
                            tmp.Add(combo.Name, (combo.SelectedItem as ComboBoxItem).Content.ToString());
                            break;

                        case Type tipo when tipo == typeof(TextBox):
                            TextBox textBox = elem as TextBox;
                            tmp.Add(textBox.Name, textBox.Text);
                            break;
                    }
                }

                if (tmp.Count > 1)
                {
                    result = tmp;
                }
                else
                {
                    result = tmp.Values.ToList()[0];
                }

                this.Close();
            }
            else if (e.Key == Key.Escape)
            {
                this.Close();
            }
        }
        private void InputBox_Loaded(object sender, RoutedEventArgs e)
        {
            this.container.Children[1].Focus();
        }
    }
}
