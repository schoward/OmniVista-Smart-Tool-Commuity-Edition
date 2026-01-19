using System;
using System.Linq;
using System.Reflection;
using System.Windows;
using static PoEWizard.Data.Constants;

namespace PoEWizard.Components
{
    /// <summary>
    /// Interaction logic for AboutBox.xaml
    /// </summary>
    public partial class AboutBox : Window
    {
        public AboutBox()
        {
            InitializeComponent();
            if (MainWindow.Theme == ThemeType.Dark)
            {
                Resources.MergedDictionaries.Remove(Resources.MergedDictionaries[0]);
            }
            else
            {
                Resources.MergedDictionaries.Remove(Resources.MergedDictionaries[1]);
            }
            Resources.MergedDictionaries.Remove(Resources.MergedDictionaries[1]);
            Resources.MergedDictionaries.Add(MainWindow.Strings);

            Assembly assembly = Assembly.GetExecutingAssembly();
            string version = assembly.GetName().Version.ToString();
            string ver = (string)Resources.MergedDictionaries[1]["i18n_ver"];

            _version.Text = $"{ver} {string.Join(".", version.Split('.').ToList().Take(3))} CE";
            _copyRight.Text = "© 2024 ALE USA, 2024-2026 Community Edition. MIT License.";
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
