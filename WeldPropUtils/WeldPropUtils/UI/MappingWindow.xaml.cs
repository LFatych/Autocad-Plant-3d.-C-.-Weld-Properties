using System;
using System.Windows;
using WeldPropUtils.Settings;

namespace WeldPropUtils.UI
{
    public partial class MappingWindow : Window
    {
        private static readonly Uri DarkTheme = new Uri("/WeldPropUtils;component/UI/Themes/SpdsDark.xaml", UriKind.Relative);
        private static readonly Uri LightTheme = new Uri("/WeldPropUtils;component/UI/Themes/SpdsLight.xaml", UriKind.Relative);

        private readonly UserPreferences _prefs;
        private bool _dark;

        public MappingWindow(MappingViewModel viewModel, bool dark, UserPreferences prefs)
        {
            InitializeComponent();
            _prefs = prefs;
            DataContext = viewModel;
            ApplyTheme(dark);
            viewModel.CloseRequested += (saved, update) =>
            {
                Saved = saved;
                UpdateWeldsRequested = update;
                Close();
            };
        }

        public bool Saved { get; private set; }
        public bool UpdateWeldsRequested { get; private set; }

        // The theme dictionary is MergedDictionaries[0] (see MappingWindow.xaml).
        private void ApplyTheme(bool dark)
        {
            _dark = dark;
            Resources.MergedDictionaries[0] = new ResourceDictionary { Source = dark ? DarkTheme : LightTheme };
            ThemeButton.Content = dark ? "Light theme" : "Dark theme";
        }

        private void OnToggleTheme(object sender, RoutedEventArgs e)
        {
            ApplyTheme(!_dark);
            _prefs.Theme = _dark ? UserPreferences.ThemeDark : UserPreferences.ThemeLight;
            _prefs.Save();
        }
    }
}
