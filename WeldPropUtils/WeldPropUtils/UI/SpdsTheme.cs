using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace WeldPropUtils.UI
{
    // SPDS dark and light themes (colours: .claude/skills/spds-wpf-ui). Brushes are resources named "Spds.<Key>";
    // elements reference them with SetResourceReference so a theme switch repaints everything at once.
    // The dictionary goes into the window's own resources, never Application.Current (AutoCAD owns that).
    public static class SpdsTheme
    {
        public const string HeadingFont = "Century Gothic, Segoe UI";
        public const string BodyFont = "Segoe UI";
        public const string MonoFont = "Consolas";

        private static readonly Dictionary<string, string> Dark = new Dictionary<string, string>
        {
            ["WindowBg"] = "#1F262B", ["ChromeBg"] = "#29323A", ["ChromeText"] = "#EEF0EE", ["ChromeMuted"] = "#A9B3BA",
            ["ChromeBorder"] = "#4A5761", ["HeaderRule"] = "#FDD541",
            ["PanelBg"] = "#232B31", ["CardBg"] = "#29323A", ["FieldBg"] = "#313C45",
            ["Border"] = "#3E4A54", ["BorderStrong"] = "#52606B",
            ["Text"] = "#EEF0EE", ["Muted"] = "#A9B3BA", ["Faint"] = "#8F9AA2", ["Mono"] = "#D3D9DD",
            ["ChipBg"] = "#2B343B", ["ChipBorder"] = "#3E4A54", ["SelectedBg"] = "#3A3A2C", ["SelectedBorder"] = "#FDD541",
            ["DropMappedBg"] = "#313C45", ["DropEmptyBg"] = "#20282D", ["DropEmptyBorder"] = "#5A6772",
            ["DropOverBg"] = "#3A3A2C", ["DropOverBorder"] = "#FDD541",
            ["TabBorder"] = "#4A5761", ["TabText"] = "#D3D9DD", ["TabOnBg"] = "#FDD541", ["TabOnText"] = "#1F262B",
            ["Side1"] = "#FDD541", ["Side1Ring"] = "#FDD541", ["Side2"] = "#9FB3C2",
            ["PrimaryBg"] = "#FDD541", ["PrimaryText"] = "#1F262B", ["PrimaryBorder"] = "#FDD541",
            ["NoteBg"] = "#2C3740", ["NoteBorder"] = "#465460", ["NoteText"] = "#DCE2E5",
            ["Unmapped"] = "#F2A07B", ["RowLine"] = "#323C44",
        };

        private static readonly Dictionary<string, string> Light = new Dictionary<string, string>
        {
            ["WindowBg"] = "#F5F4EF", ["ChromeBg"] = "#FFFFFF", ["ChromeText"] = "#2E373E", ["ChromeMuted"] = "#5A6770",
            ["ChromeBorder"] = "#C9CFD3", ["HeaderRule"] = "#FDD541",
            ["PanelBg"] = "#EFEEE8", ["CardBg"] = "#FFFFFF", ["FieldBg"] = "#FFFFFF",
            ["Border"] = "#D9DCDD", ["BorderStrong"] = "#B3BABF",
            ["Text"] = "#2E373E", ["Muted"] = "#5A6770", ["Faint"] = "#5E6B74", ["Mono"] = "#384650",
            ["ChipBg"] = "#FFFFFF", ["ChipBorder"] = "#D9DCDD", ["SelectedBg"] = "#FFF1BF", ["SelectedBorder"] = "#C99A12",
            ["DropMappedBg"] = "#F7F6F1", ["DropEmptyBg"] = "#FFFFFF", ["DropEmptyBorder"] = "#AEB6BB",
            ["DropOverBg"] = "#FFF1BF", ["DropOverBorder"] = "#44525C",
            ["TabBorder"] = "#C9CFD3", ["TabText"] = "#384650", ["TabOnBg"] = "#44525C", ["TabOnText"] = "#FFFFFF",
            ["Side1"] = "#FDD541", ["Side1Ring"] = "#B8900F", ["Side2"] = "#44525C",
            ["PrimaryBg"] = "#FDD541", ["PrimaryText"] = "#2E373E", ["PrimaryBorder"] = "#E0B21F",
            ["NoteBg"] = "#FFF6D6", ["NoteBorder"] = "#EBCF6A", ["NoteText"] = "#4A3E12",
            ["Unmapped"] = "#B4532A", ["RowLine"] = "#ECEBE6",
        };

        public static void Apply(FrameworkElement root, bool dark)
        {
            var dictionary = new ResourceDictionary();
            foreach (KeyValuePair<string, string> entry in dark ? Dark : Light)
            {
                var brush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(entry.Value));
                brush.Freeze();
                dictionary["Spds." + entry.Key] = brush;
            }
            dictionary[typeof(Button)] = FlatButtonStyle();
            root.Resources.MergedDictionaries.Clear();
            root.Resources.MergedDictionaries.Add(dictionary);
        }

        public static void Bind(FrameworkElement element, DependencyProperty property, string key)
        {
            element.SetResourceReference(property, "Spds." + key);
        }

        // Flat button: background, border and padding come from the button itself; no system chrome.
        private static Style FlatButtonStyle()
        {
            var border = new FrameworkElementFactory(typeof(Border), "bd");
            border.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(Control.BackgroundProperty));
            border.SetValue(Border.BorderBrushProperty, new TemplateBindingExtension(Control.BorderBrushProperty));
            border.SetValue(Border.BorderThicknessProperty, new TemplateBindingExtension(Control.BorderThicknessProperty));
            border.SetValue(Border.PaddingProperty, new TemplateBindingExtension(Control.PaddingProperty));
            border.SetValue(Border.CornerRadiusProperty, new CornerRadius(4));
            border.SetValue(UIElement.SnapsToDevicePixelsProperty, true);

            var content = new FrameworkElementFactory(typeof(ContentPresenter));
            content.SetValue(FrameworkElement.HorizontalAlignmentProperty, new TemplateBindingExtension(Control.HorizontalContentAlignmentProperty));
            content.SetValue(FrameworkElement.VerticalAlignmentProperty, new TemplateBindingExtension(Control.VerticalContentAlignmentProperty));
            border.AppendChild(content);

            var template = new ControlTemplate(typeof(Button)) { VisualTree = border };
            var hover = new Trigger { Property = UIElement.IsMouseOverProperty, Value = true };
            hover.Setters.Add(new Setter(UIElement.OpacityProperty, 0.85, "bd"));
            template.Triggers.Add(hover);
            var disabled = new Trigger { Property = UIElement.IsEnabledProperty, Value = false };
            disabled.Setters.Add(new Setter(UIElement.OpacityProperty, 0.45, "bd"));
            template.Triggers.Add(disabled);

            var style = new Style(typeof(Button));
            style.Setters.Add(new Setter(Control.TemplateProperty, template));
            style.Setters.Add(new Setter(Control.BorderThicknessProperty, new Thickness(1)));
            style.Setters.Add(new Setter(Control.PaddingProperty, new Thickness(12, 0, 12, 0)));
            style.Setters.Add(new Setter(Control.HorizontalContentAlignmentProperty, HorizontalAlignment.Center));
            style.Setters.Add(new Setter(Control.VerticalContentAlignmentProperty, VerticalAlignment.Center));
            style.Setters.Add(new Setter(FrameworkElement.CursorProperty, System.Windows.Input.Cursors.Hand));
            return style;
        }
    }
}
