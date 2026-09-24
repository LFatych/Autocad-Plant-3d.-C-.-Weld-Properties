using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using WeldPropUtils.Schema;
using WeldPropUtils.Settings;

namespace WeldPropUtils.UI
{
    // Weld Property Mapping window (approved design: .claude/skills/spds-wpf-ui).
    // Built in code rather than XAML so it compiles and is checked like the rest of the plugin.
    // Edits the WeldPropSettings passed in; the caller saves them when Saved is true.
    public sealed class MappingWindow : Window
    {
        private const string AllTab = "All";

        private readonly WeldPropSettings _settings;
        private readonly List<SourceProperty> _sources;
        private readonly HashSet<string> _schemaTargets;
        private readonly UserPreferences _prefs;
        private MappingProfile _profile;
        private bool _dark;
        private string _tab = AllTab;
        private string _selected;
        private Point _dragStart;

        private TextBox _search;
        private WrapPanel _tabs;
        private StackPanel _sourceList;
        private StackPanel _side1Rows;
        private StackPanel _side2Rows;
        private StackPanel _summary;
        private WrapPanel _profiles;
        private TextBox _profileName;
        private Button _deleteProfile;
        private CheckBox _mirror;
        private Button _themeButton;
        private TextBlock _hint;
        private TextBlock _note;

        public bool Saved { get; private set; }
        public bool UpdateWeldsRequested { get; private set; }

        public MappingWindow(WeldPropSettings settings, List<SourceProperty> sources, List<string> schemaTargets,
            string projectName, bool dark, UserPreferences prefs)
        {
            _settings = settings;
            _sources = sources;
            _schemaTargets = new HashSet<string>(schemaTargets, StringComparer.Ordinal);
            _prefs = prefs;
            _dark = dark;
            _profile = settings.GetActiveProfile();

            Title = "Weld Property Mapping";
            Width = 1280;
            Height = 800;
            MinWidth = 1000;
            MinHeight = 600;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            FontFamily = new FontFamily(SpdsTheme.BodyFont);
            FontSize = 13;
            SpdsTheme.Apply(this, _dark);
            SpdsTheme.Bind(this, BackgroundProperty, "WindowBg");
            SpdsTheme.Bind(this, ForegroundProperty, "Text");

            Content = BuildLayout(projectName);
            RefreshAll();
        }

        // ---------- layout ----------

        private UIElement BuildLayout(string projectName)
        {
            var root = new Grid();
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            UIElement header = BuildHeader(projectName);
            UIElement body = BuildBody();
            UIElement footer = BuildFooter();
            Grid.SetRow(header, 0);
            Grid.SetRow(body, 1);
            Grid.SetRow(footer, 2);
            root.Children.Add(header);
            root.Children.Add(body);
            root.Children.Add(footer);
            return root;
        }

        private UIElement BuildHeader(string projectName)
        {
            var bar = new Border { Padding = new Thickness(20, 10, 20, 10), BorderThickness = new Thickness(0, 0, 0, 2) };
            SpdsTheme.Bind(bar, Border.BackgroundProperty, "ChromeBg");
            SpdsTheme.Bind(bar, Border.BorderBrushProperty, "HeaderRule");

            var dock = new DockPanel { LastChildFill = true };
            _themeButton = MakeButton("", ButtonKind.Chrome);
            _themeButton.Click += (s, e) => ToggleTheme();
            DockPanel.SetDock(_themeButton, Dock.Right);
            dock.Children.Add(_themeButton);

            var titles = new StackPanel { VerticalAlignment = VerticalAlignment.Center };
            titles.Children.Add(Text("Weld Property Mapping", "ChromeText", 18, heading: true));
            titles.Children.Add(Text("SPDS Plant tools · Project: " + projectName + " · classes read from Project Setup", "ChromeMuted", 12));
            dock.Children.Add(titles);
            bar.Child = dock;
            return bar;
        }

        private UIElement BuildBody()
        {
            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(300) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(320) });

            UIElement left = BuildSourcePanel();
            UIElement center = BuildTargetPanel();
            UIElement right = BuildSummaryPanel();
            Grid.SetColumn(left, 0);
            Grid.SetColumn(center, 1);
            Grid.SetColumn(right, 2);
            grid.Children.Add(left);
            grid.Children.Add(center);
            grid.Children.Add(right);
            return grid;
        }

        private UIElement BuildSourcePanel()
        {
            var panel = new Border { Padding = new Thickness(16), BorderThickness = new Thickness(0, 0, 1, 0) };
            SpdsTheme.Bind(panel, Border.BackgroundProperty, "PanelBg");
            SpdsTheme.Bind(panel, Border.BorderBrushProperty, "Border");

            var dock = new DockPanel();
            var top = new StackPanel();
            top.Children.Add(Text("Connected part properties", "Text", 14, heading: true, bottom: 10));
            _search = new TextBox { Height = 32, Padding = new Thickness(8, 6, 8, 0), Margin = new Thickness(0, 0, 0, 10) };
            StyleField(_search);
            AutomationProperties.SetName(_search, "Search properties");
            _search.TextChanged += (s, e) => RenderSources();
            top.Children.Add(_search);
            _tabs = new WrapPanel { Margin = new Thickness(0, 0, 0, 10) };
            top.Children.Add(_tabs);
            DockPanel.SetDock(top, Dock.Top);
            dock.Children.Add(top);

            TextBlock info = Text("Properties come from the pipe, fitting, flange, valve and nozzle classes in Project Setup.", "Faint", 12, top: 10);
            info.TextWrapping = TextWrapping.Wrap;
            DockPanel.SetDock(info, Dock.Bottom);
            dock.Children.Add(info);

            _sourceList = new StackPanel();
            dock.Children.Add(new ScrollViewer { Content = _sourceList, VerticalScrollBarVisibility = ScrollBarVisibility.Auto });
            panel.Child = dock;
            return panel;
        }

        private UIElement BuildTargetPanel()
        {
            var stack = new StackPanel { Margin = new Thickness(20, 16, 20, 16) };

            var toolbar = new DockPanel { Margin = new Thickness(0, 0, 0, 14) };
            Button clearAll = MakeButton("Clear all", ButtonKind.Normal);
            clearAll.Click += (s, e) => { _profile.Mappings.Clear(); RenderMapping(); };
            Button reset = MakeButton("Reset to default", ButtonKind.Normal);
            reset.Margin = new Thickness(8, 0, 8, 0);
            reset.Click += (s, e) => { _profile.Mappings = MappingProfile.CreateDefault().Mappings; RenderMapping(); };
            DockPanel.SetDock(clearAll, Dock.Right);
            DockPanel.SetDock(reset, Dock.Right);
            toolbar.Children.Add(clearAll);
            toolbar.Children.Add(reset);
            _mirror = new CheckBox { Content = "Same mapping for both sides", VerticalAlignment = VerticalAlignment.Center };
            SpdsTheme.Bind(_mirror, ForegroundProperty, "Text");
            _mirror.Click += (s, e) => _profile.MirrorSides = _mirror.IsChecked == true;
            toolbar.Children.Add(_mirror);
            stack.Children.Add(toolbar);

            stack.Children.Add(SideCard("Side 1 · larger part", "port with bigger OD, then wall thickness", "Side1", out _side1Rows));
            stack.Children.Add(SideCard("Side 2 · smaller part", "the other connected part", "Side2", out _side2Rows));

            TextBlock numbering = Text("WeldNumber is set by the numbering rules (SetWeldNumber), not mapped here.", "Muted", 12, top: 4);
            stack.Children.Add(numbering);
            return new ScrollViewer { Content = stack, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
        }

        private Border SideCard(string title, string subtitle, string colorKey, out StackPanel rows)
        {
            var card = Card();
            card.Margin = new Thickness(0, 0, 0, 14);
            var stack = new StackPanel();
            var head = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 10) };
            var dot = new Border { Width = 10, Height = 10, CornerRadius = new CornerRadius(5), BorderThickness = new Thickness(1), Margin = new Thickness(0, 0, 10, 0), VerticalAlignment = VerticalAlignment.Center };
            SpdsTheme.Bind(dot, Border.BackgroundProperty, colorKey);
            SpdsTheme.Bind(dot, Border.BorderBrushProperty, colorKey == "Side1" ? "Side1Ring" : colorKey);
            head.Children.Add(dot);
            head.Children.Add(Text(title, "Text", 14, heading: true, right: 10));
            head.Children.Add(Text(subtitle, "Muted", 12));
            stack.Children.Add(head);
            rows = new StackPanel();
            stack.Children.Add(rows);
            card.Child = stack;
            return card;
        }

        private UIElement BuildSummaryPanel()
        {
            var panel = new Border { Padding = new Thickness(16), BorderThickness = new Thickness(1, 0, 0, 0) };
            SpdsTheme.Bind(panel, Border.BackgroundProperty, "PanelBg");
            SpdsTheme.Bind(panel, Border.BorderBrushProperty, "Border");
            var dock = new DockPanel();

            var top = new StackPanel();
            top.Children.Add(Text("Profile", "Text", 14, heading: true, bottom: 8));
            _profiles = new WrapPanel { Margin = new Thickness(0, 0, 0, 8) };
            top.Children.Add(_profiles);
            _profileName = new TextBox { Height = 30, Padding = new Thickness(8, 5, 8, 0), Margin = new Thickness(0, 0, 0, 8) };
            StyleField(_profileName);
            AutomationProperties.SetName(_profileName, "Profile name");
            _profileName.LostFocus += (s, e) => RenameProfile();
            _profileName.KeyDown += (s, e) => { if (e.Key == Key.Enter) RenameProfile(); };
            top.Children.Add(_profileName);
            var profileButtons = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 16) };
            Button newProfile = MakeButton("New profile", ButtonKind.Normal);
            newProfile.Click += (s, e) => NewProfile();
            _deleteProfile = MakeButton("Delete", ButtonKind.Normal);
            _deleteProfile.Margin = new Thickness(8, 0, 0, 0);
            _deleteProfile.Click += (s, e) => DeleteProfile();
            profileButtons.Children.Add(newProfile);
            profileButtons.Children.Add(_deleteProfile);
            top.Children.Add(profileButtons);
            top.Children.Add(Text("Mapping summary", "Text", 14, heading: true, bottom: 8));
            DockPanel.SetDock(top, Dock.Top);
            dock.Children.Add(top);

            var noteBox = new Border { Padding = new Thickness(10), CornerRadius = new CornerRadius(6), BorderThickness = new Thickness(1), Margin = new Thickness(0, 12, 0, 0) };
            SpdsTheme.Bind(noteBox, Border.BackgroundProperty, "NoteBg");
            SpdsTheme.Bind(noteBox, Border.BorderBrushProperty, "NoteBorder");
            _note = Text("", "NoteText", 12);
            _note.TextWrapping = TextWrapping.Wrap;
            noteBox.Child = _note;
            DockPanel.SetDock(noteBox, Dock.Bottom);
            dock.Children.Add(noteBox);

            _summary = new StackPanel();
            var summaryCard = new Border { BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(6), Child = _summary };
            SpdsTheme.Bind(summaryCard, Border.BackgroundProperty, "CardBg");
            SpdsTheme.Bind(summaryCard, Border.BorderBrushProperty, "Border");
            dock.Children.Add(new ScrollViewer { Content = summaryCard, VerticalScrollBarVisibility = ScrollBarVisibility.Auto });
            panel.Child = dock;
            return panel;
        }

        private UIElement BuildFooter()
        {
            var bar = new Border { Padding = new Thickness(20, 12, 20, 12), BorderThickness = new Thickness(0, 1, 0, 0) };
            SpdsTheme.Bind(bar, Border.BackgroundProperty, "ChromeBg");
            SpdsTheme.Bind(bar, Border.BorderBrushProperty, "ChromeBorder");
            var dock = new DockPanel();

            Button update = MakeButton("Save and update all welds", ButtonKind.Primary);
            update.Margin = new Thickness(8, 0, 0, 0);
            update.Click += (s, e) => { UpdateWeldsRequested = true; Saved = true; Close(); };
            Button save = MakeButton("Save", ButtonKind.Chrome);
            save.Margin = new Thickness(8, 0, 0, 0);
            save.Click += (s, e) => { Saved = true; Close(); };
            Button cancel = MakeButton("Cancel", ButtonKind.Chrome);
            cancel.IsCancel = true;
            cancel.Click += (s, e) => Close();
            foreach (Button b in new[] { update, save, cancel })
            {
                DockPanel.SetDock(b, Dock.Right);
                dock.Children.Add(b);
            }
            _hint = Text("", "ChromeMuted", 12);
            _hint.VerticalAlignment = VerticalAlignment.Center;
            dock.Children.Add(_hint);
            bar.Child = dock;
            return bar;
        }

        // ---------- rendering ----------

        private void RefreshAll()
        {
            _themeButton.Content = _dark ? "Light theme" : "Dark theme";
            RenderProfiles();
            RenderTabs();
            RenderSources();
            RenderMapping();
        }

        private void RenderMapping()
        {
            _mirror.IsChecked = _profile.MirrorSides;
            RenderRows();
            RenderSummary();
            _hint.Text = _selected != null
                ? "Now click a weld field to assign \"" + _selected + "\". Click the property again to cancel."
                : "Drag a property onto a weld field, or click a property and then a field.";
        }

        private void RenderProfiles()
        {
            _profiles.Children.Clear();
            foreach (MappingProfile profile in _settings.Profiles)
            {
                MappingProfile p = profile;
                Button pill = MakePill(p.Name, ReferenceEquals(p, _profile));
                pill.Click += (s, e) => { _profile = p; _settings.ActiveProfile = p.Name; _selected = null; RefreshAll(); };
                _profiles.Children.Add(pill);
            }
            _profileName.Text = _profile.Name;
            _deleteProfile.IsEnabled = _settings.Profiles.Count > 1;
        }

        private void RenderTabs()
        {
            _tabs.Children.Clear();
            foreach (string tab in new[] { AllTab }.Concat(ProjectSchema.SourceGroups.Select(g => g.Key)))
            {
                string t = tab;
                Button pill = MakePill(t, t == _tab);
                pill.Click += (s, e) => { _tab = t; RenderTabs(); RenderSources(); };
                _tabs.Children.Add(pill);
            }
        }

        private void RenderSources()
        {
            _sourceList.Children.Clear();
            string query = (_search.Text ?? "").Trim();
            var visible = _sources.Where(p => (_tab == AllTab || p.Groups.Contains(_tab))
                && (query.Length == 0 || p.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0)).ToList();
            if (_sources.Count == 0)
                _sourceList.Children.Add(Text("No classes could be read from Project Setup.", "Unmapped", 12));
            foreach (SourceProperty source in visible)
                _sourceList.Children.Add(SourceChip(source));
        }

        private Button SourceChip(SourceProperty source)
        {
            bool selected = source.Name == _selected;
            var content = new StackPanel();
            content.Children.Add(Text(source.Name, "Text", 13));
            string scope = source.Groups.Count == ProjectSchema.SourceGroups.Count ? "all classes" : string.Join(", ", source.Groups);
            content.Children.Add(Text(scope, "Faint", 11));

            var chip = new Button
            {
                Content = content,
                HorizontalContentAlignment = HorizontalAlignment.Stretch,
                Padding = new Thickness(10, 6, 10, 6),
                Margin = new Thickness(0, 0, 0, 6),
                MinHeight = 44,
                Cursor = Cursors.SizeAll,
            };
            SpdsTheme.Bind(chip, Control.BackgroundProperty, selected ? "SelectedBg" : "ChipBg");
            SpdsTheme.Bind(chip, Control.BorderBrushProperty, selected ? "SelectedBorder" : "ChipBorder");
            AutomationProperties.SetName(chip, source.Name + (selected ? ", selected" : ""));

            chip.Click += (s, e) =>
            {
                _selected = selected ? null : source.Name;
                RenderSources();
                RenderMapping();
            };
            chip.PreviewMouseLeftButtonDown += (s, e) => _dragStart = e.GetPosition(this);
            chip.PreviewMouseMove += (s, e) =>
            {
                if (e.LeftButton != MouseButtonState.Pressed) return;
                Vector moved = e.GetPosition(this) - _dragStart;
                if (Math.Abs(moved.X) < SystemParameters.MinimumHorizontalDragDistance
                    && Math.Abs(moved.Y) < SystemParameters.MinimumVerticalDragDistance) return;
                DragDrop.DoDragDrop(chip, new DataObject(DataFormats.StringFormat, source.Name), DragDropEffects.Copy);
            };
            return chip;
        }

        private void RenderRows()
        {
            var targets = new SortedSet<string>(_schemaTargets, StringComparer.OrdinalIgnoreCase);
            foreach (MappingProfile profile in _settings.Profiles)
                foreach (PropertyMapping m in profile.Mappings)
                    targets.Add(m.Target);

            _side1Rows.Children.Clear();
            _side2Rows.Children.Clear();
            foreach (string target in targets)
            {
                int side = MappingProfile.SideOf(target);
                (side == 2 ? _side2Rows : _side1Rows).Children.Add(TargetRow(target));
            }
        }

        private UIElement TargetRow(string target)
        {
            string source = _profile.GetSource(target);
            bool missingInSetup = _schemaTargets.Count > 0 && !_schemaTargets.Contains(target);

            var row = new Grid { Margin = new Thickness(0, 0, 0, 8) };
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(170) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(40) });

            var name = new StackPanel { VerticalAlignment = VerticalAlignment.Center };
            TextBlock label = Text(target, "Mono", 13);
            label.FontFamily = new FontFamily(SpdsTheme.MonoFont);
            name.Children.Add(label);
            if (missingInSetup) name.Children.Add(Text("not in Project Setup", "Unmapped", 11));
            row.Children.Add(name);

            var zoneContent = new StackPanel { Orientation = Orientation.Horizontal };
            if (source != null)
            {
                TextBlock mapped = Text(source, "Text", 13);
                mapped.FontWeight = FontWeights.SemiBold;
                zoneContent.Children.Add(mapped);
            }
            else
            {
                zoneContent.Children.Add(Text("Drop a property here", "Faint", 13));
            }

            var zone = new Button
            {
                Content = zoneContent,
                Height = 40,
                HorizontalContentAlignment = HorizontalAlignment.Left,
                AllowDrop = true,
            };
            SetZoneLook(zone, source != null, false);
            AutomationProperties.SetName(zone, source != null
                ? target + " mapped to " + source + ". Activate to replace with the selected property."
                : target + " not mapped. Activate to assign the selected property.");
            zone.Click += (s, e) => { if (_selected != null) Assign(target, _selected); };
            zone.DragEnter += (s, e) => { if (e.Data.GetDataPresent(DataFormats.StringFormat)) SetZoneLook(zone, source != null, true); };
            zone.DragLeave += (s, e) => SetZoneLook(zone, source != null, false);
            zone.DragOver += (s, e) =>
            {
                e.Effects = e.Data.GetDataPresent(DataFormats.StringFormat) ? DragDropEffects.Copy : DragDropEffects.None;
                e.Handled = true;
            };
            zone.Drop += (s, e) =>
            {
                if (e.Data.GetData(DataFormats.StringFormat) is string dropped) Assign(target, dropped);
            };
            Grid.SetColumn(zone, 1);
            row.Children.Add(zone);

            var clear = new Button { Content = "×", FontSize = 16, Width = 32, Height = 32, Padding = new Thickness(0), BorderThickness = new Thickness(0), Background = Brushes.Transparent, Margin = new Thickness(8, 0, 0, 0) };
            SpdsTheme.Bind(clear, ForegroundProperty, "Muted");
            AutomationProperties.SetName(clear, "Clear " + target);
            clear.IsEnabled = source != null;
            clear.Click += (s, e) => { _profile.Unassign(target, _profile.MirrorSides); RenderMapping(); };
            Grid.SetColumn(clear, 2);
            row.Children.Add(clear);
            return row;
        }

        private static void SetZoneLook(Button zone, bool mapped, bool over)
        {
            SpdsTheme.Bind(zone, Control.BackgroundProperty, over ? "DropOverBg" : mapped ? "DropMappedBg" : "DropEmptyBg");
            SpdsTheme.Bind(zone, Control.BorderBrushProperty, over ? "DropOverBorder" : mapped ? "BorderStrong" : "DropEmptyBorder");
        }

        private void RenderSummary()
        {
            _summary.Children.Clear();
            foreach (PropertyMapping m in _profile.Mappings.OrderBy(m => m.Side).ThenBy(m => m.Target, StringComparer.OrdinalIgnoreCase))
            {
                var line = new Border { Padding = new Thickness(10, 6, 10, 6), BorderThickness = new Thickness(0, 0, 0, 1) };
                SpdsTheme.Bind(line, Border.BorderBrushProperty, "RowLine");
                var grid = new Grid();
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                TextBlock target = Text(m.Target, "Mono", 12);
                target.FontFamily = new FontFamily(SpdsTheme.MonoFont);
                TextBlock source = Text("← " + m.Source, "Text", 12);
                source.TextWrapping = TextWrapping.Wrap;
                Grid.SetColumn(source, 1);
                grid.Children.Add(target);
                grid.Children.Add(source);
                line.Child = grid;
                _summary.Children.Add(line);
            }

            int total = _side1Rows.Children.Count + _side2Rows.Children.Count;
            int unmapped = total - _profile.Mappings.Count;
            int missing = _profile.Mappings.Count(m => _schemaTargets.Count > 0 && !_schemaTargets.Contains(m.Target));
            var notes = new List<string>();
            notes.Add(unmapped <= 0 ? "All weld fields are mapped." : unmapped + " weld field(s) not mapped; they keep their current value.");
            if (missing > 0) notes.Add(missing + " mapped weld propert" + (missing == 1 ? "y is" : "ies are") + " missing in Project Setup; add them to the weld class or writing fails.");
            if (_schemaTargets.Count == 0) notes.Add("Weld classes could not be read from Project Setup; showing the profile's fields only.");
            _note.Text = string.Join("\n", notes);
        }

        // ---------- actions ----------

        private void Assign(string target, string source)
        {
            _profile.Assign(target, source, _profile.MirrorSides);
            _selected = null;
            RenderSources();
            RenderMapping();
        }

        private void ToggleTheme()
        {
            _dark = !_dark;
            SpdsTheme.Apply(this, _dark);
            _prefs.Theme = _dark ? UserPreferences.ThemeDark : UserPreferences.ThemeLight;
            _prefs.Save();
            _themeButton.Content = _dark ? "Light theme" : "Dark theme";
        }

        private void NewProfile()
        {
            int n = _settings.Profiles.Count + 1;
            while (_settings.Profiles.Any(p => p.Name == "Profile " + n)) n++;
            _profile = _profile.Clone("Profile " + n);
            _settings.Profiles.Add(_profile);
            _settings.ActiveProfile = _profile.Name;
            RefreshAll();
        }

        private void DeleteProfile()
        {
            if (_settings.Profiles.Count <= 1) return;
            _settings.Profiles.Remove(_profile);
            _profile = _settings.Profiles[0];
            _settings.ActiveProfile = _profile.Name;
            RefreshAll();
        }

        private void RenameProfile()
        {
            string name = (_profileName.Text ?? "").Trim();
            if (name.Length == 0 || name == _profile.Name || _settings.Profiles.Any(p => p.Name == name))
            {
                _profileName.Text = _profile.Name;
                return;
            }
            _profile.Name = name;
            _settings.ActiveProfile = name;
            RenderProfiles();
        }

        // ---------- small builders ----------

        private enum ButtonKind { Normal, Chrome, Primary }

        private static Button MakeButton(string text, ButtonKind kind)
        {
            var button = new Button { Content = text, Height = 34, MinWidth = 80, FontSize = 13 };
            switch (kind)
            {
                case ButtonKind.Primary:
                    SpdsTheme.Bind(button, Control.BackgroundProperty, "PrimaryBg");
                    SpdsTheme.Bind(button, Control.ForegroundProperty, "PrimaryText");
                    SpdsTheme.Bind(button, Control.BorderBrushProperty, "PrimaryBorder");
                    button.FontWeight = FontWeights.Bold;
                    break;
                case ButtonKind.Chrome:
                    button.Background = Brushes.Transparent;
                    SpdsTheme.Bind(button, Control.ForegroundProperty, "ChromeText");
                    SpdsTheme.Bind(button, Control.BorderBrushProperty, "ChromeBorder");
                    break;
                default:
                    button.Background = Brushes.Transparent;
                    SpdsTheme.Bind(button, Control.ForegroundProperty, "Text");
                    SpdsTheme.Bind(button, Control.BorderBrushProperty, "BorderStrong");
                    break;
            }
            return button;
        }

        private static Button MakePill(string text, bool on)
        {
            var pill = new Button { Content = text, Height = 28, Padding = new Thickness(10, 0, 10, 0), Margin = new Thickness(0, 0, 6, 6), FontSize = 12 };
            SpdsTheme.Bind(pill, Control.BackgroundProperty, on ? "TabOnBg" : "PanelBg");
            SpdsTheme.Bind(pill, Control.ForegroundProperty, on ? "TabOnText" : "TabText");
            SpdsTheme.Bind(pill, Control.BorderBrushProperty, on ? "TabOnBg" : "TabBorder");
            return pill;
        }

        private static Border Card()
        {
            var card = new Border { Padding = new Thickness(14), CornerRadius = new CornerRadius(8), BorderThickness = new Thickness(1) };
            SpdsTheme.Bind(card, Border.BackgroundProperty, "CardBg");
            SpdsTheme.Bind(card, Border.BorderBrushProperty, "Border");
            return card;
        }

        private static void StyleField(TextBox box)
        {
            SpdsTheme.Bind(box, Control.BackgroundProperty, "FieldBg");
            SpdsTheme.Bind(box, Control.ForegroundProperty, "Text");
            SpdsTheme.Bind(box, Control.BorderBrushProperty, "BorderStrong");
            SpdsTheme.Bind(box, TextBoxBase.CaretBrushProperty, "Text");
        }

        private static TextBlock Text(string text, string colorKey, double size, bool heading = false,
            double top = 0, double right = 0, double bottom = 0)
        {
            var block = new TextBlock { Text = text, FontSize = size, Margin = new Thickness(0, top, right, bottom), VerticalAlignment = VerticalAlignment.Center };
            if (heading)
            {
                block.FontFamily = new FontFamily(SpdsTheme.HeadingFont);
                block.FontWeight = FontWeights.Bold;
            }
            SpdsTheme.Bind(block, TextBlock.ForegroundProperty, colorKey);
            return block;
        }
    }
}
