using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using WeldPropUtils.Schema;
using WeldPropUtils.Settings;

namespace WeldPropUtils.UI
{
    // View model of the Weld Property Mapping window. Edits the WeldPropSettings passed in;
    // the command that opened the window saves them when CloseRequested reports saved = true.
    public sealed class MappingViewModel : ObservableObject
    {
        public const string AllTab = "All";

        private readonly WeldPropSettings _settings;
        private readonly List<SourceProperty> _allSources;
        private readonly HashSet<string> _schemaTargets;
        private MappingProfile _profile;
        private string _activeTab = AllTab;
        private string _selectedSource;
        private string _searchText = "";
        private string _profileName;
        private string _hintText;
        private string _noteText;

        // (saved, updateWelds)
        public event Action<bool, bool> CloseRequested;

        public MappingViewModel(WeldPropSettings settings, List<SourceProperty> sources, List<string> schemaTargets, string projectName)
        {
            _settings = settings;
            _allSources = sources;
            _schemaTargets = new HashSet<string>(schemaTargets, StringComparer.Ordinal);
            _profile = settings.GetActiveProfile();
            Subtitle = "SPDS Plant tools · Project: " + projectName + " · classes read from Project Setup";

            NewProfileCommand = new RelayCommand(NewProfile);
            DeleteProfileCommand = new RelayCommand(DeleteProfile, () => _settings.Profiles.Count > 1);
            ResetCommand = new RelayCommand(() => { _profile.Mappings = MappingProfile.CreateDefault().Mappings; RefreshMapping(); });
            ClearAllCommand = new RelayCommand(() => { _profile.Mappings.Clear(); RefreshMapping(); });
            SaveCommand = new RelayCommand(() => CloseRequested?.Invoke(true, false));
            SaveAndUpdateCommand = new RelayCommand(() => CloseRequested?.Invoke(true, true));
            CancelCommand = new RelayCommand(() => CloseRequested?.Invoke(false, false));

            RefreshAll();
        }

        // ---------- bindable state ----------

        public string Subtitle { get; }
        public ObservableCollection<PillItem> Tabs { get; } = new ObservableCollection<PillItem>();
        public ObservableCollection<PillItem> Profiles { get; } = new ObservableCollection<PillItem>();
        public ObservableCollection<SourceItem> Sources { get; } = new ObservableCollection<SourceItem>();
        public ObservableCollection<TargetRow> Side1Rows { get; } = new ObservableCollection<TargetRow>();
        public ObservableCollection<TargetRow> Side2Rows { get; } = new ObservableCollection<TargetRow>();
        public ObservableCollection<SummaryLine> Summary { get; } = new ObservableCollection<SummaryLine>();
        public bool HasNoSources => _allSources.Count == 0;

        public string SearchText
        {
            get => _searchText;
            set { if (SetField(ref _searchText, value ?? "")) RefreshSources(); }
        }

        public bool MirrorSides
        {
            get => _profile.MirrorSides;
            set { _profile.MirrorSides = value; OnPropertyChanged(); }
        }

        // Bound with UpdateSourceTrigger=LostFocus: renames the active profile when the name is valid and unique.
        public string ProfileName
        {
            get => _profileName;
            set
            {
                string name = (value ?? "").Trim();
                if (name.Length > 0 && name != _profile.Name && !_settings.Profiles.Any(p => p.Name == name))
                {
                    _profile.Name = name;
                    _settings.ActiveProfile = name;
                    RefreshProfiles();
                }
                else
                {
                    _profileName = _profile.Name;
                    OnPropertyChanged();
                }
            }
        }

        public string HintText
        {
            get => _hintText;
            private set => SetField(ref _hintText, value);
        }

        public string NoteText
        {
            get => _noteText;
            private set => SetField(ref _noteText, value);
        }

        public ICommand NewProfileCommand { get; }
        public ICommand DeleteProfileCommand { get; }
        public ICommand ResetCommand { get; }
        public ICommand ClearAllCommand { get; }
        public ICommand SaveCommand { get; }
        public ICommand SaveAndUpdateCommand { get; }
        public ICommand CancelCommand { get; }

        // ---------- refresh ----------

        private void RefreshAll()
        {
            RefreshProfiles();
            RefreshTabs();
            RefreshSources();
            RefreshMapping();
        }

        private void RefreshProfiles()
        {
            Profiles.Clear();
            foreach (MappingProfile profile in _settings.Profiles)
            {
                MappingProfile p = profile;
                Profiles.Add(new PillItem(p.Name, ReferenceEquals(p, _profile), () =>
                {
                    _profile = p;
                    _settings.ActiveProfile = p.Name;
                    _selectedSource = null;
                    RefreshAll();
                }));
            }
            _profileName = _profile.Name;
            OnPropertyChanged(nameof(ProfileName));
            OnPropertyChanged(nameof(MirrorSides));
        }

        private void RefreshTabs()
        {
            Tabs.Clear();
            foreach (string tab in new[] { AllTab }.Concat(ProjectSchema.SourceGroups.Select(g => g.Key)))
            {
                string t = tab;
                Tabs.Add(new PillItem(t, t == _activeTab, () => { _activeTab = t; RefreshTabs(); RefreshSources(); }));
            }
        }

        private void RefreshSources()
        {
            Sources.Clear();
            string query = _searchText.Trim();
            foreach (SourceProperty source in _allSources)
            {
                if (_activeTab != AllTab && !source.Groups.Contains(_activeTab)) continue;
                if (query.Length > 0 && source.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) < 0) continue;
                string name = source.Name;
                string scope = source.Groups.Count == ProjectSchema.SourceGroups.Count ? "all classes" : string.Join(", ", source.Groups);
                Sources.Add(new SourceItem(name, scope, name == _selectedSource, () =>
                {
                    _selectedSource = _selectedSource == name ? null : name;
                    RefreshSources();
                    RefreshMapping();
                }));
            }
        }

        private void RefreshMapping()
        {
            var targets = new SortedSet<string>(_schemaTargets, StringComparer.OrdinalIgnoreCase);
            foreach (MappingProfile profile in _settings.Profiles)
                foreach (PropertyMapping m in profile.Mappings)
                    targets.Add(m.Target);

            Side1Rows.Clear();
            Side2Rows.Clear();
            foreach (string target in targets)
            {
                string t = target;
                var row = new TargetRow(
                    t,
                    _profile.GetSource(t),
                    _schemaTargets.Count > 0 && !_schemaTargets.Contains(t),
                    dropped => Assign(t, dropped),
                    () => { if (_selectedSource != null) Assign(t, _selectedSource); },
                    () => { _profile.Unassign(t, _profile.MirrorSides); RefreshMapping(); });
                (MappingProfile.SideOf(t) == 2 ? Side2Rows : Side1Rows).Add(row);
            }

            Summary.Clear();
            foreach (PropertyMapping m in _profile.Mappings.OrderBy(m => m.Side).ThenBy(m => m.Target, StringComparer.OrdinalIgnoreCase))
                Summary.Add(new SummaryLine(m.Target, m.Source));

            int unmapped = Side1Rows.Count + Side2Rows.Count - _profile.Mappings.Count;
            int missing = _schemaTargets.Count == 0 ? 0 : _profile.Mappings.Count(m => !_schemaTargets.Contains(m.Target));
            var notes = new List<string>
            {
                unmapped <= 0 ? "All weld fields are mapped." : unmapped + " weld field(s) not mapped; they keep their current value."
            };
            if (missing > 0) notes.Add(missing + " mapped weld propert" + (missing == 1 ? "y is" : "ies are") + " missing in Project Setup; add them to the weld class or writing fails.");
            if (_schemaTargets.Count == 0) notes.Add("Weld classes could not be read from Project Setup; showing the profile's fields only.");
            NoteText = string.Join("\n", notes);

            HintText = _selectedSource != null
                ? "Now click a weld field to assign \"" + _selectedSource + "\". Click the property again to cancel."
                : "Drag a property onto a weld field, or click a property and then a field.";
            OnPropertyChanged(nameof(MirrorSides));
        }

        // ---------- actions ----------

        private void Assign(string target, string source)
        {
            if (string.IsNullOrEmpty(source)) return;
            _profile.Assign(target, source, _profile.MirrorSides);
            _selectedSource = null;
            RefreshSources();
            RefreshMapping();
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
    }

    // A pill button (class tab or profile): label, active state, click action.
    public sealed class PillItem
    {
        public PillItem(string name, bool isActive, Action select)
        {
            Name = name;
            IsActive = isActive;
            SelectCommand = new RelayCommand(select);
        }

        public string Name { get; }
        public bool IsActive { get; }
        public ICommand SelectCommand { get; }
    }

    public sealed class SourceItem
    {
        public SourceItem(string name, string scope, bool isSelected, Action toggle)
        {
            Name = name;
            Scope = scope;
            IsSelected = isSelected;
            SelectCommand = new RelayCommand(toggle);
        }

        public string Name { get; }
        public string Scope { get; }
        public bool IsSelected { get; }
        public ICommand SelectCommand { get; }
    }

    public sealed class TargetRow
    {
        public TargetRow(string target, string source, bool missingInSetup, Action<string> drop, Action activate, Action clear)
        {
            Target = target;
            Source = source;
            MissingInSetup = missingInSetup;
            DropCommand = new RelayCommand(p => drop(p as string));
            ActivateCommand = new RelayCommand(activate);
            ClearCommand = new RelayCommand(clear, () => source != null);
        }

        public string Target { get; }
        public string Source { get; }
        public bool MissingInSetup { get; }
        public ICommand DropCommand { get; }
        public ICommand ActivateCommand { get; }
        public ICommand ClearCommand { get; }
    }

    public sealed class SummaryLine
    {
        public SummaryLine(string target, string source)
        {
            Target = target;
            Source = "← " + source;
        }

        public string Target { get; }
        public string Source { get; }
    }
}
