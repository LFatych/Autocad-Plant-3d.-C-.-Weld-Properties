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
        private readonly HashSet<string> _weldProps;
        // Weld property rows the user added that have no source yet (target -> side). Not saved: an empty row writes nothing.
        private readonly Dictionary<string, int> _emptyRows = new Dictionary<string, int>(StringComparer.Ordinal);
        private MappingProfile _profile;
        private string _activeTab = AllTab;
        private string _selectedSource;
        private string _searchText = "";
        private string _profileName;
        private string _hintText;
        private string _noteText;
        private bool _hasMissing;

        public event Action<MappingWindowResult> CloseRequested;

        // weldProperties: all properties of the weld classes in Project Setup (empty when they could not be read).
        public MappingViewModel(WeldPropSettings settings, List<SourceProperty> sources, List<string> weldProperties, string projectName)
        {
            _settings = settings;
            _allSources = sources;
            _weldProps = new HashSet<string>(weldProperties, StringComparer.Ordinal);
            Side1 = new SideViewModel(1, RefreshOptions);
            Side2 = new SideViewModel(2, RefreshOptions);
            _profile = settings.GetActiveProfile();
            Subtitle = "SPDS Plant tools · Project: " + projectName + " · classes read from Project Setup";

            NewProfileCommand = new RelayCommand(NewProfile);
            DeleteProfileCommand = new RelayCommand(DeleteProfile, () => _settings.Profiles.Count > 1);
            ResetCommand = new RelayCommand(() => { _profile.Mappings = MappingProfile.CreateDefault().Mappings; _emptyRows.Clear(); RefreshMapping(); });
            ClearAllCommand = new RelayCommand(ClearAll);
            RemoveMissingCommand = new RelayCommand(RemoveMissing);
            SaveCommand = new RelayCommand(() => CloseRequested?.Invoke(MappingWindowResult.Save));
            SaveAndUpdateCommand = new RelayCommand(() => CloseRequested?.Invoke(MappingWindowResult.SaveAndUpdate));
            SaveAndNumberCommand = new RelayCommand(() => CloseRequested?.Invoke(MappingWindowResult.SaveAndNumber));
            CancelCommand = new RelayCommand(() => CloseRequested?.Invoke(MappingWindowResult.Cancel));

            RefreshAll();
        }

        // ---------- bindable state ----------

        public string Subtitle { get; }
        public ObservableCollection<PillItem> Tabs { get; } = new ObservableCollection<PillItem>();
        public ObservableCollection<PillItem> Profiles { get; } = new ObservableCollection<PillItem>();
        public ObservableCollection<SourceItem> Sources { get; } = new ObservableCollection<SourceItem>();
        public SideViewModel Side1 { get; }
        public SideViewModel Side2 { get; }
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

        // Project setting "autoUpdate": new welds get their properties automatically (saved with Save).
        public bool AutoUpdate
        {
            get => _settings.AutoUpdate;
            set { _settings.AutoUpdate = value; OnPropertyChanged(); }
        }

        // Project setting "numbering.auto": new welds get the number of their group, or the next free one.
        public bool AutoNumber
        {
            get => _settings.Numbering.Auto;
            set { _settings.Numbering.Auto = value; OnPropertyChanged(); }
        }

        // Start numbers of the weld types (TextBoxes; WPF rejects non-numbers before they reach the setter).
        public int ButtweldStart
        {
            get => _settings.Numbering.ButtweldStart;
            set { _settings.Numbering.ButtweldStart = Math.Max(0, value); OnPropertyChanged(); }
        }

        public int TapStart
        {
            get => _settings.Numbering.TapStart;
            set { _settings.Numbering.TapStart = Math.Max(0, value); OnPropertyChanged(); }
        }

        public int SocketweldStart
        {
            get => _settings.Numbering.SocketweldStart;
            set { _settings.Numbering.SocketweldStart = Math.Max(0, value); OnPropertyChanged(); }
        }

        // Some rows name a weld property that the weld classes in Project Setup don't have.
        public bool HasMissing
        {
            get => _hasMissing;
            private set => SetField(ref _hasMissing, value);
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
        public ICommand RemoveMissingCommand { get; }
        public ICommand SaveCommand { get; }
        public ICommand SaveAndUpdateCommand { get; }
        public ICommand SaveAndNumberCommand { get; }
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
                    _emptyRows.Clear();
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

        private bool IsMissing(string target)
        {
            return _weldProps.Count > 0 && !_weldProps.Contains(target);
        }

        private bool IsUsed(string target)
        {
            return _profile.Find(target) != null || _emptyRows.ContainsKey(target);
        }

        private void RefreshMapping()
        {
            foreach (SideViewModel side in new[] { Side1, Side2 })
            {
                SideViewModel sv = side;
                IEnumerable<string> targets = _profile.Mappings.Where(m => m.Side == sv.Side).Select(m => m.Target)
                    .Concat(_emptyRows.Where(e => e.Value == sv.Side).Select(e => e.Key))
                    .Distinct()
                    .OrderBy(t => t, StringComparer.OrdinalIgnoreCase);
                sv.Rows.Clear();
                foreach (string target in targets)
                {
                    string t = target;
                    sv.Rows.Add(new TargetRow(
                        t,
                        _profile.GetSource(t),
                        IsMissing(t),
                        dropped => Assign(t, sv.Side, dropped),
                        () => { if (_selectedSource != null) Assign(t, sv.Side, _selectedSource); },
                        () => Clear(t, sv.Side)));
                }
                RefreshOptions(sv);
            }

            Summary.Clear();
            foreach (PropertyMapping m in _profile.Mappings.OrderBy(m => m.Side).ThenBy(m => m.Target, StringComparer.OrdinalIgnoreCase))
                Summary.Add(new SummaryLine(m.Target, m.Side, m.Source, IsMissing(m.Target)));

            int missing = Side1.Rows.Concat(Side2.Rows).Count(r => r.MissingInSetup);
            HasMissing = missing > 0;
            var notes = new List<string>();
            if (_emptyRows.Count > 0) notes.Add(_emptyRows.Count + " weld field(s) not mapped; they keep their current value.");
            if (missing > 0)
                notes.Add(missing + " weld propert" + (missing == 1 ? "y (red) is" : "ies (red) are") + " not in the weld classes of Project Setup "
                    + "and will be skipped. Add them in Project Setup, or remove them and use \"+ Add weld property\".");
            if (_weldProps.Count == 0) notes.Add("Weld classes could not be read from Project Setup; showing the profile's fields only.");
            if (Side1.Rows.Count + Side2.Rows.Count == 0) notes.Add("No weld fields yet: use \"+ Add weld property\" on a side.");
            if (notes.Count == 0) notes.Add("All weld fields are mapped.");
            NoteText = string.Join("\n", notes);

            HintText = _selectedSource != null
                ? "Now click a weld field to assign \"" + _selectedSource + "\". Click the property again to cancel."
                : "Drag a property onto a weld field, or click a property and then a field.";
            OnPropertyChanged(nameof(MirrorSides));
        }

        // Weld properties from Project Setup that are not on either side yet, filtered by the side's search text.
        private void RefreshOptions(SideViewModel side)
        {
            string query = (side.PickerSearch ?? "").Trim();
            side.Options.Clear();
            foreach (string prop in _weldProps.OrderBy(p => p, StringComparer.OrdinalIgnoreCase))
            {
                if (IsUsed(prop) || prop == "WeldNumber") continue;
                if (query.Length > 0 && prop.IndexOf(query, StringComparison.OrdinalIgnoreCase) < 0) continue;
                string p = prop;
                side.Options.Add(new PickItem(p, () => AddRow(p, side.Side)));
            }
            side.HasNoOptions = side.Options.Count == 0;
        }

        // ---------- actions ----------

        // With "same mapping for both sides" the counterpart on the other side ("Port1_X" <-> "Port2_X") gets the same source.
        private void Assign(string target, int side, string source)
        {
            if (string.IsNullOrEmpty(source)) return;
            AssignOne(target, side, source);
            string other = _profile.MirrorSides ? WeldSides.Counterpart(target, _weldProps) : null;
            if (other != null) AssignOne(other, SideOfExisting(other) ?? 3 - side, source);
            _selectedSource = null;
            RefreshSources();
            RefreshMapping();
        }

        private void AssignOne(string target, int side, string source)
        {
            _profile.Assign(target, side, source);
            _emptyRows.Remove(target);
        }

        // The side a target is already shown on, or null.
        private int? SideOfExisting(string target)
        {
            PropertyMapping mapping = _profile.Find(target);
            if (mapping != null) return mapping.Side;
            return _emptyRows.TryGetValue(target, out int side) ? side : (int?)null;
        }

        // × on a row: a mapped row is emptied (kept for dropping another property), an empty row is removed.
        private void Clear(string target, int side)
        {
            if (_profile.Find(target) == null)
            {
                _emptyRows.Remove(target);
            }
            else
            {
                _profile.Unassign(target);
                _emptyRows[target] = side;
                string other = _profile.MirrorSides ? WeldSides.Counterpart(target, _weldProps) : null;
                PropertyMapping otherMapping = other == null ? null : _profile.Find(other);
                if (otherMapping != null)
                {
                    _profile.Unassign(other);
                    _emptyRows[other] = otherMapping.Side;
                }
            }
            RefreshMapping();
        }

        // Adds an empty row for a weld property; with mirroring also its counterpart on the other side.
        private void AddRow(string target, int side)
        {
            if (!IsUsed(target)) _emptyRows[target] = side;
            string other = _profile.MirrorSides ? WeldSides.Counterpart(target, _weldProps) : null;
            if (other != null && !IsUsed(other)) _emptyRows[other] = 3 - side;
            Side1.ClosePicker();
            Side2.ClosePicker();
            RefreshMapping();
        }

        private void ClearAll()
        {
            foreach (PropertyMapping m in _profile.Mappings) _emptyRows[m.Target] = m.Side;
            _profile.Mappings.Clear();
            RefreshMapping();
        }

        private void RemoveMissing()
        {
            _profile.Mappings.RemoveAll(m => IsMissing(m.Target));
            foreach (string target in _emptyRows.Keys.Where(IsMissing).ToList()) _emptyRows.Remove(target);
            RefreshMapping();
        }

        private void NewProfile()
        {
            int n = _settings.Profiles.Count + 1;
            while (_settings.Profiles.Any(p => p.Name == "Profile " + n)) n++;
            _profile = _profile.Clone("Profile " + n);
            _settings.Profiles.Add(_profile);
            _settings.ActiveProfile = _profile.Name;
            _emptyRows.Clear();
            RefreshAll();
        }

        private void DeleteProfile()
        {
            if (_settings.Profiles.Count <= 1) return;
            _settings.Profiles.Remove(_profile);
            _profile = _settings.Profiles[0];
            _settings.ActiveProfile = _profile.Name;
            _emptyRows.Clear();
            RefreshAll();
        }
    }

    // How the mapping window was closed. Every result except Cancel saves the settings first.
    public enum MappingWindowResult
    {
        Cancel,
        Save,
        SaveAndUpdate,   // then fill the mapped properties of all welds
        SaveAndNumber    // then fill the properties and renumber all welds (SetWeldNumber)
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
            ClearCommand = new RelayCommand(clear);
            ClearToolTip = source != null ? "Clear this field" : "Remove this row";
        }

        public string Target { get; }
        public string Source { get; }
        public bool MissingInSetup { get; }
        public ICommand DropCommand { get; }
        public ICommand ActivateCommand { get; }
        public ICommand ClearCommand { get; }
        public string ClearToolTip { get; }
    }

    public sealed class SummaryLine
    {
        public SummaryLine(string target, int side, string source, bool missingInSetup)
        {
            Target = target;
            Source = "← " + source + " (side " + side + ")";
            MissingInSetup = missingInSetup;
        }

        public string Target { get; }
        public string Source { get; }
        public bool MissingInSetup { get; }
    }

    // One side card (side 1 = larger part, side 2 = the other): its rows and the "+ Add weld property" picker.
    public sealed class SideViewModel : ObservableObject
    {
        private readonly Action<SideViewModel> _refreshOptions;
        private bool _isPickerOpen;
        private string _pickerSearch = "";
        private bool _hasNoOptions;

        public SideViewModel(int side, Action<SideViewModel> refreshOptions)
        {
            Side = side;
            _refreshOptions = refreshOptions;
            Title = side == 1 ? "Side 1 · larger part" : "Side 2 · smaller part";
            Description = side == 1 ? "port with bigger OD, then wall thickness" : "the other connected part";
            TogglePickerCommand = new RelayCommand(() => IsPickerOpen = !IsPickerOpen);
        }

        public int Side { get; }
        public string Title { get; }
        public string Description { get; }
        public ObservableCollection<TargetRow> Rows { get; } = new ObservableCollection<TargetRow>();
        public ObservableCollection<PickItem> Options { get; } = new ObservableCollection<PickItem>();
        public ICommand TogglePickerCommand { get; }

        public bool IsPickerOpen
        {
            get => _isPickerOpen;
            set { if (SetField(ref _isPickerOpen, value) && value) _refreshOptions(this); }
        }

        public string PickerSearch
        {
            get => _pickerSearch;
            set { if (SetField(ref _pickerSearch, value ?? "")) _refreshOptions(this); }
        }

        public bool HasNoOptions
        {
            get => _hasNoOptions;
            set => SetField(ref _hasNoOptions, value);
        }

        public void ClosePicker()
        {
            IsPickerOpen = false;
            PickerSearch = "";
        }
    }

    // A weld property offered by the "+ Add weld property" picker.
    public sealed class PickItem
    {
        public PickItem(string name, Action add)
        {
            Name = name;
            AddCommand = new RelayCommand(add);
        }

        public string Name { get; }
        public ICommand AddCommand { get; }
    }
}
