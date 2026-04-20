using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Win32;

namespace DllInjector
{
    public partial class MainWindow : Window
    {
        private readonly ObservableCollection<DllItem> _selectedDlls = new();
        private List<ProcessItem> _allProcesses = new();
        private Timer? _refreshTimer;
        private Timer? _saveDebounceTimer;
        private readonly string _configPath;
        private string _lastSelectedProcessName = "";
        private volatile bool _autoInjectPending = false;
        private bool _lastStatusOk = true;
        private bool _suppressSelectionClear = false;

        private readonly ConcurrentDictionary<string, ImageSource?> _iconCache = new(StringComparer.OrdinalIgnoreCase);

        public MainWindow()
        {
            InitializeComponent();
            _configPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "Downloads", "FluentInjector_Config.json");
            Loaded += MainWindow_Loaded;
            Closed += (_, _) =>
            {
                _refreshTimer?.Dispose();
                _saveDebounceTimer?.Dispose();
                Microsoft.Win32.SystemEvents.UserPreferenceChanged -= OnUserPreferenceChanged;
            };
            DllListControl.ItemsSource = _selectedDlls;
            _selectedDlls.CollectionChanged += (_, _) => ScheduleSave();
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            ApplyAcrylicBlur();
            ThemeComboBox.SelectionChanged -= ThemeComboBox_SelectionChanged;
            ThemeComboBox.SelectedIndex = 0;
            ThemeComboBox.SelectionChanged += ThemeComboBox_SelectionChanged;
            LoadConfig();
            if (ThemeComboBox.SelectedItem is System.Windows.Controls.ComboBoxItem item)
                ApplyTheme(item.Tag?.ToString() ?? "Dark");
            Task.Run(LoadProcesses);
            _refreshTimer = new Timer(_ => Task.Run(LoadProcesses), null, 2000, 2000);
        }

        // ── Process list ───────────────────────────────────────────────────────

        private void LoadProcesses()
        {
            try
            {
                var snapshot = Process.GetProcesses();
                var processes = new List<ProcessItem>();
                foreach (var p in snapshot.OrderBy(p => p.ProcessName))
                {
                    try
                    {
                        if (string.IsNullOrEmpty(p.MainWindowTitle)) continue;
                        processes.Add(new ProcessItem
                        {
                            DisplayName = $"{p.ProcessName}.exe",
                            Id = p.Id,
                            ProcessIcon = GetProcessIconCached(p)
                        });
                    }
                    catch { }
                    finally
                    {
                        p.Dispose();
                    }
                }

                Dispatcher.Invoke(() =>
                {
                    var currentId = (ProcessListBox.SelectedItem as ProcessItem)?.Id;

                    // Only update if the process set changed (by ID)
                    var newIds = processes.Select(p => p.Id).ToHashSet();
                    var oldIds = _allProcesses.Select(p => p.Id).ToHashSet();
                    if (!newIds.SetEquals(oldIds))
                    {
                        _allProcesses = processes;

                        _suppressSelectionClear = true;
                        ApplySearchFilter();
                        _suppressSelectionClear = false;

                        if (currentId.HasValue)
                        {
                            var match = _allProcesses.FirstOrDefault(p => p.Id == currentId.Value);
                            if (match != null)
                            {
                                ProcessListBox.SelectedItem = match;
                            }
                            else
                            {
                                ProcessListBox.SelectedItem = null;
                                if (!string.IsNullOrEmpty(_lastSelectedProcessName))
                                    SetStatus($"Status: Waiting for {_lastSelectedProcessName} (not running)", true);
                            }
                        }

                        if (ProcessListBox.SelectedItem == null && !string.IsNullOrEmpty(_lastSelectedProcessName))
                        {
                            var match = _allProcesses.FirstOrDefault(p => p.DisplayName == _lastSelectedProcessName);
                            if (match != null)
                                ProcessListBox.SelectedItem = match;
                        }
                    }

                    if (_autoInjectPending && ProcessListBox.SelectedItem is ProcessItem autoTarget)
                    {
                        _autoInjectPending = false;
                        _ = RunAutoInjectAsync(autoTarget);
                    }
                });
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() => SetStatus($"Error loading processes: {ex.Message}", false));
            }
        }

        private ImageSource? GetProcessIconCached(Process process)
        {
            try
            {
                string? path = process.MainModule?.FileName;
                if (string.IsNullOrEmpty(path)) return null;

                if (_iconCache.TryGetValue(path, out var cached))
                    return cached;

                ImageSource? icon = LoadIconFromPath(path);
                _iconCache[path] = icon;
                return icon;
            }
            catch
            {
                return null;
            }
        }

        private static ImageSource? LoadIconFromPath(string path)
        {
            try
            {
                if (!File.Exists(path)) return null;
                using var icon = System.Drawing.Icon.ExtractAssociatedIcon(path);
                if (icon == null) return null;
                var bitmap = Imaging.CreateBitmapSourceFromHIcon(
                    icon.Handle, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                bitmap.Freeze();
                return bitmap;
            }
            catch
            {
                return null;
            }
        }

        private void ApplySearchFilter()
        {
            var query = SearchBox.Text.Trim().ToLowerInvariant();
            ProcessListBox.ItemsSource = string.IsNullOrEmpty(query)
                ? _allProcesses
                : _allProcesses.Where(p => p.DisplayName.ToLowerInvariant().Contains(query)).ToList();
        }

        private void SearchBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            ApplySearchFilter();
        }

        // ── Custom process / list toggle ───────────────────────────────────────

        private bool _processListVisible = true;

        private void ToggleProcessList_Click(object sender, RoutedEventArgs e)
        {
            _processListVisible = !_processListVisible;
            ProcessListBox.Visibility = _processListVisible ? Visibility.Visible : Visibility.Collapsed;

            // Swap eye / eye-off icon to reflect current state
            var btn = (System.Windows.Controls.Button)sender;
            var eyeOn  = (System.Windows.Controls.Viewbox)btn.Template.FindName("EyeIcon",    btn);
            var eyeOff = (System.Windows.Controls.Viewbox)btn.Template.FindName("EyeOffIcon", btn);
            if (eyeOn != null)  eyeOn.Visibility  = _processListVisible ? Visibility.Visible   : Visibility.Collapsed;
            if (eyeOff != null) eyeOff.Visibility = _processListVisible ? Visibility.Collapsed : Visibility.Visible;

            ScheduleSave();
        }

        private void AddCustomProcess_Click(object sender, RoutedEventArgs e)
        {
            bool show = CustomProcessRow.Visibility != Visibility.Visible;
            CustomProcessRow.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
            if (show)
            {
                CustomProcessBox.Clear();
                CustomProcessBox.Focus();
            }
        }

        private void ConfirmCustomProcess_Click(object sender, RoutedEventArgs e)
        {
            ApplyCustomProcess();
        }

        private void CustomProcessBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                ApplyCustomProcess();
            else if (e.Key == Key.Escape)
                CustomProcessRow.Visibility = Visibility.Collapsed;
        }

        private void ApplyCustomProcess()
        {
            var name = CustomProcessBox.Text.Trim();
            if (string.IsNullOrEmpty(name)) return;

            if (!name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                name += ".exe";

            var match = _allProcesses.FirstOrDefault(
                p => p.DisplayName.Equals(name, StringComparison.OrdinalIgnoreCase));

            if (match != null)
            {
                ProcessListBox.SelectedItem = match;
                ProcessListBox.ScrollIntoView(match);
                SetStatus($"Status: Selected {match.DisplayName} (PID {match.Id})", true);
            }
            else
            {
                _lastSelectedProcessName = name;
                SetStatus($"Status: Waiting for {name} (not running)", true);
                ScheduleSave();
            }

            CustomProcessRow.Visibility = Visibility.Collapsed;
        }

        private void ProcessListBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (ProcessListBox.SelectedItem is ProcessItem p)
            {
                _lastSelectedProcessName = p.DisplayName;
                SetStatus($"Status: Selected {p.DisplayName} (PID {p.Id})", true);
                ScheduleSave();
            }
            else if (e.RemovedItems.Count > 0 && !_suppressSelectionClear)
            {
                _lastSelectedProcessName = "";
            }
        }

        // ── Title bar ──────────────────────────────────────────────────────────

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                DragMove();
        }

        private void MinimizeButton_Click(object sender, RoutedEventArgs e) => WindowState = WindowState.Minimized;
        private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();

        // ── DLL management ─────────────────────────────────────────────────────

        private void DropZone_Click(object sender, MouseButtonEventArgs e) => BrowseForDll();

        private void BrowseForDll()
        {
            var dlg = new OpenFileDialog
            {
                Filter = "DLL Files (*.dll)|*.dll|All Files (*.*)|*.*",
                Title = "Select DLL to Inject",
                Multiselect = true
            };
            if (dlg.ShowDialog() == true)
                AddDlls(dlg.FileNames);
        }

        private void DropZone_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
            DropZoneBorder.BorderThickness = new Thickness(2);
            DropZoneBorder.BorderBrush = (SolidColorBrush)Application.Current.Resources["AccentBrush"];
            e.Handled = true;
        }

        private void DropZone_DragLeave(object sender, DragEventArgs e)
        {
            DropZoneBorder.BorderThickness = new Thickness(0);
            DropZoneBorder.BorderBrush = null;
        }

        private void DropZone_Drop(object sender, DragEventArgs e)
        {
            DropZone_DragLeave(sender, e);
            if (e.Data.GetData(DataFormats.FileDrop) is string[] files)
                AddDlls(files.Where(f => f.EndsWith(".dll", StringComparison.OrdinalIgnoreCase)));
        }

        private void AddDlls(IEnumerable<string> paths)
        {
            foreach (var path in paths)
            {
                if (!File.Exists(path)) continue;
                if (_selectedDlls.Any(d => d.FullPath.Equals(path, StringComparison.OrdinalIgnoreCase))) continue;
                _selectedDlls.Add(new DllItem { FileName = Path.GetFileName(path), FullPath = path });
            }
            SetStatus($"Status: {_selectedDlls.Count} DLL(s) selected", true);
        }

        private void RemoveDll_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button btn && btn.Tag is string path)
            {
                var item = _selectedDlls.FirstOrDefault(d => d.FullPath == path);
                if (item != null) _selectedDlls.Remove(item);
            }
        }

        // ── Auto Inject ────────────────────────────────────────────────────────

        private void AutoInjectCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            bool enabled = AutoInjectCheckBox.IsChecked == true;
            if (enabled)
            {
                if (string.IsNullOrEmpty(_lastSelectedProcessName))
                {
                    SetStatus("Status: Select a process before enabling Auto Inject", false);
                    AutoInjectCheckBox.IsChecked = false;
                    return;
                }
                if (_selectedDlls.Count == 0)
                {
                    SetStatus("Status: Add at least one DLL before enabling Auto Inject", false);
                    AutoInjectCheckBox.IsChecked = false;
                    return;
                }
                _autoInjectPending = true;
                SetStatus($"Status: Auto Inject armed — waiting for {_lastSelectedProcessName}", true);
            }
            else
            {
                _autoInjectPending = false;
                SetStatus("Status: Auto Inject disabled", true);
            }
        }

        // ── Injection ──────────────────────────────────────────────────────────

        private async void InjectButton_Click(object sender, RoutedEventArgs e)
        {
            InjectBtn.IsEnabled = false;
            try
            {
                if (ProcessListBox.SelectedItem is not ProcessItem target)
                {
                    SetStatus(!string.IsNullOrEmpty(_lastSelectedProcessName)
                        ? $"Status: Process '{_lastSelectedProcessName}' not found - please start it first"
                        : "Status: Please select a target process", false);
                    return;
                }

                if (_selectedDlls.Count == 0)
                {
                    SetStatus("Status: Please add at least one DLL", false);
                    return;
                }

                var dllSnapshot = _selectedDlls.ToList();
                InjectSpinner.Visibility = Visibility.Visible;
                SetStatus("Status: Preparing injection...", true);

                var results = await Task.Run(() =>
                {
                    var log = new List<string>();
                    foreach (var dll in dllSnapshot)
                    {
                        bool ok = StandardInject(target.Id, dll.FullPath);
                        log.Add($"{dll.FileName}: {(ok ? "OK" : "FAILED")}");
                    }
                    return log;
                });

                bool allOk = results.All(r => r.EndsWith("OK"));
                SetStatus($"Status: {string.Join("  |  ", results)}", allOk);

                await Task.Delay(2500);
            }
            catch (Exception ex)
            {
                SetStatus($"Status: Injection error — {ex.Message}", false);
            }
            finally
            {
                InjectSpinner.Visibility = Visibility.Collapsed;
                InjectBtn.IsEnabled = true;
            }
        }

        private async Task RunAutoInjectAsync(ProcessItem target)
        {
            var dllSnapshot = _selectedDlls.ToList();
            if (dllSnapshot.Count == 0) return;

            // Verify the process is still alive before attempting injection
            try
            {
                using var p = Process.GetProcessById(target.Id);
                if (p.HasExited)
                {
                    _autoInjectPending = true;
                    return;
                }
            }
            catch
            {
                _autoInjectPending = true;
                return;
            }

            InjectBtn.IsEnabled = false;
            InjectSpinner.Visibility = Visibility.Visible;
            SetStatus($"Status: Auto Inject — injecting into {target.DisplayName}...", true);

            try
            {
                var results = await Task.Run(() =>
                {
                    var log = new List<string>();
                    foreach (var dll in dllSnapshot)
                    {
                        bool ok = StandardInject(target.Id, dll.FullPath);
                        log.Add($"{dll.FileName}: {(ok ? "OK" : "FAILED")}");
                    }
                    return log;
                });

                bool allOk = results.All(r => r.EndsWith("OK"));
                SetStatus($"Status: Auto Inject — {string.Join("  |  ", results)}", allOk);
                await Task.Delay(2500);

                // FIX #5: re-arm regardless of success/failure so auto-inject keeps working
                // after a failed attempt (e.g. wrong bitness, missing dependency).
                // Always wait for exit first so we don't inject twice into the same instance.
                if (AutoInjectCheckBox.IsChecked == true)
                {
                    SetStatus($"Status: Auto Inject — waiting for {_lastSelectedProcessName} to restart", true);
                    _ = WaitForProcessExitThenRearmAsync(target.Id);
                }
            }
            catch (Exception ex)
            {
                SetStatus($"Status: Auto Inject error — {ex.Message}", false);
                if (AutoInjectCheckBox.IsChecked == true)
                    _ = WaitForProcessExitThenRearmAsync(target.Id);
            }
            finally
            {
                InjectSpinner.Visibility = Visibility.Collapsed;
                InjectBtn.IsEnabled = true;
            }
        }

        private async Task WaitForProcessExitThenRearmAsync(int processId)
        {
            try
            {
                using var p = Process.GetProcessById(processId);
                await Task.Run(() =>
                {
                    try { p.WaitForExit(); }
                    catch { }
                });
            }
            catch { }

            if (AutoInjectCheckBox.IsChecked == true)
            {
                _autoInjectPending = true;
                Dispatcher.Invoke(() =>
                    SetStatus($"Status: Auto Inject armed — waiting for {_lastSelectedProcessName}", true));
            }
        }

        private bool StandardInject(int processId, string dllPath)
        {
            IntPtr hProcess = OpenProcess(
                ProcessAccessFlags.CreateThread |
                ProcessAccessFlags.QueryInformation |
                ProcessAccessFlags.VirtualMemoryOperation |
                ProcessAccessFlags.VirtualMemoryWrite |
                ProcessAccessFlags.VirtualMemoryRead,
                false, processId);

            if (hProcess == IntPtr.Zero) return false;

            try
            {
                IntPtr loadLibW = GetProcAddress(GetModuleHandle("kernel32.dll"), "LoadLibraryW");
                if (loadLibW == IntPtr.Zero) return false;

                byte[] pathBytes = Encoding.Unicode.GetBytes(dllPath + "\0");
                IntPtr mem = VirtualAllocEx(hProcess, IntPtr.Zero, (uint)pathBytes.Length,
                    AllocationType.Commit | AllocationType.Reserve, MemoryProtection.ReadWrite);
                if (mem == IntPtr.Zero) return false;

                try
                {
                    bool written = WriteProcessMemory(hProcess, mem, pathBytes, (uint)pathBytes.Length, out _);
                    if (!written) return false;

                    IntPtr hThread = CreateRemoteThread(hProcess, IntPtr.Zero, 0, loadLibW, mem, 0, IntPtr.Zero);
                    if (hThread == IntPtr.Zero) return false;

                    uint waitResult = WaitForSingleObject(hThread, 5000);
                    if (waitResult != 0)
                    {
                        CloseHandle(hThread);
                        return false;
                    }

                    GetExitCodeThread(hThread, out uint exitCode);
                    CloseHandle(hThread);
                    VirtualFreeEx(hProcess, mem, 0, AllocationType.Release);
                    return exitCode != 0;
                }
                catch
                {
                    VirtualFreeEx(hProcess, mem, 0, AllocationType.Release);
                    throw;
                }
            }
            finally
            {
                CloseHandle(hProcess);
            }
        }

        // ── Config ─────────────────────────────────────────────────────────────

        private void ScheduleSave()
        {
            _saveDebounceTimer?.Dispose();
            _saveDebounceTimer = new Timer(_ => Dispatcher.Invoke(SaveConfig), null, 300, Timeout.Infinite);
        }

        private void SaveConfig()
        {
            try
            {
                if (ThemeComboBox == null || ProcessListBox == null) return;

                var selectedTheme = (ThemeComboBox.SelectedItem as System.Windows.Controls.ComboBoxItem)?.Tag?.ToString() ?? "Dark";
                var selectedProcess = ProcessListBox.SelectedItem as ProcessItem;
                var config = new
                {
                    ProcessName         = selectedProcess?.DisplayName ?? _lastSelectedProcessName,
                    ProcessId           = selectedProcess?.Id ?? 0,
                    DllPaths            = _selectedDlls.Select(d => d.FullPath).ToList(),
                    Theme               = selectedTheme,
                    AutoInject          = AutoInjectCheckBox.IsChecked == true,
                    ProcessListVisible  = _processListVisible,
                    SavedAt             = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                };

                string json = System.Text.Json.JsonSerializer.Serialize(config,
                    new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_configPath, json);
            }
            catch { /* Silent fail on auto-save */ }
        }

        private void LoadConfig()
        {
            try
            {
                if (!File.Exists(_configPath)) return;

                string json = File.ReadAllText(_configPath);
                var config = System.Text.Json.JsonSerializer.Deserialize<System.Text.Json.JsonElement>(json);

                if (config.TryGetProperty("DllPaths", out var dllPaths))
                {
                    foreach (var path in dllPaths.EnumerateArray())
                    {
                        string dllPath = path.GetString() ?? "";
                        if (File.Exists(dllPath))
                            _selectedDlls.Add(new DllItem { FileName = Path.GetFileName(dllPath), FullPath = dllPath });
                    }
                }

                if (config.TryGetProperty("Theme", out var theme))
                {
                    string themeName = theme.GetString() ?? "Dark";
                    var item = ThemeComboBox.Items.OfType<System.Windows.Controls.ComboBoxItem>()
                        .FirstOrDefault(i => i.Tag?.ToString() == themeName);
                    if (item != null) ThemeComboBox.SelectedItem = item;
                }

                if (config.TryGetProperty("ProcessName", out var procName))
                {
                    string targetProcess = procName.GetString() ?? "";
                    if (!string.IsNullOrEmpty(targetProcess))
                    {
                        _lastSelectedProcessName = targetProcess;
                        SetStatus($"Status: Waiting for {targetProcess} (not running)", true);
                    }
                }

                if (config.TryGetProperty("AutoInject", out var autoInject) && autoInject.GetBoolean())
                {
                    AutoInjectCheckBox.IsChecked = true;
                    _autoInjectPending = true;
                }

                if (config.TryGetProperty("ProcessListVisible", out var listVisible))
                {
                    _processListVisible = listVisible.GetBoolean();
                    ProcessListBox.Visibility = _processListVisible ? Visibility.Visible : Visibility.Collapsed;

                    Dispatcher.InvokeAsync(() =>
                    {
                        var eyeOn  = (System.Windows.Controls.Viewbox)ToggleListBtn.Template.FindName("EyeIcon",    ToggleListBtn);
                        var eyeOff = (System.Windows.Controls.Viewbox)ToggleListBtn.Template.FindName("EyeOffIcon", ToggleListBtn);
                        if (eyeOn  != null) eyeOn.Visibility  = _processListVisible ? Visibility.Visible   : Visibility.Collapsed;
                        if (eyeOff != null) eyeOff.Visibility = _processListVisible ? Visibility.Collapsed : Visibility.Visible;
                    }, System.Windows.Threading.DispatcherPriority.Loaded);
                }
            }
            catch { /* Silent fail on load */ }
        }

        // ── Status ─────────────────────────────────────────────────────────────

        private void SetStatus(string message, bool ok)
        {
            _lastStatusOk = ok;
            StatusTextBlock.Text = message;
            StatusTextBlock.Foreground = ok
                ? (SolidColorBrush)Application.Current.Resources["StatusOkText"]
                : (SolidColorBrush)Application.Current.Resources["ErrorText"];
        }

        // ── Theme ──────────────────────────────────────────────────────────────

        private void ThemeComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (ThemeComboBox.SelectedItem is not System.Windows.Controls.ComboBoxItem item) return;
            var tag = item.Tag?.ToString() ?? "Dark";
            if (tag == "Auto")
            {
                // Subscribe to system theme changes
                Microsoft.Win32.SystemEvents.UserPreferenceChanged -= OnUserPreferenceChanged;
                Microsoft.Win32.SystemEvents.UserPreferenceChanged += OnUserPreferenceChanged;
                ApplyTheme("Auto");
            }
            else
            {
                Microsoft.Win32.SystemEvents.UserPreferenceChanged -= OnUserPreferenceChanged;
                ApplyTheme(tag);
            }
            ScheduleSave();
        }

        private void OnUserPreferenceChanged(object sender, Microsoft.Win32.UserPreferenceChangedEventArgs e)
        {
            if (e.Category == Microsoft.Win32.UserPreferenceCategory.General)
                Dispatcher.Invoke(() => ApplyTheme("Auto"));
        }

        private static bool IsSystemDarkMode()
        {
            try
            {
                using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
                return key?.GetValue("AppsUseLightTheme") is int i && i == 0;
            }
            catch { return true; }
        }

        private void ApplyTheme(string theme)
        {
            if (theme == "Auto")
                theme = IsSystemDarkMode() ? "Dark" : "Light";

            switch (theme)
            {
                case "Dark":
                    SetThemeColors("#AA1E1E1E", "#551E1E1E", "#F0F0F0", "#AAAAAA", "#AAAAAA",
                                   "#15000000", "#207B68EE", "#107B68EE", "#25000000", "#45000000", "#7B68EE");
                    break;
                case "Darker":
                    SetThemeColors("#CC0F0F0F", "#550F0F0F", "#F0F0F0", "#AAAAAA", "#AAAAAA",
                                   "#15000000", "#207B68EE", "#107B68EE", "#25000000", "#45000000", "#7B68EE");
                    break;
                case "Light":
                    SetThemeColors("#88C8C8C8", "#55FFFFFF", "#1A1A2E", "#666666", "#444444",
                                   "#18000000", "#207B68EE", "#107B68EE", "#15000000", "#30000000", "#7B68EE");
                    break;
            }
        }

        private void SetThemeColors(
            string rootBg, string panelBg, string text, string mutedText, string statusOk,
            string hover, string selection, string inactiveSel, string comboHover, string comboSel, string accent)
        {
            var res = Application.Current.Resources;

            SolidColorBrush Brush(string hex) =>
                new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));

            res["RootBackground"]      = Brush(rootBg);
            res["PanelBackground"]     = Brush(panelBg);
            res["PrimaryText"]         = Brush(text);
            res["MutedText"]           = Brush(mutedText);
            res["StatusOkText"]        = Brush(statusOk);
            res["HoverBackground"]     = Brush(hover);
            res["SelectionBackground"] = Brush(selection);
            res["InactiveSelection"]   = Brush(inactiveSel);
            res["ComboHover"]          = Brush(comboHover);
            res["ComboSelected"]       = Brush(comboSel);
            res["AccentBrush"]         = Brush(accent);

            // ListBox system brushes can't use DynamicResource — update them directly on the ListBox
            var selColor   = (Color)ColorConverter.ConvertFromString(selection);
            var inactColor = (Color)ColorConverter.ConvertFromString(inactiveSel);
            var textColor  = (Color)ColorConverter.ConvertFromString(text);
            ProcessListBox.Resources[SystemColors.HighlightBrushKey]                      = new SolidColorBrush(selColor);
            ProcessListBox.Resources[SystemColors.HighlightTextBrushKey]                  = new SolidColorBrush(textColor);
            ProcessListBox.Resources[SystemColors.InactiveSelectionHighlightBrushKey]     = new SolidColorBrush(inactColor);
            ProcessListBox.Resources[SystemColors.InactiveSelectionHighlightTextBrushKey] = new SolidColorBrush(textColor);

            SetStatus(StatusTextBlock.Text, _lastStatusOk);
        }

        // ── Acrylic blur ───────────────────────────────────────────────────────

        private void ApplyAcrylicBlur()
        {
            var helper = new WindowInteropHelper(this);
            if (helper.Handle == IntPtr.Zero) return;

            int backdropType = 3;
            DwmSetWindowAttribute(helper.Handle, 38, ref backdropType, Marshal.SizeOf<int>());

            int cornerPref = 2;
            DwmSetWindowAttribute(helper.Handle, 33, ref cornerPref, Marshal.SizeOf<int>());
        }

        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int dwAttribute, ref int pvAttribute, int cbAttribute);

        // ── Windows API ────────────────────────────────────────────────────────

        [DllImport("kernel32.dll")]
        private static extern IntPtr OpenProcess(ProcessAccessFlags access, bool inherit, int pid);

        [DllImport("kernel32.dll", CharSet = CharSet.Ansi)]
        private static extern IntPtr GetProcAddress(IntPtr hModule, string name);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr GetModuleHandle(string name);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr VirtualAllocEx(IntPtr hProcess, IntPtr lpAddress, uint size,
            AllocationType type, MemoryProtection protect);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool VirtualFreeEx(IntPtr hProcess, IntPtr lpAddress, uint size, AllocationType type);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool WriteProcessMemory(IntPtr hProcess, IntPtr lpBase, byte[] buf, uint size, out UIntPtr written);

        [DllImport("kernel32.dll")]
        private static extern IntPtr CreateRemoteThread(IntPtr hProcess, IntPtr attr, uint stackSize,
            IntPtr startAddr, IntPtr param, uint flags, IntPtr threadId);

        [DllImport("kernel32.dll")]
        private static extern uint WaitForSingleObject(IntPtr handle, uint ms);

        [DllImport("kernel32.dll")]
        private static extern bool GetExitCodeThread(IntPtr hThread, out uint lpExitCode);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr handle);
    }

    // ── Data models ────────────────────────────────────────────────────────────

    public class ProcessItem
    {
        public string DisplayName { get; set; } = string.Empty;
        public int Id { get; set; }
        public ImageSource? ProcessIcon { get; set; }
    }

    public class DllItem
    {
        public string FileName { get; set; } = string.Empty;
        public string FullPath { get; set; } = string.Empty;
    }

    // ── Win32 enums ────────────────────────────────────────────────────────────

    [Flags]
    internal enum ProcessAccessFlags : uint
    {
        CreateThread           = 0x00000002,
        VirtualMemoryOperation = 0x00000008,
        VirtualMemoryRead      = 0x00000010,
        VirtualMemoryWrite     = 0x00000020,
        QueryInformation       = 0x00000400,
    }

    [Flags]
    internal enum AllocationType
    {
        Commit  = 0x1000,
        Reserve = 0x2000,
        Release = 0x8000,
    }

    [Flags]
    internal enum MemoryProtection
    {
        ReadWrite        = 0x04,
        ExecuteReadWrite = 0x40,
    }
}
