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
        private int _saveDebounceVersion = 0;
        private readonly string _configPath;
        private string _lastSelectedProcessName = "";
        private bool _autoInjectPending = false;
        private bool _lastStatusOk = true;
        private bool _suppressSelectionClear = false;

        private readonly ConcurrentDictionary<string, ImageSource?> _iconCache = new(StringComparer.OrdinalIgnoreCase);

        public MainWindow()
        {
            InitializeComponent();
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string configDir = Path.Combine(appData, "FluentInjector");
            Directory.CreateDirectory(configDir);
            _configPath = Path.Combine(configDir, "config.json");
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
                        processes.Add(new ProcessItem(
                            DisplayName: $"{p.ProcessName}.exe",
                            Id: p.Id,
                            ProcessIcon: GetProcessIconCached(p)
                        ));
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

                if (_iconCache.Count > 200)
                {
                    var oldest = _iconCache.Keys.FirstOrDefault();
                    if (oldest != null) _iconCache.TryRemove(oldest, out _);
                }

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

        private bool _processListVisible = true;

        private void ToggleProcessList_Click(object sender, RoutedEventArgs e)
        {
            _processListVisible = !_processListVisible;
            ProcessListBox.Visibility = _processListVisible ? Visibility.Visible : Visibility.Collapsed;
            SyncToggleIcon();
            ScheduleSave();
        }

        /// <summary>
        /// Syncs the eye / eye-off icon on the toggle button to match <see cref="_processListVisible"/>.
        /// Uses Template.FindName; must be called after the template is applied (i.e. after Loaded).
        /// </summary>
        private void SyncToggleIcon()
        {
            if (!IsLoaded) return;
            var eyeOn  = ToggleListBtn.Template.FindName("EyeIcon",    ToggleListBtn) as System.Windows.Controls.Viewbox;
            var eyeOff = ToggleListBtn.Template.FindName("EyeOffIcon", ToggleListBtn) as System.Windows.Controls.Viewbox;
            if (eyeOn  != null) eyeOn.Visibility  = _processListVisible ? Visibility.Visible   : Visibility.Collapsed;
            if (eyeOff != null) eyeOff.Visibility = _processListVisible ? Visibility.Collapsed : Visibility.Visible;
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
                ScheduleSave();
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

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
                DragMove();
        }

        private void MinimizeButton_Click(object sender, RoutedEventArgs e) => WindowState = WindowState.Minimized;
        private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();

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
                _selectedDlls.Add(new DllItem(FileName: Path.GetFileName(path), FullPath: path));
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
                        string? err = StandardInject(target.Id, dll.FullPath);
                        log.Add(err == null ? $"{dll.FileName}: OK" : $"{dll.FileName}: {err}");
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
                        string? err = StandardInject(target.Id, dll.FullPath);
                        log.Add(err == null ? $"{dll.FileName}: OK" : $"{dll.FileName}: {err}");
                    }
                    return log;
                });

                bool allOk = results.All(r => r.EndsWith("OK"));
                SetStatus($"Status: Auto Inject — {string.Join("  |  ", results)}", allOk);
                await Task.Delay(2500);

                if (AutoInjectCheckBox.IsChecked == true)
                {
                    SetStatus($"Status: Auto Inject — waiting for {_lastSelectedProcessName} to restart", true);
                    _ = WaitForProcessExitThenRearmAsync(target.Id);
                }
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                {
                    SetStatus($"Status: Auto Inject error — {ex.Message}", false);
                    if (AutoInjectCheckBox.IsChecked == true)
                        _ = WaitForProcessExitThenRearmAsync(target.Id);
                });
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
                // If the process exits between GetProcessById and WaitForExit the exception
                // is swallowed intentionally — the process is gone, so fall through to re-arm.
                await Task.Run(() =>
                {
                    try { p.WaitForExit(); }
                    catch { }
                });
            }
            catch { }

            Dispatcher.Invoke(() =>
            {
                if (AutoInjectCheckBox.IsChecked == true)
                {
                    _autoInjectPending = true;
                    SetStatus($"Status: Auto Inject armed — waiting for {_lastSelectedProcessName}", true);
                }
            });
        }

        // Returns null on success, or a human-readable error string on failure.
        private string? StandardInject(int processId, string dllPath)
        {
            if (!Path.IsPathRooted(dllPath) || !File.Exists(dllPath))
                return "DLL not found";

            // Bitness check — injecting a mismatched DLL causes a silent LoadLibrary failure
            bool dllIs32Bit = IsDll32Bit(dllPath);

            IntPtr hProcess = OpenProcess(
                ProcessAccessFlags.CreateThread |
                ProcessAccessFlags.QueryInformation |
                ProcessAccessFlags.VirtualMemoryOperation |
                ProcessAccessFlags.VirtualMemoryWrite |
                ProcessAccessFlags.VirtualMemoryRead,
                false, processId);

            if (hProcess == IntPtr.Zero)
                return $"OpenProcess failed (error {Marshal.GetLastWin32Error()})";

            try
            {
                IsWow64Process(hProcess, out bool processIs32Bit);
                if (dllIs32Bit != processIs32Bit)
                    return dllIs32Bit
                        ? "Bitness mismatch: 32-bit DLL \u2192 64-bit process"
                        : "Bitness mismatch: 64-bit DLL \u2192 32-bit process";

                IntPtr loadLibW = GetProcAddress(GetModuleHandle("kernel32.dll"), "LoadLibraryW");
                if (loadLibW == IntPtr.Zero) return "GetProcAddress(LoadLibraryW) failed";

                byte[] pathBytes = Encoding.Unicode.GetBytes(dllPath + "\0");
                IntPtr mem = VirtualAllocEx(hProcess, IntPtr.Zero, (uint)pathBytes.Length,
                    AllocationType.Commit | AllocationType.Reserve, MemoryProtection.ReadWrite);
                if (mem == IntPtr.Zero) return "VirtualAllocEx failed";

                try
                {
                    if (!WriteProcessMemory(hProcess, mem, pathBytes, (uint)pathBytes.Length, out _))
                        return "WriteProcessMemory failed";

                    IntPtr hThread = CreateRemoteThread(hProcess, IntPtr.Zero, 0, loadLibW, mem, 0, IntPtr.Zero);
                    if (hThread == IntPtr.Zero) return "CreateRemoteThread failed";

                    try
                    {
                        // Wait indefinitely — freeing memory while DllMain is still running crashes the target.
                        // DllMain is expected to return quickly; INFINITE is the safe choice here.
                        const uint INFINITE = 0xFFFFFFFF;
                        WaitForSingleObject(hThread, INFINITE);

                        GetExitCodeThread(hThread, out uint exitCode);
                        return exitCode != 0 ? null : "LoadLibrary returned NULL (check DLL dependencies / path)";
                    }
                    finally
                    {
                        CloseHandle(hThread);
                    }
                }
                finally
                {
                    VirtualFreeEx(hProcess, mem, 0, AllocationType.Release);
                }
            }
            finally
            {
                CloseHandle(hProcess);
            }
        }

        /// <summary>Reads the PE Machine field to determine if a DLL targets x86 (32-bit).</summary>
        private static bool IsDll32Bit(string dllPath)
        {
            try
            {
                using var fs = new FileStream(dllPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                using var br = new BinaryReader(fs);
                // DOS header: e_magic = 0x5A4D ("MZ"), e_lfanew at offset 0x3C
                fs.Seek(0x3C, SeekOrigin.Begin);
                int peOffset = br.ReadInt32();
                fs.Seek(peOffset, SeekOrigin.Begin);
                uint peSig = br.ReadUInt32(); // "PE\0\0"
                if (peSig != 0x00004550) return false;
                ushort machine = br.ReadUInt16(); // IMAGE_FILE_HEADER.Machine
                return machine == 0x014C; // IMAGE_FILE_MACHINE_I386
            }
            catch
            {
                return false; // assume 64-bit on read failure
            }
        }

        private void ScheduleSave()
        {
            int version = Interlocked.Increment(ref _saveDebounceVersion);
            _saveDebounceTimer?.Dispose();
            _saveDebounceTimer = new Timer(_ =>
            {
                if (Volatile.Read(ref _saveDebounceVersion) == version)
                    Dispatcher.Invoke(SaveConfig);
            }, null, 300, Timeout.Infinite);
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
                            _selectedDlls.Add(new DllItem(FileName: Path.GetFileName(dllPath), FullPath: dllPath));
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
                    // Defer icon sync until after the template is applied
                    Dispatcher.InvokeAsync(SyncToggleIcon, System.Windows.Threading.DispatcherPriority.Loaded);
                }
            }
            catch (Exception ex)
            {
                SetStatus($"Status: Failed to load config — {ex.Message}", false);
            }
        }

        private void SetStatus(string message, bool ok)
        {
            _lastStatusOk = ok;
            StatusTextBlock.Text = message;
            StatusTextBlock.Foreground = ok
                ? (SolidColorBrush)Application.Current.Resources["StatusOkText"]
                : (SolidColorBrush)Application.Current.Resources["ErrorText"];
        }

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

        private record ThemeColors(
            string RootBg,     string PanelBg,
            string Text,       string MutedText,   string StatusOk,
            string Hover,      string Selection,   string InactiveSel,
            string ComboHover, string ComboSel,
            string Accent);

        private static readonly ThemeColors DarkTheme = new(
            RootBg:      "#AA1E1E1E", PanelBg:     "#551E1E1E",
            Text:        "#F0F0F0",   MutedText:   "#AAAAAA",   StatusOk:    "#AAAAAA",
            Hover:       "#15000000", Selection:   "#207B68EE", InactiveSel: "#107B68EE",
            ComboHover:  "#25000000", ComboSel:    "#45000000",
            Accent:      "#7B68EE");

        private static readonly ThemeColors DarkerTheme = new(
            RootBg:      "#CC0F0F0F", PanelBg:     "#550F0F0F",
            Text:        "#F0F0F0",   MutedText:   "#AAAAAA",   StatusOk:    "#AAAAAA",
            Hover:       "#15000000", Selection:   "#207B68EE", InactiveSel: "#107B68EE",
            ComboHover:  "#25000000", ComboSel:    "#45000000",
            Accent:      "#7B68EE");

        private static readonly ThemeColors LightTheme = new(
            RootBg:      "#88C8C8C8", PanelBg:     "#55FFFFFF",
            Text:        "#1A1A2E",   MutedText:   "#666666",   StatusOk:    "#444444",
            Hover:       "#18000000", Selection:   "#207B68EE", InactiveSel: "#107B68EE",
            ComboHover:  "#15000000", ComboSel:    "#30000000",
            Accent:      "#7B68EE");

        private void ApplyTheme(string theme)
        {
            if (theme == "Auto")
                theme = IsSystemDarkMode() ? "Dark" : "Light";

            ThemeColors colors = theme switch
            {
                "Darker" => DarkerTheme,
                "Light"  => LightTheme,
                _        => DarkTheme,
            };

            SetThemeColors(colors);
        }

        private void SetThemeColors(ThemeColors t)
        {
            var res = Application.Current.Resources;

            SolidColorBrush Brush(string hex)
            {
                var b = new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));
                b.Freeze();
                return b;
            }

            res["RootBackground"]      = Brush(t.RootBg);
            res["PanelBackground"]     = Brush(t.PanelBg);
            res["PrimaryText"]         = Brush(t.Text);
            res["MutedText"]           = Brush(t.MutedText);
            res["StatusOkText"]        = Brush(t.StatusOk);
            res["HoverBackground"]     = Brush(t.Hover);
            res["SelectionBackground"] = Brush(t.Selection);
            res["InactiveSelection"]   = Brush(t.InactiveSel);
            res["ComboHover"]          = Brush(t.ComboHover);
            res["ComboSelected"]       = Brush(t.ComboSel);
            res["AccentBrush"]         = Brush(t.Accent);

            // ListBox system brushes can't use DynamicResource — update them directly on the ListBox
            ProcessListBox.Resources[SystemColors.HighlightBrushKey]                      = Brush(t.Selection);
            ProcessListBox.Resources[SystemColors.HighlightTextBrushKey]                  = Brush(t.Text);
            ProcessListBox.Resources[SystemColors.InactiveSelectionHighlightBrushKey]     = Brush(t.InactiveSel);
            ProcessListBox.Resources[SystemColors.InactiveSelectionHighlightTextBrushKey] = Brush(t.Text);

            SetStatus(StatusTextBlock.Text, _lastStatusOk);
        }

        private const int DWMWA_SYSTEMBACKDROP_TYPE = 38;
        private const int DWMWA_WINDOW_CORNER_PREF  = 33;
        private const int DWMWCP_ROUND              = 2;
        private const int DWMSBT_TRANSIENTWINDOW    = 3; // Acrylic

        private void ApplyAcrylicBlur()
        {
            var helper = new WindowInteropHelper(this);
            if (helper.Handle == IntPtr.Zero) return;

            int backdropType = DWMSBT_TRANSIENTWINDOW;
            DwmSetWindowAttribute(helper.Handle, DWMWA_SYSTEMBACKDROP_TYPE, ref backdropType, Marshal.SizeOf<int>());

            int cornerPref = DWMWCP_ROUND;
            DwmSetWindowAttribute(helper.Handle, DWMWA_WINDOW_CORNER_PREF, ref cornerPref, Marshal.SizeOf<int>());
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWow64Process(IntPtr hProcess, [MarshalAs(UnmanagedType.Bool)] out bool wow64Process);

        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int dwAttribute, ref int pvAttribute, int cbAttribute);

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

    public record ProcessItem(string DisplayName, int Id, ImageSource? ProcessIcon);

    public record DllItem(string FileName, string FullPath);

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
