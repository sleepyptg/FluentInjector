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
        private bool _autoInjectPending = false;

        // Icon cache: exe path → ImageSource (null = no icon available)
        private readonly ConcurrentDictionary<string, ImageSource?> _iconCache = new(StringComparer.OrdinalIgnoreCase);

        // Status color for the "ok" (non-error) state is managed via Application.Current.Resources["StatusOkText"]

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
            };
            DllListControl.ItemsSource = _selectedDlls;
            _selectedDlls.CollectionChanged += (_, _) => ScheduleSave();
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            ApplyAcrylicBlur();
            // Detach handler so setting default index doesn't overwrite saved config
            ThemeComboBox.SelectionChanged -= ThemeComboBox_SelectionChanged;
            ThemeComboBox.SelectedIndex = 0;
            ThemeComboBox.SelectionChanged += ThemeComboBox_SelectionChanged;
            LoadConfig();
            // Apply whatever theme is now selected (restored from config or default)
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
                        // Skip processes without a visible window
                        if (string.IsNullOrEmpty(p.MainWindowTitle)) continue;
                        processes.Add(new ProcessItem
                        {
                            DisplayName = $"{p.ProcessName}.exe",
                            Id = p.Id,
                            ProcessIcon = GetProcessIconCached(p)
                        });
                    }
                    catch
                    {
                        // Process exited between snapshot and property access — skip it
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

                        // Preserve name across ItemsSource reassignment — ApplySearchFilter
                        // clears SelectedItem which fires SelectionChanged and wipes _lastSelectedProcessName
                        var savedName = _lastSelectedProcessName;
                        ApplySearchFilter();
                        if (string.IsNullOrEmpty(_lastSelectedProcessName))
                            _lastSelectedProcessName = savedName;

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

                        // Auto-select if saved process name appears in the new list
                        if (ProcessListBox.SelectedItem == null && !string.IsNullOrEmpty(_lastSelectedProcessName))
                        {
                            var match = _allProcesses.FirstOrDefault(p => p.DisplayName == _lastSelectedProcessName);
                            if (match != null)
                                ProcessListBox.SelectedItem = match;
                        }
                    }

                    // Auto Inject: check every tick — not gated by list change
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

                // Return cached result (including null for "no icon") without hitting disk again
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

        private void ProcessListBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (ProcessListBox.SelectedItem is ProcessItem p)
            {
                _lastSelectedProcessName = p.DisplayName;
                SetStatus($"Status: Selected {p.DisplayName} (PID {p.Id})", true);
                ScheduleSave();
            }
            else if (e.RemovedItems.Count > 0)
            {
                // User manually deselected — clear saved name so we don't show misleading "waiting" messages
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
                InjectBtn.IsEnabled = false;
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

                // Hold the result message for at least 2.5s so the user can read it
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
                var p = Process.GetProcessById(target.Id);
                if (p.HasExited)
                {
                    _autoInjectPending = true;
                    return;
                }
            }
            catch
            {
                // Process no longer exists — re-arm and wait for next launch
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

                // Re-arm so the next process launch triggers injection again
                if (AutoInjectCheckBox.IsChecked == true)
                {
                    _autoInjectPending = true;
                    SetStatus($"Status: Auto Inject armed — waiting for {_lastSelectedProcessName}", true);
                }
            }
            catch (Exception ex)
            {
                SetStatus($"Status: Auto Inject error — {ex.Message}", false);
            }
            finally
            {
                InjectSpinner.Visibility = Visibility.Collapsed;
                InjectBtn.IsEnabled = true;
            }
        }

        // Standard LoadLibraryW injection
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

                    try
                    {
                        // WAIT_OBJECT_0 = 0; timeout or failure means injection didn't complete
                        uint waitResult = WaitForSingleObject(hThread, 5000);
                        if (waitResult != 0) return false;

                        // Thread exit code = LoadLibraryW return value; 0 means it failed to load
                        GetExitCodeThread(hThread, out uint exitCode);
                        return exitCode != 0;
                    }
                    finally
                    {
                        CloseHandle(hThread);
                    }
                }
                finally
                {
                    // Always free remote memory — LoadLibraryW has already copied the path internally
                    VirtualFreeEx(hProcess, mem, 0, AllocationType.Release);
                }
            }
            finally
            {
                CloseHandle(hProcess);
            }
        }

        // ── Config ─────────────────────────────────────────────────────────────

        // Debounce saves: reset a 300ms timer on each call, only write when it fires
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
                var config = new
                {
                    ProcessName = (ProcessListBox.SelectedItem as ProcessItem)?.DisplayName ?? "",
                    ProcessId   = (ProcessListBox.SelectedItem as ProcessItem)?.Id ?? 0,
                    DllPaths    = _selectedDlls.Select(d => d.FullPath).ToList(),
                    Theme       = selectedTheme,
                    AutoInject  = AutoInjectCheckBox.IsChecked == true,
                    SavedAt     = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
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

                // Just set the name — LoadProcesses will auto-select when the list populates
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
            }
            catch { /* Silent fail on load */ }
        }

        // ── Status ─────────────────────────────────────────────────────────────

        private void SetStatus(string message, bool ok)
        {
            StatusTextBlock.Text = message;
            StatusTextBlock.Foreground = ok
                ? (SolidColorBrush)Application.Current.Resources["StatusOkText"]
                : (SolidColorBrush)Application.Current.Resources["ErrorText"];
        }

        // ── Theme ──────────────────────────────────────────────────────────────

        private void ThemeComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (ThemeComboBox.SelectedItem is not System.Windows.Controls.ComboBoxItem item) return;
            ApplyTheme(item.Tag?.ToString() ?? "Dark");
            ScheduleSave();
        }

        private void ApplyTheme(string theme)
        {
            switch (theme)
            {
                // rootBg, panelBg, text, mutedText, statusOk, hover, selection, inactiveSel, comboHover, comboSel, accent
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

            // Re-apply status to pick up the new StatusOkText brush
            bool isOk = StatusTextBlock.Foreground is SolidColorBrush sb
                && sb.Color != (Color)ColorConverter.ConvertFromString("#D32F2F");
            SetStatus(StatusTextBlock.Text, isOk);
        }

        // ── Acrylic blur ───────────────────────────────────────────────────────

        private void ApplyAcrylicBlur()
        {
            var helper = new WindowInteropHelper(this);
            if (helper.Handle == IntPtr.Zero) return;

            // DWMWA_SYSTEMBACKDROP_TYPE = 38, requires Windows 11 22H2 (build 22621+)
            // Values: 0=Auto, 1=None, 2=Mica, 3=Acrylic, 4=Tabbed
            int backdropType = 3;
            DwmSetWindowAttribute(helper.Handle, 38, ref backdropType, Marshal.SizeOf<int>());

            // DWMWA_WINDOW_CORNER_PREFERENCE = 33, requires Windows 11
            // Values: 0=Default, 1=DoNotRound, 2=Round, 3=RoundSmall
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
