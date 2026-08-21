using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace FocusShield
{
    /// <summary>
    /// Reads and writes %APPDATA%\FocusShield\config.json — the app lists plus the
    /// handful of settings exposed through the tray menu.
    /// Whitelist: apps that are always allowed to come to the front.
    /// Blacklist: apps that are always kept in the background.
    /// </summary>
    internal class ConfigManager
    {
        // Sanity bounds for the hand-edited idle threshold.
        private const int MinIdleThresholdMs = 0;
        private const int MaxIdleThresholdMs = 5000;

        private readonly string _configPath;
        private HashSet<string> _whitelist;
        private HashSet<string> _blacklist;

        private bool _protectionEnabled    = true;
        private bool _backgroundNewWindows = true;
        private bool _showNotifications    = true;
        private int  _idleThresholdMs      = 300;

        private bool _loading;

        private class Config
        {
            [JsonPropertyName("whitelist")]
            public List<string> Whitelist { get; set; } = new();

            [JsonPropertyName("blacklist")]
            public List<string> Blacklist { get; set; } = new();

            [JsonPropertyName("protectionEnabled")]
            public bool? ProtectionEnabled { get; set; }

            [JsonPropertyName("backgroundNewWindows")]
            public bool? BackgroundNewWindows { get; set; }

            [JsonPropertyName("showNotifications")]
            public bool? ShowNotifications { get; set; }

            [JsonPropertyName("idleThresholdMs")]
            public int? IdleThresholdMs { get; set; }
        }

        public ConfigManager()
        {
            _configPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "FocusShield",
                "config.json");

            _whitelist = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            _blacklist = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            Load();
        }

        // ─── settings ────────────────────────────────────────────────────────────

        /// <summary>Master switch for the whole protection mechanism.</summary>
        public bool ProtectionEnabled
        {
            get => _protectionEnabled;
            set => Set(ref _protectionEnabled, value);
        }

        /// <summary>
        /// Keep windows that have just finished loading behind the user's work
        /// instead of letting them jump to the front.
        /// </summary>
        public bool BackgroundNewWindows
        {
            get => _backgroundNewWindows;
            set => Set(ref _backgroundNewWindows, value);
        }

        /// <summary>Show a balloon tip when a window is kept in the background.</summary>
        public bool ShowNotifications
        {
            get => _showNotifications;
            set => Set(ref _showNotifications, value);
        }

        /// <summary>
        /// How long after the last keypress or mouse move an activation may still
        /// count as user-initiated. Only editable in config.json.
        /// </summary>
        public int IdleThresholdMs => _idleThresholdMs;

        // ─── app lists ───────────────────────────────────────────────────────────

        public bool IsWhitelisted(string appName) =>
            !string.IsNullOrEmpty(appName) && _whitelist.Contains(appName);

        public bool IsBlacklisted(string appName) =>
            !string.IsNullOrEmpty(appName) && _blacklist.Contains(appName);

        public void AddToWhitelist(string appName) => Add(_whitelist, _blacklist, appName);

        public void AddToBlacklist(string appName) => Add(_blacklist, _whitelist, appName);

        public void RemoveFromWhitelist(string appName) => Remove(_whitelist, appName);

        public void RemoveFromBlacklist(string appName) => Remove(_blacklist, appName);

        /// <summary>Gets a sorted copy of the current whitelist.</summary>
        public IEnumerable<string> GetWhitelist() => Sorted(_whitelist);

        /// <summary>Gets a sorted copy of the current blacklist.</summary>
        public IEnumerable<string> GetBlacklist() => Sorted(_blacklist);

        /// <summary>
        /// Strips a typed-in ".exe" and surrounding whitespace so "Steam.exe",
        /// " steam " and "steam" all end up as the same entry.
        /// </summary>
        public static string Normalize(string appName)
        {
            if (string.IsNullOrWhiteSpace(appName))
                return string.Empty;

            string name = appName.Trim().Trim('"');
            if (name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                name = name.Substring(0, name.Length - 4);

            return name.Trim();
        }

        // ─── helpers ─────────────────────────────────────────────────────────────

        private void Set(ref bool field, bool value)
        {
            if (field == value) return;
            field = value;
            Save();
        }

        /// <summary>
        /// Adds to <paramref name="target"/> and drops the entry from
        /// <paramref name="other"/> — an app cannot be both allowed and blocked.
        /// </summary>
        private void Add(HashSet<string> target, HashSet<string> other, string appName)
        {
            string name = Normalize(appName);
            if (name.Length == 0) return;

            other.Remove(name);
            if (target.Add(name))
                Save();
        }

        private void Remove(HashSet<string> list, string appName)
        {
            if (list.Remove(Normalize(appName)))
                Save();
        }

        private static List<string> Sorted(HashSet<string> list)
        {
            var copy = new List<string>(list);
            copy.Sort(StringComparer.OrdinalIgnoreCase);
            return copy;
        }

        private void Load()
        {
            _loading = true;
            try
            {
                if (!File.Exists(_configPath))
                    return;

                var config = JsonSerializer.Deserialize<Config>(File.ReadAllText(_configPath));
                if (config == null)
                    return;

                if (config.Whitelist != null)
                    foreach (string app in config.Whitelist)
                        AddNormalized(_whitelist, app);

                if (config.Blacklist != null)
                    foreach (string app in config.Blacklist)
                        AddNormalized(_blacklist, app);

                _protectionEnabled    = config.ProtectionEnabled    ?? _protectionEnabled;
                _backgroundNewWindows = config.BackgroundNewWindows ?? _backgroundNewWindows;
                _showNotifications    = config.ShowNotifications    ?? _showNotifications;

                if (config.IdleThresholdMs.HasValue)
                    _idleThresholdMs = Math.Min(MaxIdleThresholdMs,
                        Math.Max(MinIdleThresholdMs, config.IdleThresholdMs.Value));
            }
            catch (Exception ex)
            {
                // A damaged config must never stop FocusShield from starting — fall
                // back to defaults and tell the user where to look.
                System.Windows.Forms.MessageBox.Show(
                    $"Failed to load FocusShield config, using defaults.\n\n{_configPath}\n\n{ex.Message}",
                    "FocusShield", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
            }
            finally
            {
                _loading = false;
            }
        }

        private static void AddNormalized(HashSet<string> list, string appName)
        {
            string name = Normalize(appName);
            if (name.Length > 0)
                list.Add(name);
        }

        private void Save()
        {
            if (_loading) return;

            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_configPath));

                var config = new Config
                {
                    Whitelist            = Sorted(_whitelist),
                    Blacklist            = Sorted(_blacklist),
                    ProtectionEnabled    = _protectionEnabled,
                    BackgroundNewWindows = _backgroundNewWindows,
                    ShowNotifications    = _showNotifications,
                    IdleThresholdMs      = _idleThresholdMs
                };

                var options = new JsonSerializerOptions { WriteIndented = true };
                File.WriteAllText(_configPath, JsonSerializer.Serialize(config, options));
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Failed to save FocusShield config: {ex.Message}",
                    "FocusShield", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
    }
}
