using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace FocusShield
{
    /// <summary>
    /// Manages the whitelist and blacklist of applications.
    /// Whitelist: apps that are always allowed to steal focus.
    /// Blacklist: apps that should always be blocked.
    /// </summary>
    internal class AppListManager
    {
        private readonly string _configPath;
        private HashSet<string> _whitelist;
        private HashSet<string> _blacklist;

        private class Config
        {
            [JsonPropertyName("whitelist")]
            public List<string> Whitelist { get; set; } = new();

            [JsonPropertyName("blacklist")]
            public List<string> Blacklist { get; set; } = new();
        }

        public AppListManager()
        {
            _configPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "FocusShield",
                "config.json");

            _whitelist = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            _blacklist = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            Load();
        }

        /// <summary>
        /// Gets the executable name from a process ID.
        /// </summary>
        private static string GetProcessName(uint pid)
        {
            try
            {
                var proc = System.Diagnostics.Process.GetProcessById((int)pid);
                return Path.GetFileNameWithoutExtension(proc.ProcessName);
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Checks if an app should be whitelisted (always allowed to steal focus).
        /// </summary>
        public bool IsWhitelisted(uint pid)
        {
            string name = GetProcessName(pid);
            return !string.IsNullOrEmpty(name) && _whitelist.Contains(name);
        }

        /// <summary>
        /// Checks if an app should be blacklisted (always blocked).
        /// </summary>
        public bool IsBlacklisted(uint pid)
        {
            string name = GetProcessName(pid);
            return !string.IsNullOrEmpty(name) && _blacklist.Contains(name);
        }

        /// <summary>
        /// Adds an app to the whitelist.
        /// </summary>
        public void AddToWhitelist(string appName)
        {
            _whitelist.Add(appName);
            Save();
        }

        /// <summary>
        /// Adds an app to the blacklist.
        /// </summary>
        public void AddToBlacklist(string appName)
        {
            _blacklist.Add(appName);
            Save();
        }

        /// <summary>
        /// Removes an app from the whitelist.
        /// </summary>
        public void RemoveFromWhitelist(string appName)
        {
            _whitelist.Remove(appName);
            Save();
        }

        /// <summary>
        /// Removes an app from the blacklist.
        /// </summary>
        public void RemoveFromBlacklist(string appName)
        {
            _blacklist.Remove(appName);
            Save();
        }

        /// <summary>
        /// Gets a copy of the current whitelist.
        /// </summary>
        public IEnumerable<string> GetWhitelist() => new List<string>(_whitelist);

        /// <summary>
        /// Gets a copy of the current blacklist.
        /// </summary>
        public IEnumerable<string> GetBlacklist() => new List<string>(_blacklist);

        private void Load()
        {
            try
            {
                if (File.Exists(_configPath))
                {
                    string json = File.ReadAllText(_configPath);
                    var config = JsonSerializer.Deserialize<Config>(json);

                    if (config?.Whitelist != null)
                        _whitelist = new HashSet<string>(config.Whitelist, StringComparer.OrdinalIgnoreCase);

                    if (config?.Blacklist != null)
                        _blacklist = new HashSet<string>(config.Blacklist, StringComparer.OrdinalIgnoreCase);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Failed to load FocusShield config: {ex.Message}",
                    "Error", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private void Save()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_configPath));

                var config = new Config
                {
                    Whitelist = new List<string>(_whitelist),
                    Blacklist = new List<string>(_blacklist)
                };

                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(config, options);
                File.WriteAllText(_configPath, json);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Failed to save FocusShield config: {ex.Message}",
                    "Error", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
    }
}
