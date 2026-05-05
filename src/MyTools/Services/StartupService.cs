using System;
using System.Collections.Generic;
using Microsoft.Win32;

namespace MyTools.Services
{
    public class StartupItem
    {
        public string Name { get; set; }
        public string Command { get; set; }
        public string Location { get; set; }
        public bool IsEnabled { get; set; }
        public bool IsUserLevel { get; set; }
    }

    public static class StartupService
    {
        private const string RunKeyPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
        private const string BackupKeyPath = @"SOFTWARE\AJunTools\DisabledRun";

        public static List<StartupItem> GetStartupItems()
        {
            var items = new List<StartupItem>();

            // Read Enabled items
            AddItemsFromKey(Registry.CurrentUser, RunKeyPath, items, true, true);
            AddItemsFromKey(Registry.LocalMachine, RunKeyPath, items, true, false);

            // Read Disabled items
            AddItemsFromKey(Registry.CurrentUser, BackupKeyPath, items, false, true);
            AddItemsFromKey(Registry.LocalMachine, BackupKeyPath, items, false, false);

            return items;
        }

        private static void AddItemsFromKey(RegistryKey root, string path, List<StartupItem> list, bool isEnabled, bool isUserLevel)
        {
            using (var key = root.OpenSubKey(path))
            {
                if (key != null)
                {
                    foreach (var name in key.GetValueNames())
                    {
                        list.Add(new StartupItem
                        {
                            Name = name,
                            Command = key.GetValue(name)?.ToString(),
                            Location = isUserLevel ? "当前用户" : "所有用户",
                            IsEnabled = isEnabled,
                            IsUserLevel = isUserLevel
                        });
                    }
                }
            }
        }

        public static void ToggleStartupItem(StartupItem item)
        {
            try
            {
                var root = item.IsUserLevel ? Registry.CurrentUser : Registry.LocalMachine;
                string sourcePath = item.IsEnabled ? RunKeyPath : BackupKeyPath;
                string targetPath = item.IsEnabled ? BackupKeyPath : RunKeyPath;

                using (var sourceKey = root.OpenSubKey(sourcePath, true))
                using (var targetKey = root.CreateSubKey(targetPath))
                {
                    if (sourceKey != null && targetKey != null)
                    {
                        var value = sourceKey.GetValue(item.Name);
                        if (value != null)
                        {
                            targetKey.SetValue(item.Name, value);
                            sourceKey.DeleteValue(item.Name);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"操作失败: {ex.Message}\n请尝试以管理员身份运行程序。", "错误", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }

        public static void DeleteStartupItem(StartupItem item)
        {
            try
            {
                var root = item.IsUserLevel ? Registry.CurrentUser : Registry.LocalMachine;
                string path = item.IsEnabled ? RunKeyPath : BackupKeyPath;
                using (var key = root.OpenSubKey(path, true))
                {
                    if (key != null)
                    {
                        key.DeleteValue(item.Name, false);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"删除失败: {ex.Message}\n请尝试以管理员身份运行程序。", "错误", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }
    }
}
