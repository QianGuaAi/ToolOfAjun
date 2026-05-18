using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;

namespace MyTools.ViewModels
{
    /// <summary>
    /// 看图模块用的资源管理器风格的文件夹树节点。懒加载子目录：
    /// 每个未展开节点持有一个占位 dummy 子节点，TreeViewItem 展开时一次性替换为真实子目录。
    /// </summary>
    public class ImageFolderNode : INotifyPropertyChanged
    {
        private static readonly ImageFolderNode DummyChild = new ImageFolderNode { Name = "...", FullPath = null };

        public string Name { get; set; }
        public string FullPath { get; set; }
        public ObservableCollection<ImageFolderNode> Children { get; } = new ObservableCollection<ImageFolderNode>();

        private bool _childrenLoaded;
        private bool _isExpanded;
        public bool IsExpanded
        {
            get => _isExpanded;
            set
            {
                if (_isExpanded == value) return;
                _isExpanded = value;
                OnPropertyChanged();
                if (value) LoadChildren();
            }
        }

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set { if (_isSelected == value) return; _isSelected = value; OnPropertyChanged(); }
        }

        public bool IsDummy => FullPath == null;

        public static ObservableCollection<ImageFolderNode> CreateRoots()
        {
            var roots = new ObservableCollection<ImageFolderNode>();
            try
            {
                foreach (var drive in DriveInfo.GetDrives())
                {
                    try
                    {
                        if (!drive.IsReady) continue;
                        var label = string.IsNullOrEmpty(drive.VolumeLabel) ? drive.Name.TrimEnd('\\') : $"{drive.Name.TrimEnd('\\')} ({drive.VolumeLabel})";
                        var node = new ImageFolderNode { Name = label, FullPath = drive.RootDirectory.FullName };
                        node.Children.Add(DummyChild);
                        roots.Add(node);
                    }
                    catch { /* 单个驱动器失败不阻塞其他 */ }
                }
            }
            catch { /* swallow */ }
            return roots;
        }

        private void LoadChildren()
        {
            if (_childrenLoaded) return;
            _childrenLoaded = true;
            Children.Clear();
            if (string.IsNullOrEmpty(FullPath) || !Directory.Exists(FullPath)) return;
            try
            {
                var subDirs = Directory.EnumerateDirectories(FullPath)
                    .Where(IsListableDirectory)
                    .OrderBy(d => Path.GetFileName(d), StringComparer.CurrentCultureIgnoreCase)
                    .ToList();
                foreach (var dir in subDirs)
                {
                    var name = Path.GetFileName(dir);
                    if (string.IsNullOrEmpty(name)) name = dir;
                    var node = new ImageFolderNode { Name = name, FullPath = dir };
                    node.Children.Add(DummyChild);
                    Children.Add(node);
                }
            }
            catch { /* 没权限的目录直接跳过 */ }
        }

        private static bool IsListableDirectory(string path)
        {
            try
            {
                var info = new DirectoryInfo(path);
                if ((info.Attributes & (FileAttributes.System | FileAttributes.Hidden)) != 0) return false;
                return true;
            }
            catch { return false; }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
