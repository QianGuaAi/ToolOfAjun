using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Win32;

class TestWeChatLocator
{
    public static void Main()
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        Console.WriteLine("=== WeChat Data Locator Debug ===\n");

        // 1. Documents 路径
        var docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        Console.WriteLine($"MyDocuments: {docs}");
        Console.WriteLine($"  GetFullPath: {Path.GetFullPath(docs)}");
        Console.WriteLine($"  Exists: {Directory.Exists(docs)}");

        // 2. 穿透 Junction 后实际有什么
        Console.WriteLine($"\n  Contents through junction:");
        try {
            foreach (var e in Directory.EnumerateFileSystemEntries(docs).Take(20))
                Console.WriteLine($"    {Path.GetFileName(e)}");
        } catch (Exception ex) {
            Console.WriteLine($"    ERROR: {ex.Message}");
        }

        // 3. 检查各关键路径
        var paths = new[] {
            Path.Combine(docs, "WeChat Files"),
            Path.Combine(docs, "Tencent Files"),
            Path.Combine(docs, "xwechat_files"),
            Path.Combine(docs, "xwechat_files", "all_users"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Tencent", "xwechat_files"),
        };

        Console.WriteLine("\n=== Key paths ===");
        foreach (var p in paths) {
            var full = Path.GetFullPath(p);
            Console.WriteLine($"  {full}");
            Console.WriteLine($"    Exists={Directory.Exists(full)}, GetFullPath={full}");
            if (Directory.Exists(full)) {
                try {
                    var subdirs = Directory.GetDirectories(full, "*", SearchOption.TopDirectoryOnly);
                    Console.WriteLine($"    Subdirs ({subdirs.Length}): {string.Join(", ", subdirs.Select(Path.GetFileName))}");
                } catch (Exception ex) {
                    Console.WriteLine($"    Subdirs ERROR: {ex.Message}");
                }
            }
        }

        // 4. 注册表
        Console.WriteLine("\n=== Registry ===");
        var keys = new[] { @"Software\Tencent\WeChat", @"Software\Tencent\Weixin" };
        foreach (var k in keys) {
            try {
                using (var key = Registry.CurrentUser.OpenSubKey(k)) {
                    if (key == null) {
                        Console.WriteLine($"  {k}: (not found)");
                    } else {
                        Console.WriteLine($"  {k}:");
                        foreach (var name in key.GetValueNames()) {
                            Console.WriteLine($"    {name} = {key.GetValue(name)}");
                        }
                    }
                }
            } catch (Exception ex) {
                Console.WriteLine($"  {k}: ERROR {ex.Message}");
            }
        }

        // 5. WxId 匹配测试
        Console.WriteLine("\n=== WxId Pattern Test ===");
        var pattern = new System.Text.RegularExpressions.Regex(@"^(wxid_[A-Za-z0-9]+|\d+_.+)$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);

        var testNames = new[] {
            "wxid_6evo1mkpqh1c22_d613",
            "262679118",
            "3523174748",
            "nt_qq",
            "Tencent Files",
            "xwechat_files",
        };
        foreach (var n in testNames)
            Console.WriteLine($"  '{n}': {pattern.IsMatch(n)}");

        // 6. 模拟 AddRootsFromBasePath 行为
        Console.WriteLine("\n=== Simulated scan ===");
        var basePath = Path.Combine(docs, "xwechat_files");
        Console.WriteLine($"BasePath: {basePath}");
        Console.WriteLine($"Exists: {Directory.Exists(basePath)}");
        if (Directory.Exists(basePath)) {
            try {
                var children = Directory.GetDirectories(basePath, "*", SearchOption.TopDirectoryOnly);
                Console.WriteLine($"Children ({children.Length}):");
                foreach (var child in children) {
                    var name = Path.GetFileName(child.TrimEnd('\\'));
                    var match = pattern.IsMatch(name);
                    Console.WriteLine($"  [{match}] {name} -> {child}");
                }
            } catch (Exception ex) {
                Console.WriteLine($"  ERROR: {ex.Message}");
            }
        }

        // 7. User Shell Folders
        Console.WriteLine("\n=== User Shell Folders ===");
        try {
            using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders")) {
                var val = key?.GetValue("Personal") as string;
                Console.WriteLine($"  Personal (User Shell Folders): {val}");
                if (!string.IsNullOrEmpty(val)) {
                    var expanded = Environment.ExpandEnvironmentVariables(val);
                    Console.WriteLine($"  Expanded: {expanded}");
                    Console.WriteLine($"  GetFullPath: {Path.GetFullPath(expanded)}");
                }
            }
        } catch (Exception ex) {
            Console.WriteLine($"  ERROR: {ex.Message}");
        }

        Console.WriteLine("\n=== Done ===");
        Console.ReadLine();
    }
}
