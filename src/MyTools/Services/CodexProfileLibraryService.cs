using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using MyTools.ViewModels;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace MyTools.Services
{
    public static class CodexProfileLibraryService
    {
        public const int CurrentSchemaVersion = 2;
        public const string StatusOk = "正常";
        public const string StatusWarn = "即将过期";
        public const string StatusExpired = "已过期";
        public const string StatusUnknown = "未知";
        public const string ExportExtension = ".codexbox";

        private const int ExportSchemaVersion = 1;
        private const int ExportIterations = 200000;
        private const string ExportHeader = "CDXB";
        private const string ExportPortableKind = "portable-codex-profiles-v2";
        private static readonly SemaphoreSlim FileLock = new SemaphoreSlim(1, 1);
        private static CodexProfilesFile _cachedProfiles;

        public static string RootFolderPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MyTools", "Codex");
        public static string ProfilesFilePath => Path.Combine(RootFolderPath, "profiles.json");
        public static string ActiveFilePath => Path.Combine(RootFolderPath, "active.json");
        public static string BackupsFolderPath => Path.Combine(RootFolderPath, "Backups");
        public static string CodexFolderPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".codex");

        public static async Task<CodexProfilesFile> LoadAsync(CancellationToken ct)
        {
            await FileLock.WaitAsync(ct).ConfigureAwait(false);
            try
            {
                if (File.Exists(ProfilesFilePath))
                {
                    var protectedText = await ReadAllTextAsync(ProfilesFilePath, ct).ConfigureAwait(false);
                    var jsonBytes = CodexConfigProfileService.UnprotectBytesFromBase64(protectedText);
                    var file = DeserializeProfilesFile(jsonBytes);
                    NormalizeProfilesFile(file);
                    _cachedProfiles = file;
                    return file;
                }

                var migrated = await TryLoadLegacyProfilesAsync(ct).ConfigureAwait(false);
                if (migrated.items.Count > 0)
                {
                    await SaveCoreAsync(migrated, ct).ConfigureAwait(false);
                }

                _cachedProfiles = migrated;
                return migrated;
            }
            catch (Exception ex)
            {
                AppLogService.Error(new InvalidOperationException(ex.Message), "Loading Codex profile library failed with {ErrorType}", ex.GetType().Name);
                _cachedProfiles = CreateEmptyProfilesFile();
                return _cachedProfiles;
            }
            finally
            {
                FileLock.Release();
            }
        }

        public static CodexProfilesFile GetCachedProfiles()
        {
            return _cachedProfiles ?? CreateEmptyProfilesFile();
        }

        public static async Task SaveAsync(CodexProfilesFile file, CancellationToken ct)
        {
            if (file == null)
            {
                return;
            }

            await FileLock.WaitAsync(ct).ConfigureAwait(false);
            try
            {
                await SaveCoreAsync(file, ct).ConfigureAwait(false);
                _cachedProfiles = file;
            }
            finally
            {
                FileLock.Release();
            }
        }

        public static async Task<string> BackupProfilesBeforeConfigTemplateSyncAsync(CodexProfilesFile file, string templateDisplayName, CancellationToken ct)
        {
            if (file == null)
            {
                throw new ArgumentNullException(nameof(file));
            }

            await FileLock.WaitAsync(ct).ConfigureAwait(false);
            try
            {
                NormalizeProfilesFile(file);
                Directory.CreateDirectory(BackupsFolderPath);

                var package = new CodexProfilesLibraryBackupPackage
                {
                    BackupKind = "codex-profile-library-before-config-template-sync",
                    TemplateDisplayName = templateDisplayName ?? string.Empty,
                    CreatedAtUtc = DateTime.UtcNow,
                    Profiles = file
                };

                var safeName = SanitizeFileName(string.IsNullOrWhiteSpace(templateDisplayName) ? "codex" : templateDisplayName);
                var backupPath = Path.Combine(
                    BackupsFolderPath,
                    "profiles_config_sync_" + safeName + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".profiles.dpapi");
                var json = JsonConvert.SerializeObject(package, Formatting.Indented);
                var protectedText = CodexConfigProfileService.ProtectBytesToBase64(Encoding.UTF8.GetBytes(json));
                await WriteAllTextAsync(backupPath, protectedText, ct).ConfigureAwait(false);
                return backupPath;
            }
            finally
            {
                FileLock.Release();
            }
        }

        public static async Task<string> BackupCurrentCodexFolderAsync(string activeDisplayName, CancellationToken ct)
        {
            var configPath = Path.Combine(CodexFolderPath, CodexConfigProfileService.ConfigFileName);
            var authPath = Path.Combine(CodexFolderPath, CodexConfigProfileService.AuthFileName);
            if (!File.Exists(configPath) && !File.Exists(authPath))
            {
                return string.Empty;
            }

            var package = new CodexProfileBackupPackage
            {
                ActiveDisplayName = activeDisplayName ?? string.Empty,
                CreatedAtUtc = DateTime.UtcNow,
                ConfigTomlBase64 = File.Exists(configPath) ? Convert.ToBase64String(await ReadAllBytesAsync(configPath, ct).ConfigureAwait(false)) : string.Empty,
                AuthJsonBase64 = File.Exists(authPath) ? Convert.ToBase64String(await ReadAllBytesAsync(authPath, ct).ConfigureAwait(false)) : string.Empty
            };

            Directory.CreateDirectory(BackupsFolderPath);
            var safeName = SanitizeFileName(string.IsNullOrWhiteSpace(activeDisplayName) ? "codex" : activeDisplayName);
            var backupPath = Path.Combine(BackupsFolderPath, safeName + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".bak.dpapi");
            var json = JsonConvert.SerializeObject(package, Formatting.None);
            var protectedText = CodexConfigProfileService.ProtectBytesToBase64(Encoding.UTF8.GetBytes(json));
            await WriteAllTextAsync(backupPath, protectedText, ct).ConfigureAwait(false);
            return backupPath;
        }

        public static async Task<string> RestoreLatestBackupAsync(CancellationToken ct)
        {
            if (!Directory.Exists(BackupsFolderPath))
            {
                throw new FileNotFoundException("未找到 Codex 切换备份目录。");
            }

            var backupPath = Directory.GetFiles(BackupsFolderPath, "*.bak.dpapi")
                .Where(path =>
                {
                    var fileName = Path.GetFileName(path) ?? string.Empty;
                    return !fileName.EndsWith(".profiles.bak.dpapi", StringComparison.OrdinalIgnoreCase)
                        && !fileName.EndsWith(".profiles.dpapi", StringComparison.OrdinalIgnoreCase);
                })
                .OrderByDescending(File.GetLastWriteTimeUtc)
                .FirstOrDefault();
            if (string.IsNullOrWhiteSpace(backupPath))
            {
                throw new FileNotFoundException("未找到可回滚的 Codex 备份文件。");
            }

            var protectedText = await ReadAllTextAsync(backupPath, ct).ConfigureAwait(false);
            var jsonBytes = CodexConfigProfileService.UnprotectBytesFromBase64(protectedText);
            if (jsonBytes == null)
            {
                throw new InvalidOperationException("Codex 备份解密失败。");
            }

            var package = JsonConvert.DeserializeObject<CodexProfileBackupPackage>(Encoding.UTF8.GetString(jsonBytes));
            if (package == null)
            {
                throw new InvalidOperationException("Codex 备份内容无效。");
            }

            Directory.CreateDirectory(CodexFolderPath);
            if (!string.IsNullOrWhiteSpace(package.ConfigTomlBase64))
            {
                await WriteAllBytesAsync(Path.Combine(CodexFolderPath, CodexConfigProfileService.ConfigFileName), Convert.FromBase64String(package.ConfigTomlBase64), ct).ConfigureAwait(false);
            }

            if (!string.IsNullOrWhiteSpace(package.AuthJsonBase64))
            {
                await WriteAllBytesAsync(Path.Combine(CodexFolderPath, CodexConfigProfileService.AuthFileName), Convert.FromBase64String(package.AuthJsonBase64), ct).ConfigureAwait(false);
            }

            return backupPath;
        }

        public static async Task<CodexActiveFile> LoadActiveAsync(CancellationToken ct)
        {
            try
            {
                if (!File.Exists(ActiveFilePath))
                {
                    return new CodexActiveFile();
                }

                var json = await ReadAllTextAsync(ActiveFilePath, ct).ConfigureAwait(false);
                return JsonConvert.DeserializeObject<CodexActiveFile>(json) ?? new CodexActiveFile();
            }
            catch
            {
                return new CodexActiveFile();
            }
        }

        public static async Task SaveActiveAsync(CodexActiveFile active, CancellationToken ct)
        {
            Directory.CreateDirectory(RootFolderPath);
            var json = JsonConvert.SerializeObject(active ?? new CodexActiveFile(), Formatting.Indented);
            await WriteAllTextAsync(ActiveFilePath, json, ct).ConfigureAwait(false);
        }

        public static async Task ExportBoxAsync(CodexProfilesFile file, string outputPath, string password, CancellationToken ct)
        {
            if (file == null)
            {
                throw new ArgumentNullException(nameof(file));
            }

            if (string.IsNullOrWhiteSpace(outputPath))
            {
                throw new ArgumentException("导出路径不能为空。", nameof(outputPath));
            }

            if (string.IsNullOrEmpty(password))
            {
                throw new InvalidOperationException("导出口令不能为空。");
            }

            var portableFile = CreatePortableExportFile(file);
            var plainJson = JsonConvert.SerializeObject(portableFile, Formatting.Indented);
            var plaintext = Encoding.UTF8.GetBytes(plainJson);
            var salt = CreateRandomBytes(16);
            var iv = CreateRandomBytes(16);
            var keyMaterial = DeriveExportKey(password, salt);
            var ciphertext = EncryptAesCbc(plaintext, keyMaterial.Take(32).ToArray(), iv);

            byte[] preMac;
            using (var memory = new MemoryStream())
            using (var writer = new BinaryWriter(memory, Encoding.UTF8))
            {
                writer.Write(Encoding.ASCII.GetBytes(ExportHeader));
                writer.Write((uint)ExportSchemaVersion);
                writer.Write((ushort)salt.Length);
                writer.Write(salt);
                writer.Write((uint)ExportIterations);
                writer.Write((ushort)iv.Length);
                writer.Write(iv);
                writer.Write((uint)ciphertext.Length);
                writer.Write(ciphertext);
                writer.Flush();
                preMac = memory.ToArray();
            }

            byte[] mac;
            using (var hmac = new HMACSHA256(keyMaterial.Skip(32).Take(32).ToArray()))
            {
                mac = hmac.ComputeHash(preMac);
            }

            using (var memory = new MemoryStream())
            using (var writer = new BinaryWriter(memory, Encoding.UTF8))
            {
                writer.Write(preMac);
                writer.Write((ushort)mac.Length);
                writer.Write(mac);
                writer.Flush();
                Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? RootFolderPath);
                await WriteAllBytesAsync(outputPath, memory.ToArray(), ct).ConfigureAwait(false);
            }
        }

        public static async Task<CodexProfilesFile> ImportBoxAsync(string inputPath, string password, CancellationToken ct)
        {
            if (string.IsNullOrWhiteSpace(inputPath))
            {
                throw new ArgumentException("导入路径不能为空。", nameof(inputPath));
            }

            if (string.IsNullOrEmpty(password))
            {
                throw new InvalidOperationException("导入口令不能为空。");
            }

            var bytes = await ReadAllBytesAsync(inputPath, ct).ConfigureAwait(false);
            byte[] salt;
            byte[] iv;
            byte[] ciphertext;
            byte[] mac;
            int macDataLength;
            uint iterations;

            using (var memory = new MemoryStream(bytes))
            using (var reader = new BinaryReader(memory, Encoding.UTF8))
            {
                var headerBytes = reader.ReadBytes(4);
                if (!headerBytes.SequenceEqual(Encoding.ASCII.GetBytes(ExportHeader)))
                {
                    throw new InvalidOperationException("不是有效的 Codex 加密档案包。");
                }

                var version = reader.ReadUInt32();
                if (version != ExportSchemaVersion)
                {
                    throw new InvalidOperationException("Codex 加密档案包版本不受支持。");
                }

                var saltLen = reader.ReadUInt16();
                salt = reader.ReadBytes(saltLen);
                iterations = reader.ReadUInt32();
                if (iterations != ExportIterations)
                {
                    throw new InvalidOperationException("Codex 加密档案包 KDF 参数不受支持。");
                }

                var ivLen = reader.ReadUInt16();
                iv = reader.ReadBytes(ivLen);
                var ciphertextLen = reader.ReadUInt32();
                ciphertext = reader.ReadBytes(checked((int)ciphertextLen));
                macDataLength = checked((int)memory.Position);
                var macLen = reader.ReadUInt16();
                mac = reader.ReadBytes(macLen);
            }

            if (salt.Length != 16 || iv.Length != 16 || mac.Length != 32)
            {
                throw new InvalidOperationException("Codex 加密档案包参数无效。");
            }

            var keyMaterial = DeriveExportKey(password, salt);
            byte[] computed;
            using (var hmac = new HMACSHA256(keyMaterial.Skip(32).Take(32).ToArray()))
            {
                computed = hmac.ComputeHash(bytes.Take(macDataLength).ToArray());
            }

            if (!ConstantTimeEquals(mac, computed))
            {
                throw new InvalidOperationException("口令错误或文件损坏。");
            }

            var plaintext = DecryptAesCbc(ciphertext, keyMaterial.Take(32).ToArray(), iv);
            return DeserializeImportPackage(plaintext);
        }

        public static DateTime? ParseAccessTokenExp(byte[] authJsonBytes)
        {
            var token = SelectStringFromAuth(authJsonBytes,
                "tokens.access_token",
                "tokens.accessToken",
                "access_token",
                "accessToken");
            var payload = TryReadJwtPayload(token);
            var exp = payload?["exp"]?.Value<long?>();
            return exp.HasValue ? DateTimeOffset.FromUnixTimeSeconds(exp.Value).UtcDateTime : (DateTime?)null;
        }

        public static string ParseAccountEmail(byte[] authJsonBytes)
        {
            var directEmail = SelectStringFromAuth(authJsonBytes, "email", "profile.email", "account.email");
            if (!string.IsNullOrWhiteSpace(directEmail))
            {
                return directEmail.Trim();
            }

            var idToken = SelectStringFromAuth(authJsonBytes,
                "tokens.id_token",
                "tokens.idToken",
                "id_token",
                "idToken");
            var payload = TryReadJwtPayload(idToken);
            var email = payload?["email"]?.Value<string>();
            if (!string.IsNullOrWhiteSpace(email))
            {
                return email.Trim();
            }

            var accessToken = SelectStringFromAuth(authJsonBytes,
                "tokens.access_token",
                "tokens.accessToken",
                "access_token",
                "accessToken");
            var accessPayload = TryReadJwtPayload(accessToken);
            email = (accessPayload?["https://api.openai.com/profile"] as JObject)?["email"]?.Value<string>();
            if (!string.IsNullOrWhiteSpace(email))
            {
                return email.Trim();
            }

            var accountId = SelectStringFromAuth(authJsonBytes, "account_id", "accountId", "account.id");
            return accountId ?? string.Empty;
        }

        public static string ComputeStatus(DateTime? accessExp)
        {
            if (!accessExp.HasValue)
            {
                return StatusUnknown;
            }

            var now = DateTime.UtcNow;
            if (accessExp.Value <= now)
            {
                return StatusExpired;
            }

            return accessExp.Value - now < TimeSpan.FromDays(7) ? StatusWarn : StatusOk;
        }

        public static string MaskEmail(string email)
        {
            if (string.IsNullOrWhiteSpace(email))
            {
                return string.Empty;
            }

            var trimmed = email.Trim();
            var atIndex = trimmed.IndexOf('@');
            if (atIndex <= 0)
            {
                return trimmed.Length <= 3 ? trimmed.Substring(0, 1) + "***" : trimmed.Substring(0, 3) + "***";
            }

            var local = trimmed.Substring(0, atIndex);
            var domain = trimmed.Substring(atIndex);
            var prefix = local.Length <= 1 ? local.Substring(0, 1) : local.Substring(0, Math.Min(3, local.Length));
            return prefix + "***" + domain;
        }

        public static CodexProfilesFile CreateEmptyProfilesFile()
        {
            return new CodexProfilesFile
            {
                schemaVersion = CurrentSchemaVersion,
                machineName = Environment.MachineName,
                createdAtUtc = DateTime.UtcNow,
                items = new List<CodexProfileItem>()
            };
        }

        public static async Task<CodexProfileSourceFiles> ReadCurrentCodexFilesAsync(CancellationToken ct)
        {
            var configPath = Path.Combine(CodexFolderPath, CodexConfigProfileService.ConfigFileName);
            var authPath = Path.Combine(CodexFolderPath, CodexConfigProfileService.AuthFileName);
            if (!File.Exists(configPath) || !File.Exists(authPath))
            {
                throw new FileNotFoundException("未检测到 ~/.codex/auth.json，请先在 Codex CLI 完成一次登录。");
            }

            return new CodexProfileSourceFiles
            {
                SourceFolderPath = CodexFolderPath,
                ConfigTomlBytes = await ReadAllBytesAsync(configPath, ct).ConfigureAwait(false),
                AuthJsonBytes = await ReadAllBytesAsync(authPath, ct).ConfigureAwait(false)
            };
        }

        public static int CountConfigBaseUrlAssignments(byte[] configTomlBytes)
        {
            var text = configTomlBytes == null ? string.Empty : Encoding.UTF8.GetString(configTomlBytes);
            return FindBaseUrlAssignmentLineIndices(SplitTextPreservingLineEndings(text)).Count;
        }

        public static CodexConfigTemplateMergeResult MergeConfigTomlPreservingTargetBaseUrl(byte[] templateConfigTomlBytes, byte[] targetConfigTomlBytes)
        {
            var templateText = templateConfigTomlBytes == null ? string.Empty : Encoding.UTF8.GetString(templateConfigTomlBytes);
            var targetText = targetConfigTomlBytes == null ? string.Empty : Encoding.UTF8.GetString(targetConfigTomlBytes);
            var templateLines = SplitTextPreservingLineEndings(templateText);
            var targetLines = SplitTextPreservingLineEndings(targetText);
            var templateBaseUrlLines = FindBaseUrlAssignmentLineIndices(templateLines);
            var targetBaseUrlLines = FindBaseUrlAssignmentLineIndices(targetLines);

            var result = new CodexConfigTemplateMergeResult
            {
                TemplateBaseUrlCount = templateBaseUrlLines.Count,
                TargetBaseUrlCount = targetBaseUrlLines.Count
            };

            if (templateBaseUrlLines.Count == 0)
            {
                result.Message = "模板 config.toml 未找到 base_url 行。";
                return result;
            }

            if (targetBaseUrlLines.Count == 0)
            {
                result.Message = "目标 config.toml 未找到 base_url 行。";
                return result;
            }

            if (templateBaseUrlLines.Count != targetBaseUrlLines.Count)
            {
                result.Message = $"base_url 行数量不一致：模板 {templateBaseUrlLines.Count} 行，目标 {targetBaseUrlLines.Count} 行。";
                return result;
            }

            for (var i = 0; i < templateBaseUrlLines.Count; i++)
            {
                templateLines[templateBaseUrlLines[i]].Text = targetLines[targetBaseUrlLines[i]].Text;
            }

            result.Success = true;
            result.Message = $"已保留 {targetBaseUrlLines.Count} 行 base_url。";
            result.MergedConfigTomlBytes = new UTF8Encoding(false).GetBytes(JoinTextLines(templateLines));
            return result;
        }

        private static async Task SaveCoreAsync(CodexProfilesFile file, CancellationToken ct)
        {
            NormalizeProfilesFile(file);
            Directory.CreateDirectory(RootFolderPath);
            var json = JsonConvert.SerializeObject(file, Formatting.Indented);
            var protectedText = CodexConfigProfileService.ProtectBytesToBase64(Encoding.UTF8.GetBytes(json));
            await WriteAllTextAsync(ProfilesFilePath, protectedText, ct).ConfigureAwait(false);
        }

        private static CodexPortableProfilesFile CreatePortableExportFile(CodexProfilesFile file)
        {
            NormalizeProfilesFile(file);
            var portableFile = new CodexPortableProfilesFile
            {
                schemaVersion = CurrentSchemaVersion,
                packageKind = ExportPortableKind,
                machineName = string.IsNullOrWhiteSpace(file.machineName) ? Environment.MachineName : file.machineName,
                createdAtUtc = file.createdAtUtc == default(DateTime) ? DateTime.UtcNow : file.createdAtUtc,
                items = new List<CodexPortableProfileItem>()
            };

            foreach (var item in file.items.Where(item => item != null))
            {
                NormalizeProfileItem(item);
                var configTomlBytes = UnprotectProfileContentForExport(item.ProtectedConfigTomlBase64 ?? item.ConfigTomlContentProtected, item.DisplayName, "config.toml");
                var authJsonBytes = UnprotectProfileContentForExport(item.ProtectedAuthJsonBase64 ?? item.AuthJsonContentProtected, item.DisplayName, "auth.json");
                portableFile.items.Add(new CodexPortableProfileItem
                {
                    DisplayName = item.DisplayName,
                    Name = item.DisplayName,
                    AccountEmail = item.AccountEmail ?? string.Empty,
                    Note = item.Note ?? string.Empty,
                    Remark = item.Note ?? string.Empty,
                    Tags = item.Tags ?? string.Empty,
                    FolderPath = item.FolderPath ?? string.Empty,
                    LastAppliedAt = item.LastAppliedAt,
                    LastImportedAt = item.LastImportedAt == default(DateTime) ? DateTime.UtcNow : item.LastImportedAt,
                    AccessTokenExpiresAt = item.AccessTokenExpiresAt ?? ParseAccessTokenExp(authJsonBytes),
                    RefreshTokenExpiresAt = null,
                    Status = item.Status ?? ComputeStatus(item.AccessTokenExpiresAt ?? ParseAccessTokenExp(authJsonBytes)),
                    ConfigTomlBase64 = Convert.ToBase64String(configTomlBytes),
                    AuthJsonBase64 = Convert.ToBase64String(authJsonBytes),
                    EnableRotation = item.EnableRotation,
                    RotationPriority = item.RotationPriority
                });
            }

            return portableFile;
        }

        private static byte[] UnprotectProfileContentForExport(string protectedBase64, string displayName, string fileName)
        {
            try
            {
                var bytes = CodexConfigProfileService.UnprotectBytesFromBase64(protectedBase64);
                if (bytes == null || bytes.Length == 0)
                {
                    throw new InvalidOperationException("内容为空。");
                }

                return bytes;
            }
            catch (Exception ex)
            {
                var name = string.IsNullOrWhiteSpace(displayName) ? "Codex 账号" : displayName;
                throw new InvalidOperationException($"档案「{name}」的 {fileName} 无法用当前 Windows 用户解密，不能导出为可迁移加密包。请在能正常切换该档案的 Windows 用户下重新导入当前账号后再导出。", ex);
            }
        }

        private static CodexProfilesFile DeserializeImportPackage(byte[] plaintext)
        {
            var json = Encoding.UTF8.GetString(plaintext);
            var root = JObject.Parse(json);
            var packageKind = root["packageKind"]?.Value<string>() ?? string.Empty;
            if (string.Equals(packageKind, ExportPortableKind, StringComparison.Ordinal)
                || HasPortableContent(root))
            {
                return ConvertPortableImportPackage(root);
            }

            var legacyFile = JsonConvert.DeserializeObject<CodexProfilesFile>(json) ?? CreateEmptyProfilesFile();
            NormalizeProfilesFile(legacyFile);
            return ReprotectLegacyImportPackage(legacyFile);
        }

        private static bool HasPortableContent(JObject root)
        {
            var items = root["items"] as JArray;
            if (items == null)
            {
                return false;
            }

            return items.OfType<JObject>().Any(item =>
                item["ConfigTomlBase64"] != null
                || item["AuthJsonBase64"] != null
                || item["configTomlBase64"] != null
                || item["authJsonBase64"] != null);
        }

        private static CodexProfilesFile ConvertPortableImportPackage(JObject root)
        {
            var portableFile = root.ToObject<CodexPortableProfilesFile>() ?? new CodexPortableProfilesFile();
            var file = new CodexProfilesFile
            {
                schemaVersion = CurrentSchemaVersion,
                machineName = string.IsNullOrWhiteSpace(portableFile.machineName) ? Environment.MachineName : portableFile.machineName,
                createdAtUtc = portableFile.createdAtUtc == default(DateTime) ? DateTime.UtcNow : portableFile.createdAtUtc,
                items = new List<CodexProfileItem>()
            };

            foreach (var imported in portableFile.items ?? new List<CodexPortableProfileItem>())
            {
                if (imported == null)
                {
                    continue;
                }

                byte[] configTomlBytes;
                byte[] authJsonBytes;
                try
                {
                    configTomlBytes = Convert.FromBase64String(imported.ConfigTomlBase64 ?? string.Empty);
                    authJsonBytes = Convert.FromBase64String(imported.AuthJsonBase64 ?? string.Empty);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("Codex 加密档案包内的配置内容格式无效。", ex);
                }

                if (configTomlBytes.Length == 0 || authJsonBytes.Length == 0)
                {
                    continue;
                }

                var protectedConfig = CodexConfigProfileService.ProtectBytesToBase64(configTomlBytes);
                var protectedAuth = CodexConfigProfileService.ProtectBytesToBase64(authJsonBytes);
                var accessExp = imported.AccessTokenExpiresAt ?? ParseAccessTokenExp(authJsonBytes);
                var item = new CodexProfileItem
                {
                    DisplayName = string.IsNullOrWhiteSpace(imported.DisplayName) ? "Codex 账号" : imported.DisplayName,
                    Name = string.IsNullOrWhiteSpace(imported.DisplayName) ? "Codex 账号" : imported.DisplayName,
                    AccountEmail = string.IsNullOrWhiteSpace(imported.AccountEmail) ? ParseAccountEmail(authJsonBytes) : imported.AccountEmail,
                    Note = imported.Note ?? string.Empty,
                    Remark = imported.Note ?? string.Empty,
                    Tags = imported.Tags ?? string.Empty,
                    FolderPath = imported.FolderPath ?? string.Empty,
                    LastAppliedAt = imported.LastAppliedAt,
                    LastImportedAt = imported.LastImportedAt == default(DateTime) ? DateTime.UtcNow : imported.LastImportedAt,
                    AccessTokenExpiresAt = accessExp,
                    RefreshTokenExpiresAt = null,
                    ProtectedConfigTomlBase64 = protectedConfig,
                    ProtectedAuthJsonBase64 = protectedAuth,
                    ConfigTomlContentProtected = protectedConfig,
                    AuthJsonContentProtected = protectedAuth,
                    Status = ComputeStatus(accessExp),
                    EnableRotation = imported.EnableRotation,
                    RotationPriority = imported.RotationPriority
                };
                file.items.Add(item);
            }

            NormalizeProfilesFile(file);
            return file;
        }

        private static CodexProfilesFile ReprotectLegacyImportPackage(CodexProfilesFile file)
        {
            foreach (var item in file.items.Where(item => item != null))
            {
                byte[] configTomlBytes;
                byte[] authJsonBytes;
                try
                {
                    configTomlBytes = CodexConfigProfileService.UnprotectBytesFromBase64(item.ProtectedConfigTomlBase64 ?? item.ConfigTomlContentProtected);
                    authJsonBytes = CodexConfigProfileService.UnprotectBytesFromBase64(item.ProtectedAuthJsonBase64 ?? item.AuthJsonContentProtected);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("这个 .codexbox 是旧版本导出的，包内仍包含源 Windows 用户的 DPAPI 加密内容，当前 Windows 用户无法解密。请在源电脑或源 Windows 用户下使用新版 MyTools 重新导出加密包后再导入。", ex);
                }

                if (configTomlBytes == null || authJsonBytes == null)
                {
                    continue;
                }

                var protectedConfig = CodexConfigProfileService.ProtectBytesToBase64(configTomlBytes);
                var protectedAuth = CodexConfigProfileService.ProtectBytesToBase64(authJsonBytes);
                item.ProtectedConfigTomlBase64 = protectedConfig;
                item.ProtectedAuthJsonBase64 = protectedAuth;
                item.ConfigTomlContentProtected = protectedConfig;
                item.AuthJsonContentProtected = protectedAuth;
                item.AccountEmail = string.IsNullOrWhiteSpace(item.AccountEmail) ? ParseAccountEmail(authJsonBytes) : item.AccountEmail;
                item.AccessTokenExpiresAt = item.AccessTokenExpiresAt ?? ParseAccessTokenExp(authJsonBytes);
                item.RefreshTokenExpiresAt = null;
                item.Status = ComputeStatus(item.AccessTokenExpiresAt);
            }

            NormalizeProfilesFile(file);
            return file;
        }

        private static List<CodexConfigTextLine> SplitTextPreservingLineEndings(string text)
        {
            var value = text ?? string.Empty;
            var lines = new List<CodexConfigTextLine>();
            var lineStart = 0;
            var index = 0;
            while (index < value.Length)
            {
                var ch = value[index];
                if (ch != '\r' && ch != '\n')
                {
                    index++;
                    continue;
                }

                string lineEnding;
                if (ch == '\r' && index + 1 < value.Length && value[index + 1] == '\n')
                {
                    lineEnding = "\r\n";
                    lines.Add(new CodexConfigTextLine(value.Substring(lineStart, index - lineStart), lineEnding));
                    index += 2;
                }
                else
                {
                    lineEnding = ch == '\r' ? "\r" : "\n";
                    lines.Add(new CodexConfigTextLine(value.Substring(lineStart, index - lineStart), lineEnding));
                    index++;
                }

                lineStart = index;
            }

            if (lineStart < value.Length || value.Length == 0)
            {
                lines.Add(new CodexConfigTextLine(value.Substring(lineStart), string.Empty));
            }

            return lines;
        }

        private static List<int> FindBaseUrlAssignmentLineIndices(IList<CodexConfigTextLine> lines)
        {
            var result = new List<int>();
            if (lines == null)
            {
                return result;
            }

            for (var i = 0; i < lines.Count; i++)
            {
                if (IsBaseUrlAssignmentLine(lines[i]?.Text))
                {
                    result.Add(i);
                }
            }

            return result;
        }

        private static bool IsBaseUrlAssignmentLine(string line)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                return false;
            }

            var index = 0;
            if (line.Length > 0 && line[0] == '\ufeff')
            {
                index++;
            }

            while (index < line.Length && char.IsWhiteSpace(line[index]))
            {
                index++;
            }

            if (index >= line.Length || line[index] == '#')
            {
                return false;
            }

            const string key = "base_url";
            if (index + key.Length > line.Length || !string.Equals(line.Substring(index, key.Length), key, StringComparison.Ordinal))
            {
                return false;
            }

            index += key.Length;
            while (index < line.Length && char.IsWhiteSpace(line[index]))
            {
                index++;
            }

            return index < line.Length && line[index] == '=';
        }

        private static string JoinTextLines(IEnumerable<CodexConfigTextLine> lines)
        {
            var builder = new StringBuilder();
            foreach (var line in lines ?? Enumerable.Empty<CodexConfigTextLine>())
            {
                builder.Append(line?.Text ?? string.Empty);
                builder.Append(line?.LineEnding ?? string.Empty);
            }

            return builder.ToString();
        }

        private static CodexProfilesFile DeserializeProfilesFile(byte[] jsonBytes)
        {
            if (jsonBytes == null || jsonBytes.Length == 0)
            {
                return CreateEmptyProfilesFile();
            }

            var json = Encoding.UTF8.GetString(jsonBytes);
            return JsonConvert.DeserializeObject<CodexProfilesFile>(json) ?? CreateEmptyProfilesFile();
        }

        private static void NormalizeProfilesFile(CodexProfilesFile file)
        {
            if (file == null)
            {
                return;
            }

            file.schemaVersion = CurrentSchemaVersion;
            file.machineName = string.IsNullOrWhiteSpace(file.machineName) ? Environment.MachineName : file.machineName;
            if (file.createdAtUtc == default(DateTime))
            {
                file.createdAtUtc = DateTime.UtcNow;
            }

            if (file.items == null)
            {
                file.items = new List<CodexProfileItem>();
            }

            foreach (var item in file.items.Where(item => item != null))
            {
                NormalizeProfileItem(item);
            }
        }

        private static void NormalizeProfileItem(CodexProfileItem item)
        {
            if (item == null)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(item.DisplayName))
            {
                item.DisplayName = string.IsNullOrWhiteSpace(item.Name) ? "Codex 账号" : item.Name;
            }

            item.Name = item.DisplayName;
            if (string.IsNullOrWhiteSpace(item.Note))
            {
                item.Note = item.Remark ?? string.Empty;
            }

            item.Remark = string.IsNullOrWhiteSpace(item.Note) ? item.DisplayName : item.Note;
            if (item.LastImportedAt == default(DateTime))
            {
                item.LastImportedAt = item.LastAppliedAt ?? DateTime.UtcNow;
            }

            if (string.IsNullOrWhiteSpace(item.ProtectedConfigTomlBase64))
            {
                item.ProtectedConfigTomlBase64 = item.ConfigTomlContentProtected ?? string.Empty;
            }

            if (string.IsNullOrWhiteSpace(item.ProtectedAuthJsonBase64))
            {
                item.ProtectedAuthJsonBase64 = item.AuthJsonContentProtected ?? string.Empty;
            }

            item.ConfigTomlContentProtected = item.ProtectedConfigTomlBase64;
            item.AuthJsonContentProtected = item.ProtectedAuthJsonBase64;
            item.RefreshTokenExpiresAt = null;
            item.Status = ComputeStatus(item.AccessTokenExpiresAt);
            item.EnableRotation = item.EnableRotation;
            item.RotationPriority = item.RotationPriority;
        }

        private static async Task<CodexProfilesFile> TryLoadLegacyProfilesAsync(CancellationToken ct)
        {
            var file = CreateEmptyProfilesFile();
            var settings = await AppSettingsService.LoadAsync().ConfigureAwait(false);
            foreach (var profile in settings.CodexProfiles ?? new List<CodexProfileSettings>())
            {
                if (profile == null)
                {
                    continue;
                }

                var item = new CodexProfileItem
                {
                    DisplayName = string.IsNullOrWhiteSpace(profile.Name) ? "Codex 账号" : profile.Name,
                    Name = string.IsNullOrWhiteSpace(profile.Name) ? "Codex 账号" : profile.Name,
                    Note = profile.Remark ?? string.Empty,
                    Remark = profile.Remark ?? string.Empty,
                    Tags = profile.Tags ?? string.Empty,
                    FolderPath = profile.FolderPath ?? string.Empty,
                    LastAppliedAt = profile.LastAppliedAt,
                    LastImportedAt = profile.LastAppliedAt ?? DateTime.UtcNow,
                    ProtectedConfigTomlBase64 = profile.ConfigTomlContentProtected ?? string.Empty,
                    ProtectedAuthJsonBase64 = profile.AuthJsonContentProtected ?? string.Empty,
                    ConfigTomlContentProtected = profile.ConfigTomlContentProtected ?? string.Empty,
                    AuthJsonContentProtected = profile.AuthJsonContentProtected ?? string.Empty,
                    EnableRotation = false,
                    RotationPriority = 0
                };

                var authBytes = CodexConfigProfileService.UnprotectBytesFromBase64(item.ProtectedAuthJsonBase64);
                item.AccountEmail = ParseAccountEmail(authBytes);
                item.AccessTokenExpiresAt = ParseAccessTokenExp(authBytes);
                item.RefreshTokenExpiresAt = null;
                item.Status = ComputeStatus(item.AccessTokenExpiresAt);
                file.items.Add(item);
            }

            return file;
        }

        private static string SelectStringFromAuth(byte[] authJsonBytes, params string[] paths)
        {
            if (authJsonBytes == null || authJsonBytes.Length == 0)
            {
                return string.Empty;
            }

            try
            {
                var root = JObject.Parse(Encoding.UTF8.GetString(authJsonBytes));
                foreach (var path in paths)
                {
                    var value = root.SelectToken(path)?.Value<string>();
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        return value.Trim();
                    }
                }
            }
            catch
            {
            }

            return string.Empty;
        }

        private static JObject TryReadJwtPayload(string jwt)
        {
            if (string.IsNullOrWhiteSpace(jwt))
            {
                return null;
            }

            try
            {
                var parts = jwt.Split('.');
                if (parts.Length < 2)
                {
                    return null;
                }

                var payload = parts[1].Replace('-', '+').Replace('_', '/');
                switch (payload.Length % 4)
                {
                    case 2:
                        payload += "==";
                        break;
                    case 3:
                        payload += "=";
                        break;
                    case 1:
                        return null;
                }

                return JObject.Parse(Encoding.UTF8.GetString(Convert.FromBase64String(payload)));
            }
            catch
            {
                return null;
            }
        }

        private static byte[] DeriveExportKey(string password, byte[] salt)
        {
            using (var pbkdf2 = new Rfc2898DeriveBytes(password, salt, ExportIterations, HashAlgorithmName.SHA256))
            {
                return pbkdf2.GetBytes(64);
            }
        }

        private static byte[] EncryptAesCbc(byte[] plaintext, byte[] key, byte[] iv)
        {
            using (var aes = Aes.Create())
            {
                aes.KeySize = 256;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;
                aes.Key = key;
                aes.IV = iv;
                using (var output = new MemoryStream())
                using (var crypto = new CryptoStream(output, aes.CreateEncryptor(), CryptoStreamMode.Write))
                {
                    crypto.Write(plaintext, 0, plaintext.Length);
                    crypto.FlushFinalBlock();
                    return output.ToArray();
                }
            }
        }

        private static byte[] DecryptAesCbc(byte[] ciphertext, byte[] key, byte[] iv)
        {
            using (var aes = Aes.Create())
            {
                aes.KeySize = 256;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;
                aes.Key = key;
                aes.IV = iv;
                using (var input = new MemoryStream(ciphertext))
                using (var crypto = new CryptoStream(input, aes.CreateDecryptor(), CryptoStreamMode.Read))
                using (var output = new MemoryStream())
                {
                    crypto.CopyTo(output);
                    return output.ToArray();
                }
            }
        }

        private static bool ConstantTimeEquals(byte[] left, byte[] right)
        {
            if (left == null || right == null || left.Length != right.Length)
            {
                return false;
            }

            var diff = 0;
            for (var i = 0; i < left.Length; i++)
            {
                diff |= left[i] ^ right[i];
            }

            return diff == 0;
        }

        private static byte[] CreateRandomBytes(int length)
        {
            var bytes = new byte[length];
            using (var rng = RandomNumberGenerator.Create())
            {
                rng.GetBytes(bytes);
            }

            return bytes;
        }

        private static string SanitizeFileName(string value)
        {
            var result = string.IsNullOrWhiteSpace(value) ? "codex" : value.Trim();
            foreach (var invalid in Path.GetInvalidFileNameChars())
            {
                result = result.Replace(invalid, '_');
            }

            return string.IsNullOrWhiteSpace(result) ? "codex" : result;
        }

        private static async Task<byte[]> ReadAllBytesAsync(string filePath, CancellationToken ct)
        {
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, true))
            using (var memory = new MemoryStream())
            {
                await stream.CopyToAsync(memory, 81920, ct).ConfigureAwait(false);
                return memory.ToArray();
            }
        }

        private static async Task WriteAllBytesAsync(string filePath, byte[] bytes, CancellationToken ct)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(filePath) ?? RootFolderPath);
            using (var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None, 81920, true))
            {
                await stream.WriteAsync(bytes, 0, bytes.Length, ct).ConfigureAwait(false);
            }
        }

        private static async Task<string> ReadAllTextAsync(string filePath, CancellationToken ct)
        {
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, true))
            using (var reader = new StreamReader(stream, Encoding.UTF8))
            {
                return await reader.ReadToEndAsync().ConfigureAwait(false);
            }
        }

        private static async Task WriteAllTextAsync(string filePath, string text, CancellationToken ct)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(filePath) ?? RootFolderPath);
            using (var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None, 81920, true))
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
            {
                await writer.WriteAsync(text ?? string.Empty).ConfigureAwait(false);
            }
        }
    }

    public class CodexProfilesFile
    {
        public int schemaVersion { get; set; }
        public string machineName { get; set; }
        public DateTime createdAtUtc { get; set; }
        public List<CodexProfileItem> items { get; set; } = new List<CodexProfileItem>();
    }

    public class CodexPortableProfilesFile
    {
        public int schemaVersion { get; set; }
        public string packageKind { get; set; }
        public string machineName { get; set; }
        public DateTime createdAtUtc { get; set; }
        public List<CodexPortableProfileItem> items { get; set; } = new List<CodexPortableProfileItem>();
    }

    public class CodexPortableProfileItem
    {
        public string DisplayName { get; set; }
        public string Name { get; set; }
        public string AccountEmail { get; set; }
        public string Note { get; set; }
        public string Remark { get; set; }
        public string Tags { get; set; }
        public string FolderPath { get; set; }
        public DateTime? LastAppliedAt { get; set; }
        public DateTime LastImportedAt { get; set; }
        public DateTime? AccessTokenExpiresAt { get; set; }
        public DateTime? RefreshTokenExpiresAt { get; set; }
        public string Status { get; set; }
        public string ConfigTomlBase64 { get; set; }
        public string AuthJsonBase64 { get; set; }
        public bool EnableRotation { get; set; }
        public int RotationPriority { get; set; }
    }

    public class CodexActiveFile
    {
        public string ActiveDisplayName { get; set; }
        public DateTime SwitchedAtUtc { get; set; }
    }

    public class CodexProfileSourceFiles
    {
        public string SourceFolderPath { get; set; }
        public byte[] ConfigTomlBytes { get; set; }
        public byte[] AuthJsonBytes { get; set; }
    }

    public class CodexProfileBackupPackage
    {
        public string ActiveDisplayName { get; set; }
        public DateTime CreatedAtUtc { get; set; }
        public string ConfigTomlBase64 { get; set; }
        public string AuthJsonBase64 { get; set; }
    }

    public class CodexProfilesLibraryBackupPackage
    {
        public string BackupKind { get; set; }
        public string TemplateDisplayName { get; set; }
        public DateTime CreatedAtUtc { get; set; }
        public CodexProfilesFile Profiles { get; set; }
    }

    public class CodexConfigTemplateMergeResult
    {
        public bool Success { get; set; }
        public byte[] MergedConfigTomlBytes { get; set; }
        public int TemplateBaseUrlCount { get; set; }
        public int TargetBaseUrlCount { get; set; }
        public string Message { get; set; }
    }

    internal sealed class CodexConfigTextLine
    {
        public CodexConfigTextLine(string text, string lineEnding)
        {
            Text = text ?? string.Empty;
            LineEnding = lineEnding ?? string.Empty;
        }

        public string Text { get; set; }
        public string LineEnding { get; }
    }
}
