using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Newtonsoft.Json;

namespace ppt_copier_addin
{
    public partial class ThisAddIn
    {
        private string configFilePath;
        private string stateFilePath;
        private string logFolderPath;
        private string fallbackBackupPath;
        private Dictionary<string, List<string>> fileState;
        private readonly object stateLock = new object();
        private string currentDateFolder;  // 记录当前日期文件夹
        private DateTime lastDateCheck;    // 上次日期检查时间

        // 默认日志保留天数
        private const int DefaultLogRetentionDays = 7;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                // 设置配置文件路径
                string appDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "PowerPointAutoCopy");
                configFilePath = Path.Combine(appDataFolder, "config.json");
                stateFilePath = Path.Combine(appDataFolder, "state.json");
                logFolderPath = Path.Combine(appDataFolder, "Logs");
                fallbackBackupPath = Path.Combine(appDataFolder, "Backup");

                // 确保目录存在
                Directory.CreateDirectory(appDataFolder);
                Directory.CreateDirectory(logFolderPath);

                // 初始化日期跟踪
                currentDateFolder = DateTime.Now.ToString("yyyy-MM-dd");
                lastDateCheck = DateTime.Now;

                // 创建默认配置文件（如果不存在）
                CreateDefaultConfigIfNotExists();

                // 读取配置，获取日志保留天数
                int logRetentionDays = GetLogRetentionDays();

                // 清理旧日志（只在启动时执行一次）
                CleanOldLogs(logRetentionDays);

                // 注册事件
                this.Application.AfterPresentationOpen += Application_AfterPresentationOpen;
                this.Application.ProtectedViewWindowOpen += Application_ProtectedViewWindowOpen;

                // 加载状态文件
                LoadStateFile();

                // 清理过期的状态记录
                CleanExpiredStateRecords(logRetentionDays);

                LogMessage("PowerPoint自动复制插件已启动");
                LogMessage($"日志保留天数: {logRetentionDays}天，已清理过期日志");
                LogMessage($"回退备份路径: {fallbackBackupPath}");
                LogMessage($"当前日期文件夹: {currentDateFolder}");
            }
            catch (Exception ex)
            {
                LogMessage($"插件启动失败: {ex.Message}");
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                // 保存状态文件
                SaveStateFile();

                // 注销事件
                this.Application.AfterPresentationOpen -= Application_AfterPresentationOpen;
                this.Application.ProtectedViewWindowOpen -= Application_ProtectedViewWindowOpen;

                LogMessage("PowerPoint自动复制插件已关闭");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"插件关闭失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 检查并更新日期文件夹（跨天处理）
        /// </summary>
        private void CheckAndUpdateDateFolder()
        {
            DateTime now = DateTime.Now;

            // 如果日期改变了
            if (now.Date != lastDateCheck.Date)
            {
                string newDateFolder = now.ToString("yyyy-MM-dd");
                LogMessage($"检测到日期变更: {currentDateFolder} -> {newDateFolder}");
                currentDateFolder = newDateFolder;
                lastDateCheck = now;
            }
        }

        /// <summary>
        /// 清理过期的状态记录
        /// </summary>
        private void CleanExpiredStateRecords(int retentionDays)
        {
            try
            {
                lock (stateLock)
                {
                    if (fileState == null || fileState.Count == 0)
                        return;

                    DateTime cutoffDate = DateTime.Now.AddDays(-retentionDays);
                    var keysToRemove = new List<string>();

                    foreach (var key in fileState.Keys)
                    {
                        // 从key中提取日期部分（格式：yyyy-MM-dd_filename）
                        string datePart = key.Split('_')[0];
                        if (DateTime.TryParse(datePart, out DateTime fileDate))
                        {
                            if (fileDate < cutoffDate)
                            {
                                keysToRemove.Add(key);
                            }
                        }
                    }

                    foreach (var key in keysToRemove)
                    {
                        fileState.Remove(key);
                        LogMessage($"移除过期状态记录: {key}");
                    }

                    if (keysToRemove.Count > 0)
                    {
                        SaveStateFile();
                        LogMessage($"已清理 {keysToRemove.Count} 条过期状态记录");
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"清理过期状态记录失败: {ex.Message}");
            }
        }

        private void CreateDefaultConfigIfNotExists()
        {
            if (!File.Exists(configFilePath))
            {
                var defaultConfig = new Config
                {
                    TargetCopyPath = @"D:\Backup",
                    EnableLogging = true,
                    EnableAutoCopy = true,
                    DateFolderFormat = "yyyy-MM-dd",
                    LogRetentionDays = 7,
                    UseFallbackOnError = true
                };

                string json = JsonConvert.SerializeObject(defaultConfig, Formatting.Indented);
                File.WriteAllText(configFilePath, json);

                System.Diagnostics.Debug.WriteLine($"已创建默认配置文件，目标路径: D:\\Backup，日志保留天数: 7天");
            }
        }

        /// <summary>
        /// 获取日志保留天数
        /// </summary>
        private int GetLogRetentionDays()
        {
            try
            {
                if (File.Exists(configFilePath))
                {
                    string json = File.ReadAllText(configFilePath);
                    var config = JsonConvert.DeserializeObject<Config>(json);
                    if (config != null && config.LogRetentionDays > 0)
                    {
                        return config.LogRetentionDays;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"读取日志保留天数配置失败: {ex.Message}");
            }
            return DefaultLogRetentionDays;
        }

        /// <summary>
        /// 清理旧的日志文件（只在启动时调用）
        /// </summary>
        /// <param name="retentionDays">保留天数</param>
        private void CleanOldLogs(int retentionDays)
        {
            try
            {
                if (!Directory.Exists(logFolderPath))
                    return;

                DateTime cutoffDate = DateTime.Now.AddDays(-retentionDays);
                var logFiles = Directory.GetFiles(logFolderPath, "PowerPointAutoCopy_*.log");

                int deletedCount = 0;
                foreach (var logFile in logFiles)
                {
                    try
                    {
                        FileInfo fileInfo = new FileInfo(logFile);
                        if (fileInfo.CreationTime < cutoffDate)
                        {
                            File.Delete(logFile);
                            deletedCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"删除日志文件失败 {logFile}: {ex.Message}");
                    }
                }

                if (deletedCount > 0)
                {
                    System.Diagnostics.Debug.WriteLine($"已清理 {deletedCount} 个旧日志文件（超过{retentionDays}天）");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"清理旧日志失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 验证并获取可用的目标路径
        /// </summary>
        private string GetAvailableTargetPath(string configuredPath, bool useFallback)
        {
            // 首先检查配置的路径
            if (!string.IsNullOrEmpty(configuredPath))
            {
                try
                {
                    // 测试写入权限
                    string testFile = Path.Combine(configuredPath, ".write_test");
                    Directory.CreateDirectory(configuredPath);
                    File.WriteAllText(testFile, "test");
                    File.Delete(testFile);

                    LogMessage($"目标路径可用: {configuredPath}");
                    return configuredPath;
                }
                catch (Exception ex)
                {
                    LogMessage($"目标路径不可写入: {configuredPath}, 错误: {ex.Message}");
                }
            }

            // 如果配置路径不可用且允许回退，则使用回退路径
            if (useFallback)
            {
                try
                {
                    Directory.CreateDirectory(fallbackBackupPath);
                    string testFile = Path.Combine(fallbackBackupPath, ".write_test");
                    File.WriteAllText(testFile, "test");
                    File.Delete(testFile);

                    LogMessage($"使用回退路径: {fallbackBackupPath}");
                    return fallbackBackupPath;
                }
                catch (Exception ex)
                {
                    LogMessage($"回退路径也不可写入: {fallbackBackupPath}, 错误: {ex.Message}");
                }
            }

            return null;
        }

        private async void Application_AfterPresentationOpen(PowerPoint.Presentation Pres)
        {
            await HandlePresentationOpen(Pres, "Normal");
        }

        private async void Application_ProtectedViewWindowOpen(PowerPoint.ProtectedViewWindow ProtectedViewWindow)
        {
            try
            {
                // 从保护视图中获取 Presentation 对象
                PowerPoint.Presentation presentation = ProtectedViewWindow.Edit();
                await HandlePresentationOpen(presentation, "Protected");
            }
            catch (Exception ex)
            {
                LogMessage($"处理保护视图文档打开事件时出错: {ex.Message}");
            }
        }

        private async Task HandlePresentationOpen(PowerPoint.Presentation presentation, string openType)
        {
            try
            {
                // 检查并更新日期文件夹（跨天处理）
                CheckAndUpdateDateFolder();

                // 读取配置
                var config = await ReadConfigAsync();
                if (config == null)
                {
                    LogMessage("配置文件不存在");
                    return;
                }

                // 检查是否启用自动复制
                if (!config.EnableAutoCopy)
                {
                    LogMessage("自动复制功能已禁用");
                    return;
                }

                // 获取当前文档路径
                string documentPath = presentation.FullName;
                if (string.IsNullOrEmpty(documentPath))
                {
                    LogMessage("文档尚未保存，无法获取路径");
                    return;
                }

                LogMessage($"检测到文档打开: {documentPath} (打开方式: {openType})");

                // 检查是否为移动存储设备
                bool isRemovableDrive = await IsRemovableDriveAsync(documentPath);

                if (isRemovableDrive)
                {
                    LogMessage($"检测到移动存储设备文件，开始复制: {documentPath}");

                    // 获取可用的目标路径
                    string targetPath = GetAvailableTargetPath(config.TargetCopyPath, config.UseFallbackOnError);

                    if (string.IsNullOrEmpty(targetPath))
                    {
                        LogMessage("无法获取可用的目标路径，复制操作终止");
                        return;
                    }

                    // 执行复制操作
                    await CopyFileToTargetAsync(documentPath, targetPath);
                }
                else
                {
                    LogMessage($"文档不在移动存储设备上: {documentPath}");
                }
            }
            catch (Exception ex)
            {
                LogMessage($"处理文档打开事件时出错: {ex.Message}");
            }
        }

        private async Task<Config> ReadConfigAsync()
        {
            return await Task.Run(() =>
            {
                try
                {
                    if (File.Exists(configFilePath))
                    {
                        string json = File.ReadAllText(configFilePath);
                        return JsonConvert.DeserializeObject<Config>(json);
                    }
                }
                catch (Exception ex)
                {
                    LogMessage($"读取配置文件失败: {ex.Message}");
                }
                return null;
            });
        }

        private async Task<bool> IsRemovableDriveAsync(string filePath)
        {
            return await Task.Run(() =>
            {
                try
                {
                    string rootPath = Path.GetPathRoot(filePath);
                    if (string.IsNullOrEmpty(rootPath))
                        return false;

                    DriveInfo driveInfo = new DriveInfo(rootPath);
                    bool isRemovable = driveInfo.DriveType == DriveType.Removable;
                    LogMessage($"驱动器类型检测: {rootPath} - {driveInfo.DriveType}, 是否可移动: {isRemovable}");
                    return isRemovable;
                }
                catch (Exception ex)
                {
                    LogMessage($"检查移动存储设备失败: {ex.Message}");
                    return false;
                }
            });
        }

        private async Task CopyFileToTargetAsync(string sourcePath, string targetBasePath)
        {
            await Task.Run(async () =>
            {
                try
                {
                    // 确保目标目录存在
                    if (!Directory.Exists(targetBasePath))
                    {
                        Directory.CreateDirectory(targetBasePath);
                        LogMessage($"创建目标目录: {targetBasePath}");
                    }

                    // 获取文件信息
                    FileInfo sourceFile = new FileInfo(sourcePath);
                    string fileName = sourceFile.Name;
                    string fileHash = await ComputeFileHashAsync(sourcePath);

                    // 使用当前日期文件夹（支持跨天）
                    string targetFolder = Path.Combine(targetBasePath, currentDateFolder);
                    Directory.CreateDirectory(targetFolder);

                    string targetPath = Path.Combine(targetFolder, fileName);

                    // 检查是否需要复制（防重）
                    bool shouldCopy = await CheckFileDuplicateAsync(currentDateFolder, fileName, fileHash);

                    if (shouldCopy)
                    {
                        // 处理文件名冲突
                        targetPath = GetUniqueFilePath(targetPath);

                        // 复制文件，使用重试机制
                        await CopyFileWithRetryAsync(sourcePath, targetPath);

                        // 记录文件状态
                        await RecordFileStateAsync(currentDateFolder, fileName, fileHash, targetPath);

                        LogMessage($"文件复制成功: {sourcePath} -> {targetPath}");

                        // 如果使用的是回退路径，记录警告
                        if (targetBasePath == fallbackBackupPath)
                        {
                            LogMessage($"警告: 使用的是回退备份路径，请检查配置的目标路径是否可用");
                        }
                    }
                    else
                    {
                        LogMessage($"文件已存在，跳过复制: {fileName}");
                    }
                }
                catch (Exception ex)
                {
                    LogMessage($"复制文件失败: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// 带重试机制的文件复制
        /// </summary>
        private async Task CopyFileWithRetryAsync(string sourcePath, string targetPath, int maxRetries = 3)
        {
            int retryCount = 0;
            while (retryCount < maxRetries)
            {
                try
                {
                    File.Copy(sourcePath, targetPath, true);
                    return;
                }
                catch (IOException ex) when (retryCount < maxRetries - 1)
                {
                    retryCount++;
                    LogMessage($"文件复制失败，正在重试 ({retryCount}/{maxRetries}): {ex.Message}");
                    await Task.Delay(100 * retryCount); // 递增延迟
                }
            }

            // 最后一次尝试
            File.Copy(sourcePath, targetPath, true);
        }

        private async Task<string> ComputeFileHashAsync(string filePath)
        {
            return await Task.Run(() =>
            {
                try
                {
                    using (var md5 = MD5.Create())
                    {
                        using (var stream = File.OpenRead(filePath))
                        {
                            byte[] hash = md5.ComputeHash(stream);
                            return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogMessage($"计算文件哈希失败: {ex.Message}");
                    return string.Empty;
                }
            });
        }

        private async Task<bool> CheckFileDuplicateAsync(string dateFolder, string fileName, string fileHash)
        {
            return await Task.Run(() =>
            {
                lock (stateLock)
                {
                    if (fileState == null)
                        return true;

                    string key = $"{dateFolder}_{fileName}";

                    if (fileState.ContainsKey(key))
                    {
                        var existingHashes = fileState[key];
                        bool isDuplicate = existingHashes.Contains(fileHash);

                        if (isDuplicate)
                        {
                            LogMessage($"检测到重复文件: {fileName} (哈希值相同)");
                        }

                        return !isDuplicate;
                    }

                    return true;
                }
            });
        }

        private async Task RecordFileStateAsync(string dateFolder, string fileName, string fileHash, string targetPath)
        {
            await Task.Run(() =>
            {
                lock (stateLock)
                {
                    if (fileState == null)
                        fileState = new Dictionary<string, List<string>>();

                    string key = $"{dateFolder}_{fileName}";

                    if (!fileState.ContainsKey(key))
                    {
                        fileState[key] = new List<string>();
                    }

                    if (!fileState[key].Contains(fileHash))
                    {
                        fileState[key].Add(fileHash);
                        LogMessage($"记录文件状态: {fileName} (哈希: {fileHash})");
                    }

                    SaveStateFile();
                }
            });
        }

        private string GetUniqueFilePath(string filePath)
        {
            if (!File.Exists(filePath))
                return filePath;

            string directory = Path.GetDirectoryName(filePath);
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            string extension = Path.GetExtension(filePath);
            int counter = 1;

            while (File.Exists(filePath))
            {
                string newFileName = $"{fileName}_{counter}{extension}";
                filePath = Path.Combine(directory, newFileName);
                counter++;
            }

            return filePath;
        }

        private void LoadStateFile()
        {
            try
            {
                lock (stateLock)
                {
                    if (File.Exists(stateFilePath))
                    {
                        string json = File.ReadAllText(stateFilePath);
                        fileState = JsonConvert.DeserializeObject<Dictionary<string, List<string>>>(json);
                        LogMessage($"加载状态文件成功，共 {fileState?.Count ?? 0} 条记录");
                    }
                    else
                    {
                        fileState = new Dictionary<string, List<string>>();
                        LogMessage("创建新的状态文件");
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"加载状态文件失败: {ex.Message}");
                fileState = new Dictionary<string, List<string>>();
            }
        }

        private void SaveStateFile()
        {
            try
            {
                lock (stateLock)
                {
                    if (fileState != null)
                    {
                        string json = JsonConvert.SerializeObject(fileState, Formatting.Indented);
                        File.WriteAllText(stateFilePath, json);
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"保存状态文件失败: {ex.Message}");
            }
        }

        private void LogMessage(string message)
        {
            string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";
            System.Diagnostics.Debug.WriteLine(logMessage);

            // 写入日志文件
            try
            {
                string logFile = Path.Combine(logFolderPath, $"PowerPointAutoCopy_{DateTime.Now:yyyyMMdd}.log");

                using (StreamWriter writer = new StreamWriter(logFile, true))
                {
                    writer.WriteLine(logMessage);
                }
            }
            catch
            {
                // 日志写入失败不影响主要功能
            }
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}