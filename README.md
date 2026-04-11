# PowerPoint 自动复制插件 (PowerPoint Auto Copy Add-in)

## 简介

这是一个 VSTO (Visual Studio Tools for Office) 插件，用于 PowerPoint。当从移动存储设备（如 U 盘、移动硬盘等）打开 PowerPoint 文档时，插件会自动将文档复制到指定的本地目录，并按日期分文件夹保存，同时通过 MD5 哈希值防止重复复制。

## 功能特性

- ✅ **自动检测移动存储设备**：智能识别文档是否来自 U 盘、移动硬盘等可移动设备
- ✅ **自动复制备份**：将文档自动复制到本地指定目录
- ✅ **按日期分文件夹**：自动创建 `YYYY-MM-DD` 格式的日期文件夹
- ✅ **防重复机制**：使用 MD5 哈希值记录已复制的文件，避免重复复制
- ✅ **文件名冲突处理**：当文件名冲突时自动添加序号（如 `文档_1.pptx`）
- ✅ **异步操作**：所有文件操作都在后台异步执行，不阻塞 PowerPoint 主界面
- ✅ **保护视图支持**：支持从保护视图中打开的文档
- ✅ **日志记录**：详细记录所有操作，便于排查问题
- ✅ **自动日志清理**：启动时自动清理超过保留天数的日志文件（默认7天）
- ✅ **可配置**：支持通过配置文件自定义目标路径、日志保留天数等

## 系统要求

- **操作系统**：Windows 7 或更高版本
- **PowerPoint**：Microsoft PowerPoint 2010 或更高版本
- **.NET Framework**：4.7.2 或更高版本
- **Visual Studio Tools for Office Runtime**：需要安装 VSTO 运行时

## 安装方法

### 方法一：直接安装（推荐）

1. 下载编译好的压缩包并解压到固定位置
2. 双击运行setup.exe
3. 按照提示完成安装
4. 启动 PowerPoint，插件会自动运行

### 方法二：从源码编译

1. **环境准备**
   - 安装 Visual Studio 2019 或更高版本
   - 安装 Office/SharePoint 开发工作负载

2. **打开项目**
   ```bash
   git clone https://github.com/huanhuan0812/ppt-copier-addin.git
   cd ppt-copier-addin
   ```

3. **还原 NuGet 包**
   ```bash
   nuget restore
   ```

4. **编译项目**
   - 在 Visual Studio 中打开解决方案
   - 选择 `Release` 配置
   - 生成解决方案 (Ctrl+Shift+B)

5. **发布**
   - 右键项目 → 发布
   - 按照向导完成发布

## 配置说明

### 配置文件位置

插件首次运行时，会自动创建配置文件：

```
%LocalAppData%\PowerPointAutoCopy\config.json
```

例如：
- Windows 10/11: `C:\Users\<用户名>\AppData\Local\PowerPointAutoCopy\config.json`

### 配置参数

```json
{
  "TargetCopyPath": "D:\\Backup",
  "EnableLogging": true,
  "EnableAutoCopy": true,
  "DateFolderFormat": "yyyy-MM-dd",
  "LogRetentionDays": 7
}
```

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `TargetCopyPath` | string | `D:\Backup` | 目标复制路径（文件将保存到此目录） |
| `EnableLogging` | bool | `true` | 是否启用日志记录 |
| `EnableAutoCopy` | bool | `true` | 是否启用自动复制功能 |
| `DateFolderFormat` | string | `yyyy-MM-dd` | 日期文件夹格式 |
| `LogRetentionDays` | int | `7` | 日志文件保留天数（启动时自动清理） |

### 修改配置

1. 关闭 PowerPoint
2. 打开上述路径的 `config.json` 文件
3. 修改相应参数
4. 保存文件
5. 重新启动 PowerPoint

## 使用说明

### 基本使用

1. **插入 U 盘**：将移动存储设备连接到电脑
2. **打开文档**：直接从 U 盘打开 PowerPoint 文档
3. **自动复制**：插件会在后台自动将文档复制到配置的目标目录
4. **日志查看**：可在日志文件中查看复制记录

### 文件保存结构

```
D:\Backup\
├── 2024-12-25\
│   ├── 演示文稿.pptx
│   └── 报告.pptx
├── 2024-12-26\
│   ├── 会议记录.pptx
│   └── 培训材料.pptx
└── 2024-12-27\
    └── 项目计划.pptx
```

### 日志文件

日志文件位置：
```
%LocalAppData%\PowerPointAutoCopy\Logs\PowerPointAutoCopy_YYYYMMDD.log
```

日志示例：
```
[2024-12-25 14:30:15] PowerPoint自动复制插件已启动
[2024-12-25 14:30:15] 日志保留天数: 7天，已清理过期日志
[2024-12-25 14:30:15] 加载状态文件成功，共 0 条记录
[2024-12-25 14:32:20] 检测到文档打开: E:\文档\演示文稿.pptx (打开方式: Normal)
[2024-12-25 14:32:20] 驱动器类型检测: E:\ - Removable, 是否可移动: True
[2024-12-25 14:32:20] 检测到移动存储设备文件，开始复制: E:\文档\演示文稿.pptx
[2024-12-25 14:32:21] 文件复制成功: E:\文档\演示文稿.pptx -> D:\Backup\2024-12-25\演示文稿.pptx
```

### 状态文件

防重复机制使用 `state.json` 记录已复制的文件信息：

```
%LocalAppData%\PowerPointAutoCopy\state.json
```

该文件记录了每个文件的 MD5 哈希值，用于判断是否为重复文件。

## 卸载方法

### 方法一：通过控制面板

1. 打开控制面板
2. 进入"程序和功能"
3. 找到 `PowerPoint Auto Copy Add-in`
4. 右键点击"卸载"

### 方法二：通过 Visual Studio

1. 打开项目
2. 右键项目 → 卸载

## 常见问题

### Q1: 插件没有自动复制文件？

**A**: 请检查以下几点：
1. 确认配置文件中的 `EnableAutoCopy` 为 `true`
2. 确认目标路径 `TargetCopyPath` 存在且有写入权限
3. 检查文档是否真的在移动存储设备上（U盘、移动硬盘等）
4. 查看日志文件了解具体错误信息

### Q2: 如何更改备份位置？

**A**: 修改配置文件中的 `TargetCopyPath` 参数，例如：
```json
"TargetCopyPath": "E:\\MyBackups\\PowerPoint"
```

### Q3: 日志文件太多，如何减少？

**A**: 修改配置文件中的 `LogRetentionDays` 参数：
```json
"LogRetentionDays": 3   // 只保留3天日志
```

### Q4: 插件影响 PowerPoint 启动速度吗？

**A**: 不会。所有文件操作都在后台异步执行，不会阻塞 PowerPoint 主界面。

### Q5: 可以复制网络驱动器上的文件吗？

**A**: 当前版本只检测移动存储设备（`DriveType.Removable`），网络驱动器不会被自动复制。如需支持，可修改 `IsRemovableDriveAsync` 方法。

### Q6: 如何查看插件是否运行？

**A**: 
1. 打开 PowerPoint
2. 查看日志文件是否有启动记录
3. 或使用调试工具查看输出日志

### Q7: 插件会复制保护视图中的文档吗？

**A**: 会。插件支持 `ProtectedViewWindowOpen` 事件，会自动处理保护视图中的文档。

## 开发与调试

### 调试步骤

1. 在 Visual Studio 中打开项目
2. 设置断点
3. 按 F5 启动调试
4. PowerPoint 会自动启动
5. 执行测试操作
6. 查看输出窗口的调试信息

### 日志级别

插件使用 `System.Diagnostics.Debug.WriteLine()` 输出调试信息，可在 Visual Studio 输出窗口中查看。

## 技术架构

### 核心技术

- **VSTO 4.0**：Office 插件开发框架
- **.NET Framework 4.7.2**：运行时环境
- **PowerPoint Interop**：与 PowerPoint 交互
- **Newtonsoft.Json**：配置文件解析
- **MD5**：文件哈希计算

### 事件处理

- `AfterPresentationOpen`：处理正常打开的文档
- `ProtectedViewWindowOpen`：处理保护视图中的文档

### 异步处理

所有文件 I/O 操作都使用 `async/await` 异步模式，避免阻塞 UI 线程。

## 许可证

Copyright © 2024

## 更新日志

### Version 1.0.0 (2024-12-25)

- ✨ 初始版本发布
- ✅ 支持移动存储设备检测
- ✅ 支持自动复制备份
- ✅ 支持按日期分文件夹
- ✅ 支持 MD5 防重复机制
- ✅ 支持保护视图
- ✅ 支持日志记录与自动清理
- ✅ 支持配置文件自定义

## 联系方式

如有问题或建议，请通过以下方式联系：

- 提交 Issue
- 发送邮件

## 致谢

感谢所有使用和支持本项目的用户！

---

**注意**：使用本插件时，请确保遵守相关法律法规和公司政策。作者不对因使用本插件造成的数据丢失或泄露负责。建议定期检查备份文件的完整性。