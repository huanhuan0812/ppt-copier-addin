using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ppt_copier_addin
{
    public class Config
    {
        /// <summary>
        /// 目标复制路径
        /// </summary>
        public string TargetCopyPath { get; set; }

        /// <summary>
        /// 是否启用日志记录
        /// </summary>
        public bool EnableLogging { get; set; } = true;

        /// <summary>
        /// 是否启用自动复制
        /// </summary>
        public bool EnableAutoCopy { get; set; } = true;

        /// <summary>
        /// 日期文件夹格式
        /// </summary>
        public string DateFolderFormat { get; set; } = "yyyy-MM-dd";

        /// <summary>
        /// 日志保留天数（默认7天）
        /// </summary>
        public int LogRetentionDays { get; set; } = 3;

        /// <summary>
        /// 当目标路径不可写入时是否使用回退路径
        /// </summary>
        public bool UseFallbackOnError { get; set; } = true;

        /// <summary>
        /// 是否启用版本控制（保存时创建版本历史）
        /// </summary>
        public bool EnableVersionControl { get; set; } = true;

        /// <summary>
        /// 每个文件保留的最大版本数（仅当 EnableVersionControl 为 true 时生效）
        /// </summary>
        public int MaxVersionsPerFile { get; set; } = 5;

        /// <summary>
        /// 是否在保存时自动覆盖更新主文件
        /// </summary>
        public bool UpdateMainFileOnSave { get; set; } = true;
    }
}