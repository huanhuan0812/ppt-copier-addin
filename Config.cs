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
        public int LogRetentionDays { get; set; } = 7;
    }
}