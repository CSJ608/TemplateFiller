using System;
using System.Collections.Generic;
using System.Text;

namespace TemplateFiller.Abstractions
{
    public enum TemplateType
    {
        /// <summary>
        /// 支持.xlsx和.xls
        /// </summary>
        Excel = 1,
        /// <summary>
        /// 支持.docx
        /// </summary>
        Word = 2,
    }
}
