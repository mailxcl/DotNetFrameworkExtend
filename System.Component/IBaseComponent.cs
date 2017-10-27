using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace System.Component
{
    /// <summary>
    /// 基本组件接口
    /// </summary>
    public interface IBaseComponent
    {
        /// <summary>
        /// 组件ID
        /// </summary>
        [Description("组件ID")]
        string ID { get; }

        /// <summary>
        /// 组件名称
        /// </summary>
        [Description("组件名称")]
        string Name { get; }

        /// <summary>
        /// 分组所在根，用于组织界面
        /// </summary>
        [Description("分组所在根，用于组织界面")]
        string Root { get; }

        /// <summary>
        /// 所在分组，用于组织界面
        /// </summary>
        [Description("所在分组，用于组织界面")]
        string Group { get; }
      
        /// <summary>
        /// 提示
        /// </summary>
        [Description("提示")]
        string Tip { get; }

        /// <summary>
        /// 实例类型
        /// </summary>
        [Description("实例类型")]
        Type InstanceType { get; }

        /// <summary>
        /// 实例对象
        /// </summary>
        [Description("实例对象")]
        Object Instance { get; }

        /// <summary>
        /// 小图标 16*16
        /// </summary>
        [Description("小图标16*16")]
        System.Drawing.Image SmallImage { get; }

        /// <summary>
        /// 大图标 32*32
        /// </summary>
        [Description("大图标32*32")]
        System.Drawing.Image LargeImage { get; }

        /// <summary>
        /// 是否为大图标
        /// </summary>
        [Description("是否为大图标")]
        bool IsSmallImage { get; set; }
    }
}
