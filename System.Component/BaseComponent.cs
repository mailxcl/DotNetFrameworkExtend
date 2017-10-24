using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;

namespace System.Component
{
    public abstract class BaseComponent : IBaseComponent
    {
        #region 抽象属性
        /// <summary>
        /// 组件ID
        /// </summary>
        [Description("组件ID")]
        public abstract string ID { get; }

        /// <summary>
        /// 组件名称
        /// </summary>
        [Description("组件名称")]
        public abstract string Name { get; }

        /// <summary>
        /// 分组所在根，用于组织界面
        /// </summary>
        [Description("分组所在根，用于组织界面")]
        public virtual string Root { get { return "默认页"; } }

        /// <summary>
        /// 所在分组，用于组织界面
        /// </summary>
        [Description("所在分组，用于组织界面")]
        public virtual string Group { get { return "默认组"; } }

        /// <summary>
        /// 提示
        /// </summary>
        [Description("提示")]
        public virtual string Tip { get { return this.Name; } }

        /// <summary>
        /// 小图标 16*16
        /// </summary>
        [Description("小图标16*16")]
        public abstract Image SmallImage { get; }

        /// <summary>
        /// 大图标 32*32
        /// </summary>
        [Description("大图标32*32")]
        public abstract Image LargeImage { get; }

        /// <summary>
        /// 是否为大图标
        /// </summary>
        [Description("是否为大图标")]
        public virtual bool IsSmallImage { get; set; }
        #endregion

        #region 属性
        /// <summary>
        /// 实例类型
        /// </summary>
        [Description("实例类型")]
        public Type InstanceType { get { return this.GetType(); } }

        /// <summary>
        /// 实例对象
        /// </summary>
        [Description("实例对象")]
        public Object Instance { get { return this; } }
        #endregion
    }
}
