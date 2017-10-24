using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace System.Component
{
    public interface IBaseCmd
    {
        [Description("执行命令")]
        void Execute();
    }
}
