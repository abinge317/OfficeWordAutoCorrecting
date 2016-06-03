using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeWordAutoCorrecting
{
    class TestPoint
    {
        public const int BORDER_SIZE = 101; //边框粗细
        public const int BORDER_COLOR = 102;  //边框颜色
        public const int BORDER_SHADOW = 103; //边框阴影
        public const int BORDER_TYPE = 104;  //边框类型
        public const int SHADING_COLOR = 105;  //底纹填充
        public const int SHADING_TYPE = 106;  //底纹（图案）样式
        public const int SHADIN_FILL = 107;  //底纹（图案）颜色

        public static Dictionary<int, int> testPoint2Score = new Dictionary<int, int>();
    }
}
