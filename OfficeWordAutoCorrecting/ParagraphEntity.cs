using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeWordAutoCorrecting
{
    class ParagraphEntity
    {
        private bool withBorder = false;
        private String topBorderColor = "";
        private String bottomBorderColor = "";
        private String leftBorderColor = "";
        private String rightBorderColor = "";
        private bool withShadow = false;
        private uint borderSize = 0;
        private string borderType = null;
        private string paragraphText = "";

        private bool withShading = false;

        public bool WithBorder
        {
            get
            {
                return withBorder;
            }
            set
            {
                withBorder = value;
            }
        }

        public string TopBorderColor
        {
            get
            {
                return topBorderColor;
            }
            set
            {
                topBorderColor = value;
            }
        }

        public string BottomBorderColor
        {
            get
            {
                return bottomBorderColor;
            }
            set
            {
                bottomBorderColor = value;
            }
        }

        public string LeftBorderColor
        {
            get
            {
                return leftBorderColor;
            }
            set
            {
                leftBorderColor = value;
            }
        }

        public string RightBorderColor
        {
            get
            {
                return rightBorderColor;
            }
            set
            {
                rightBorderColor = value;
            }
        }

        public bool WithShadow
        {
            get
            {
                return withShadow;
            }
            set
            {
                withShadow = value;
            }
        }

        public uint BorderSize
        {
            get
            {
                return borderSize;
            }
            set
            {
                borderSize = value;
            }
        }

        public string BorderType
        {
            get
            {
                return borderType;
            }
            set
            {
                borderType = value;
            }
        }

        public bool WithShading
        {
            get
            {
                return withShading;
            }
            set
            {
                withShading = value;
            }
        }

        public string ParagraphText
        {
            get
            {
                return paragraphText;
            }
            set
            {
                paragraphText = value;
            }
        }
    }
}
