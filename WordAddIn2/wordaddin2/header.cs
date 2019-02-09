using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace eDocs_Editor
{
    public class header
    {
        public int pageNum;
        public string headingNum;

        public header(int pageNum, string headingNum)
        {
            this.pageNum = pageNum;
            this.headingNum = headingNum;
        }
        public String getHeadingString()
        {
            return headingNum;
        }
    }
}
