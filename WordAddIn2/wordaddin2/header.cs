using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordAddIn2
{
    public class header
    {
        public int pageNum;
        public string H1_num;
        public string H2_num;

        public header(int pageNum, string H1_num, string H2_num)
        {
            this.pageNum = pageNum;
            this.H1_num = H1_num;
            this.H2_num = H2_num;
        }
        public void setHeader(int pageNum, string H1_num, string H2_num)
        {
            this.pageNum = pageNum;
            this.H1_num = H1_num;
            this.H2_num = H2_num;
            /**
            if (numValue != "" && numValue.IndexOf(".", (numValue.Length - 1)) > 0)
                numValue = numValue.Remove(numValue.Length-1, 1);
            this.numValue = numValue;
    */
        }
    }
}
