using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordAddIn2
{
    public class HeadingPos
    {
        public int H1_column;
        public int H1_row;
        public int H1_pos;


        public int H2_column;
        public int H2_row;
        public int H2_pos;

        public int rev_column;
        public int rev_row;
        public int rev_pos;

        public int date_column;
        public int date_row;
        public int date_pos;

        public int page_column;
        public int page_row;
        public int page_pos;

        public int text1_column;
        public int text1_row;
        public int text1_pos;

        public int text2_column;
        public int text2_row;
        public int text2_pos;

        public int text3_column;
        public int text3_row;
        public int text3_pos;

        public int text4_column;
        public int text4_row;
        public int text4_pos;
        public HeadingPos()
        {
            text1_pos = -1;
            text2_pos = -1;
            text3_pos = -1;
            text4_pos = -1;
        }
    }
}
