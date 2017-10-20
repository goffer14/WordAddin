using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace eDocs_Editor
{
    public class multiLOEP
    {
        public string startPage;
        public string endPage;
        public string rev;
        public string date;

        public multiLOEP(string startPage, string rev, string date)
        {
            this.startPage = startPage;
            this.endPage = startPage;
            this.rev = rev;
            this.date = date;
        }
    }
}
