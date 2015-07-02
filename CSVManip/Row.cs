using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSVManip
{
    public class Row
    {
        public String pic1 = "", pic2 = "", pic3 = "", pic4 = "", pic5 = "", pic6 = "", pic7 = "", pic8 = "", pic9 = "", pic10 = "", pic11 = "", pic12 = "";
        public String title = "";
        public double price = 0;
        public String text = "";
        public String ref_no = "";
        public String ref_url = "";

        public bool isEmpty
        {
            get
            {
                return pic1.Trim().Length == 0 && pic2.Trim().Length == 0 && pic3.Trim().Length == 0 && pic4.Trim().Length == 0 && pic5.Trim().Length == 0 && pic6.Trim().Length == 0 && pic7.Trim().Length == 0 && pic8.Trim().Length == 0 && pic9.Trim().Length == 0 && pic10.Trim().Length == 0 && pic11.Trim().Length == 0 && pic12.Trim().Length == 0 &&
                    title.Trim().Length == 0 && price == 0 && text.Trim().Length == 0 && ref_no.Trim().Length == 0 && ref_url.Trim().Length == 0;
            }
        }
    }
}
