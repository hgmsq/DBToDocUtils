using DBToDocUtils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            NpoiToDoc npoiToDoc = new NpoiToDoc(1);
            npoiToDoc.CreateToWord("data source=127.0.0.1;user id=root;password=12345678;database=test;port=3306;pooling=false;charset=utf8;","test","D:");

            npoiToDoc.CreateToHtml("data source=127.0.0.1;user id=root;password=12345678;database=test;port=3306;pooling=false;charset=utf8;", "test", "D:");

            npoiToDoc.CreateToMarkDown("data source=127.0.0.1;user id=root;password=12345678;database=test;port=3306;pooling=false;charset=utf8;", "test", "D:");
        }
    }
}
