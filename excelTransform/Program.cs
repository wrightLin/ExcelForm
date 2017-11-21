using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.IO;

namespace excelTransform
{
    class Program
    {

        static void Main(string[] args)
        {
            NPOIHelper helper = new NPOIHelper();
            FileFinder finder = new FileFinder();
            
            //找出所有包含目標資料表名稱的檔案路徑
            List<string> TarFileAdds = finder.FindTargetFileAddress();
            //remove duplicate
            var DistTarFileAdds = TarFileAdds.GroupBy(x => x).Select(y => y.First());


            //將路徑餵給helper，修改excle檔案並儲存。
            foreach (string address in DistTarFileAdds)
            {
                helper.LoadExcel(address);
            }

            Console.WriteLine("執行完畢，按任意鍵退出。");
            Console.ReadKey();

        }
    }
}
