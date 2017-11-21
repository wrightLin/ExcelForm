using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.IO;

namespace excelTransform
{
    class FileFinder
    {
        //欲修改之Table名稱
        List<string> TargetTableNames = ConfigurationManager.AppSettings["TargetTableNames"].Split(',').ToList();
        //欲遍巡資料夾路徑
        string FolderAddresses = ConfigurationManager.AppSettings["ExcelFileAddress"];
        //目標excel檔案路徑清單
        List<string> TargetExcelFileAddress = new List<string>();

        /// <summary>
        /// 找出目標資料夾中，包含欲修改資料表名稱的檔案路徑。
        /// </summary>
        /// <returns></returns>
        public List<string> FindTargetFileAddress()
        {
            //取出目標文件夾
            DirectoryInfo TheFolder = new DirectoryInfo(FolderAddresses);
            //遍歷文件夹
            CheckIfTargetFileExist(TheFolder);

            return TargetExcelFileAddress;
        }


        private void CheckIfTargetFileExist(DirectoryInfo Folder)
        {
            //遍歷文件
            foreach (FileInfo NextFile in Folder.GetFiles())
            {
                //檢查檔名是否包含欲修改tableName
                foreach (string targetTableName in TargetTableNames)
                {
                    if (NextFile.FullName.Contains(targetTableName)|| targetTableName =="ALL")
                    {
                        TargetExcelFileAddress.Add(NextFile.FullName);
                    }
                }
            }

            //如果還有子資料夾，回call自己一次
            if (Folder.GetDirectories().Count() > 0)
            {
                foreach (DirectoryInfo folder in Folder.GetDirectories())
                {
                    CheckIfTargetFileExist(folder);
                }
            }
            return;
        }
        
    }
}
