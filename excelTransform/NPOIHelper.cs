using System;
using System.Data;
using System.IO;
using System.Web;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.POIFS;
using NPOI.Util;
using System.Text;
using System.Configuration;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Linq;

public class NPOIHelper
{
    Dictionary<string, string> MyDic = new Dictionary<string, string>();
    List<string> DicKeys = ConfigurationManager.AppSettings["DicKeys"].Split(',').ToList();
    List<string> DicVals = ConfigurationManager.AppSettings["DicVals"].Split(',').ToList();
    
    //建構式
    public NPOIHelper()
    {
        CreateDictionary();
    }



    // 建立字典
    private void CreateDictionary()
    {
        if (DicKeys.Count == DicVals.Count)
        {
            for (int i = 0; i < DicKeys.Count; i++)
            {
                MyDic.Add(DicKeys[i], DicVals[i]);
            }
        }
        else
        {
            Console.WriteLine("DicKey數量與DicVal數量不一致");
            Console.ReadKey();
            Environment.Exit(0);
        }

    }

  

   /// <summary>
   /// 檢查CellVal是否包含特殊字元，將其更新後返回。
   /// </summary>
   /// <param name="CellVal"></param>
   /// <returns></returns>
    private String FindInDictionaryAndReplace(String CellVal)
    {
        string result = CellVal;

        foreach (var dic in MyDic)
        {
            if (CellVal.Contains(dic.Key))
            {
                result = CellVal.Replace(dic.Key, MyDic[dic.Key]);
                break;
            }

        }

        return result;
    }

    /// <summary>
    /// 取代字串並存檔(只針對xls,xlsx)
    /// </summary>
    public void LoadExcel(string FileAddress)
    {
        try
        {
            //precheck 
            if (FileAddress.Contains(".xlsx"))
            {
                xlsxFileProccess(FileAddress);
            }
            else if (FileAddress.Contains(".xls"))
            {
                xlsFileProccess(FileAddress);
            }
            else
            {
                return;
            }
        }
        catch (Exception e)
        {
            Console.WriteLine("錯誤訊息：" + e.Message);
            Console.WriteLine("按任意鍵繼續...");
            Console.ReadKey();
        }
       

        
    }

    /// <summary>
    /// xlsx file Proccess
    /// </summary>
    /// <param name="fileAddress"></param>
    private void xlsxFileProccess(string fileAddress)
    {
        XSSFWorkbook workbook;

        //讀取excel檔案
        using (FileStream file = new FileStream(fileAddress, FileMode.Open, FileAccess.Read))
        {
            Console.WriteLine("讀取excel檔案：" + fileAddress);
            workbook = new XSSFWorkbook(file);
            file.Close();
        }
        //讀取工作表
        var sheet = workbook.GetSheetAt(0);

        //字串處理
        for (int row = 0; row <= sheet.LastRowNum; row++)
        {
            if (sheet.GetRow(row) != null)
            {
                foreach (var c in sheet.GetRow(row).Cells)
                {
                    if (c.CellType == CellType.String)
                    {
                        string result = FindInDictionaryAndReplace(c.StringCellValue);
                        c.SetCellValue(result);
                    }
                }
            }
        }

        //儲存檔案
        //using (var file = new FileStream(fileAddress, FileMode.Open, FileAccess.Write))
        using (var file = new FileStream(fileAddress, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
        {
            workbook.Write(file);
            file.Close();
        }


    }

    /// <summary>
    /// xls file proccess
    /// </summary>
    /// <param name="fileAddress"></param>
    private void xlsFileProccess(string fileAddress)
    {
        HSSFWorkbook workbook;

        //讀取excel檔案
        using (FileStream file = new FileStream(fileAddress, FileMode.Open, FileAccess.Read))
        {
            Console.WriteLine("讀取excel檔案：" + fileAddress);
            workbook = new HSSFWorkbook(file);
            file.Close();
        }

        //讀取工作表
        var sheet = workbook.GetSheetAt(0);

        //字串處理
        for (int row = 0; row <= sheet.LastRowNum; row++)
        {
            if (sheet.GetRow(row) != null)
            {
                foreach (var c in sheet.GetRow(row).Cells)
                {
                    if (c.CellType == CellType.String)
                    {
                        string result = FindInDictionaryAndReplace(c.StringCellValue);
                        c.SetCellValue(result);
                    }
                }
            }
        }

        //儲存檔案
        using (FileStream file = new FileStream(fileAddress, FileMode.Open, FileAccess.Write))
        {
            workbook.Write(file);
            file.Close();
        }


    }
}