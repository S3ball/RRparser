using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace RateRouteParser
{
    public class UpsRates
    {
        public HSSFWorkbook RatesWorkbook;
        private ISheet CurrentSheet { get; set; }
        private int LastGroundRow { get; set; }
        public string[] ServicesArray;
        private int CurrentWidth { get; set; }

        public UpsRates(string path)
        {
            using (var file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                RatesWorkbook = new HSSFWorkbook(file);
            }
        
         

            ServicesArray = new string[7];
           
            for (var i = 0; i < 7; i++)
            {
                ServicesArray[i] = SheetToString(i);
            }

        }


        private string SheetToString(int sheetIndex)
        {


            var sheetContent = "";
            CurrentSheet = RatesWorkbook.GetSheetAt(sheetIndex);
            LastGroundRow = CurrentSheet.LastRowNum;
            SetWidth();
           
            var tab = new string[CurrentWidth];
            for (var row = 0; row <= LastGroundRow; row++)
            {
                try
                {
                    var conditionalCell1 = CurrentSheet.GetRow(row).GetCell(2);
                    var conditionalCell2 = CurrentSheet.GetRow(row).GetCell(1);
                    if (conditionalCell2.ToString() == "Price Per Pound") continue;
                    if (conditionalCell2.ToString() == "Zones" && !string.IsNullOrEmpty(tab[0])) continue;
                    if (string.IsNullOrEmpty(conditionalCell1.ToString())) continue;
                    if (string.IsNullOrEmpty(conditionalCell2.ToString())) continue; //new condition added

                    for (var i = 0; i < tab.Length; i++)
                    {

                        tab[i] = CurrentSheet.GetRow(row).GetCell(i + 1)
                            .ToString()
                            .Replace("$", "")
                            .Replace(" Lbs.", "")
                            .Replace(",", ".")
                            .Replace(" ", "");
                    }
                    for (var index = 0; index < tab.Length; index++)
                    {
                        var item = tab[index];
                        if (index > 0 && !string.IsNullOrEmpty(item))
                        {
                            Console.Write(",");
                            sheetContent += ',';
                        }
                        sheetContent += item;
                        Console.Write(item);
                    }
                    sheetContent += Environment.NewLine;
                    Console.WriteLine();
                 
                    
                }
                catch (Exception)
                {
                    // ignored
                }

            }
          
            return sheetContent;
        }

        private void SetWidth()
        {
            CurrentWidth = 0;
            var list = new List<ICell>();
            for (var row = 0; row <= LastGroundRow; row++)
            {
                try
                {
                    var conditionalCell1 = CurrentSheet.GetRow(row).GetCell(1);

                    if (conditionalCell1.ToString() != "Zones" && conditionalCell1.ToString() != "Lbs.") continue;
                    list = CurrentSheet.GetRow(row).Cells;

                    //fix header
                    if (conditionalCell1.ToString() == "Lbs.") 
                    {
                        conditionalCell1.SetCellValue("Zones");                     
                    }
                  
                    break;
                }
                catch (Exception)
                {
                    // ignored
                }

            }
            foreach (var item in list)
            {
                try
                {
                    if (item.ToString() != "") CurrentWidth++;
                    var zoneID = item.ToString().Replace("Zone ", "");
                    item.SetCellValue(zoneID);
                }
                catch (Exception)
                {
                    // ignored
                }
            }
        }

        public void ExportToCsv(string path)
        {
            try
            {
                File.WriteAllText(path + "\\" + "rates_" + "ground.csv", ServicesArray[6]);
                File.WriteAllText(path + "\\" + "rates_" + "third day select.csv", ServicesArray[5]);
                File.WriteAllText(path + "\\" + "rates_" + "second day air.csv", ServicesArray[4]);
                File.WriteAllText(path + "\\" + "rates_" + "second day air am.csv", ServicesArray[3]);
                File.WriteAllText(path + "\\" + "rates_" + "next day air saver.csv", ServicesArray[2]);
                File.WriteAllText(path + "\\" + "rates_" + "next day air.csv", ServicesArray[1]);
                File.WriteAllText(path + "\\" + "rates_" + "next day early am.csv", ServicesArray[0]);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
