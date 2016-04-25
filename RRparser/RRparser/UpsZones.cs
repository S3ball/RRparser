using System.Collections.Generic;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace RRparser
{
    public class UpsZones
    {

        private int HawaiiRow1Index { get; set; }
        private int HawaiiRow2Index { get; set; }
        private int AlaskaRow1Index { get; set; }
        private int AlaskaRow2Index { get; set; }
        private int LastRowIndex { get; set; }

        private int[] _hawaiiZones1;
        private int[] _hawaiiZones2;
        private int[] _alaskaZones1;
        private int[] _alaskaZones2;


        public HSSFWorkbook ZonesWb;
        public string SheetName { get; set; }
        public ISheet ZonesSheet { get; set; }
        public string[] ServiceZoneList;
        //0 Ground
        //1 3rd Select
        //2 2nd Day Air
        //3 2nd Day Air AM
        //4 Next Day Air Saver
        //5 Next Day Air

        public UpsZones(string path)
        {
            using (var file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                ZonesWb = new HSSFWorkbook(file);
                SheetName = ZonesWb.GetSheetName(0);
            }
            InitializeObject();
        }

        public UpsZones(string path, string fromZip)
        {
            using (var file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                ZonesWb = new HSSFWorkbook(file);
            }
            SheetName = fromZip;
            InitializeObject();
        }

        private void InitializeObject()
        {
            ServiceZoneList = new string[6];
            ZonesSheet = ZonesWb.GetSheetAt(0);
            LastRowIndex = ZonesSheet.LastRowNum;
            GetRowPositions();
            for (var serviceIndex = 1; serviceIndex <= 6; serviceIndex++)
            {
                ServiceZoneList[serviceIndex - 1] = "from_zip,to_zip,zone\n";
                SetBasicZones(serviceIndex);
            }
            AlaskaHawaiiZones();
            SetAdditionalZones(HawaiiRow1Index, HawaiiRow2Index, _hawaiiZones1);
            SetAdditionalZones(HawaiiRow2Index, AlaskaRow1Index, _hawaiiZones2);
            SetAdditionalZones(AlaskaRow1Index, AlaskaRow2Index, _alaskaZones1);
            SetAdditionalZones(AlaskaRow2Index, LastRowIndex, _alaskaZones2);

        }

        private void GetRowPositions()
        {
            for (var row = 0; row <= LastRowIndex; row++)
            {
                var firstCol = ZonesSheet.GetRow(row).GetCell(0);
                if (firstCol == null) continue;
                if (firstCol.ToString().StartsWith("[2] For Hawaii"))
                    HawaiiRow1Index = row;
                if (firstCol.ToString().StartsWith("For Hawaii"))
                    HawaiiRow2Index = row;
                if (firstCol.ToString().StartsWith("[3] For Alaska"))
                    AlaskaRow1Index = row;
                if (firstCol.ToString().StartsWith("For Alaska"))
                    AlaskaRow2Index = row;
            }
        }

        private void SetBasicZones(int serviceIndex)
        {
            for (var row = 0; row <= HawaiiRow1Index; row++)
            {
                if (ZonesSheet.GetRow(row) == null) continue;
                var firstCol = ZonesSheet.GetRow(row).GetCell(0);
                var secondCol = ZonesSheet.GetRow(row).GetCell(serviceIndex);
                int num;
                if (firstCol == null || secondCol == null || !int.TryParse(secondCol.ToString(), out num)) continue;
                ServiceZoneList[serviceIndex - 1] += string.Format("{0},{1},{2}\n", SheetName, firstCol, num);

            }
        }

        private void SetAdditionalZones(int startRow, int endRow, IList<int> zoneId)
        {
            for (var row = startRow + 1; row < endRow; row++)
            {
                for (var i = 0; i < 9; i++)
                {
                    //set for Ground, Next Day Air, 2nd Day Air
                    ServiceZoneList[0] += AppendZones(ZonesSheet.GetRow(row).GetCell(i).ToString(), zoneId[0]);
                    ServiceZoneList[5] += AppendZones(ZonesSheet.GetRow(row).GetCell(i).ToString(), zoneId[1]);
                    ServiceZoneList[2] += AppendZones(ZonesSheet.GetRow(row).GetCell(i).ToString(), zoneId[2]);
                }
            }
        }

        private void AlaskaHawaiiZones()
        {
            _hawaiiZones1 = new int[3];
            _hawaiiZones2 = new int[3];
            _alaskaZones1 = new int[3];
            _alaskaZones2 = new int[3];
            var alaskaZones1 = ZonesSheet.GetRow(AlaskaRow1Index).GetCell(0).ToString();
            var alaskaZones2 = ZonesSheet.GetRow(AlaskaRow2Index).GetCell(0).ToString();
            var hawaiiZones1 = ZonesSheet.GetRow(HawaiiRow1Index).GetCell(0).ToString();
            var hawaiiZones2 = ZonesSheet.GetRow(HawaiiRow2Index).GetCell(0).ToString();
            _alaskaZones1 = SetZonesArray(alaskaZones1);
            _alaskaZones2 = SetZonesArray(alaskaZones2);
            _hawaiiZones1 = SetZonesArray(hawaiiZones1);
            _hawaiiZones2 = SetZonesArray(hawaiiZones2);
        }

        private string AppendZones(string cell, int zoneId)
        {
            return !string.IsNullOrEmpty(cell) ? string.Format("{0},{1},{2}\n", SheetName, cell, zoneId) : string.Empty;
        }

        private static int[] SetZonesArray(string row)
        {
            var zones = new List<int>();
            for (var index = 0; index < row.Split(' ').Length; index++)
            {
                var word = row.Split(' ')[index];
                if (word != "Zone") continue;
                var zoneId = row.Split(' ')[index + 1];
                zones.Add(int.Parse(zoneId));
            }
            return zones.ToArray();
        }

        public void ExportToCsv(string path)
        {
            File.WriteAllText(path + "\\" + SheetName + "UPS Ground Zones.csv", ServiceZoneList[0]);
            File.WriteAllText(path + "\\" + SheetName + "UPS 3 Day Select Zones.csv", ServiceZoneList[1]);
            File.WriteAllText(path + "\\" + SheetName + "UPS 2nd Day Air Zones.csv", ServiceZoneList[2]);
            File.WriteAllText(path + "\\" + SheetName + "UPS 2nd Day Air A.M. Zones.csv", ServiceZoneList[3]);
            File.WriteAllText(path + "\\" + SheetName + "UPS Next Day Air Saver Zones.csv", ServiceZoneList[4]);
            File.WriteAllText(path + "\\" + SheetName + "UPS Next Day Air Zones.csv", ServiceZoneList[5]);

        }

    }
}
