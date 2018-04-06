using ExcelDataReader;
using Nortal.Utilities.Csv;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;


namespace ExcelConversion
{
    public class ConvertExcel
    {
        public LocationInfo ReadInfoFromExcel(string fileInPath)
        {
            LocationInfo locInfo = new LocationInfo();

            using (var stream = File.Open(fileInPath, FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsm)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var dataSet = reader.AsDataSet();

                    // The result of each spreadsheet is in result.Tables
                    //in xlsm sheet[0] is macro, so the first sheet is index 1
                    System.Data.DataTable workSheetCoverLocInfo = dataSet.Tables[1];

                    //Selection.Range("J8,J10,B41,C41,D41,E41,B45,D45,B50,C50,D50,E50,B53,C53,D53,B55,C55,B57,C57");
                    //"J8,J10,B41,C41,D41,E41,B45,D45,B50,C50,D50,E50,B53,C53,D53,B55,C55,B57,C57"
                    //rename columns for clarity

                    //workSheetCoverLocInfo.Columns["Column1"].ColumnName = "B";
                    //workSheetCoverLocInfo.Columns["Column2"].ColumnName = "C";
                    //workSheetCoverLocInfo.Columns["Column3"].ColumnName = "D";
                    //workSheetCoverLocInfo.Columns["Column4"].ColumnName = "E";
                    //workSheetCoverLocInfo.Columns["Column9"].ColumnName = "J";

                    //List<MapVal> input = new List<MapVal>();
                    //var spliceMap = new MapVal();
                    //spliceMap.fieldName = "SplicePoint";
                    //spliceMap.fieldLabel = "Splice Point #:";
                    //spliceMap.relativePos = 1;
                    //spliceMap.offset = 1;
                    //input.Add(spliceMap);
                    //List<string> output = new List<string>();
                    //foreach (var field in input)
                    //{
                    //   output.Add(GetFieldValue(workSheetCoverLocInfo, field.fieldName, field.fieldLabel, field.relativePos, field.offset));
                    //}
                    //var splice = GetFieldValue(workSheetCoverLocInfo, "SplicePoint", "Splice Point #:", 1, 1);
                    RowColIndexes spliceIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Splice Point #:");
                    RowColIndexes releaseIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Release #:");
                    RowColIndexes addressIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Address");
                    RowColIndexes cityStateIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "City, State:");
                    RowColIndexes floorNumberIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Floor Number");
                    RowColIndexes roomCageIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Room/Cage");
                    RowColIndexes buildingCodeIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Building CLLI Code");
                    RowColIndexes buildingNameIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Building Name \\\\ Code");
                    RowColIndexes enclosureIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Rack \\\\ Enclosure");
                    RowColIndexes makeModelIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Enclosure Make/Model:");
                    RowColIndexes ospCablesIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "OSP Cables in FDP");
                    RowColIndexes fibEngineerIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Fiber Engineer");
                    RowColIndexes feRelDateIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Release Date");
                    RowColIndexes ospPmIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "OSP PM");
                    RowColIndexes releaseNoIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Release #:");
                    RowColIndexes subfloorIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Sub Floor / Parking");
                    //RowColIndexes splicePointIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Splice Point #:");
                    RowColIndexes floorCLLIIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Floor CLLI Output");
                    RowColIndexes manholeIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Manhole #");
                    RowColIndexes manhandHoleNumberIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Manhole / Handhole#:");
                    RowColIndexes ownerIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Owner");
                    RowColIndexes mhhhOwnerIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "MH / HH Owned By:");
                    RowColIndexes mfnBackboneIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "MFN Backbone Cables(s) in Enclosure");
                    RowColIndexes facEngineerIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Facilities Engineer:");
                    RowColIndexes otherSpliceIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Other Splice Points in Node:");
                    RowColIndexes dirFacMgmtIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Dir. Facilities Management");
                    RowColIndexes dirFacMgmtDateIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Dir. Facilities Management Date");
                    RowColIndexes dirNetEngIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Dir. Network Engineering");
                    RowColIndexes dirNetEngDateIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Dir. Network Engineering Date");
                    RowColIndexes vpNetEngIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "VP Network Engineering");
                    RowColIndexes vpNetEngDateIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "VP Network Engineering Date");
                    RowColIndexes floorIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Floor:");
                    RowColIndexes rackIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Racks:");
                    RowColIndexes otherSplicesNodeIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Other Splice Points in Node");
                    RowColIndexes notesIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Notes");
                    RowColIndexes customerInfoIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Customer Information");
                    RowColIndexes hhPoleNumberIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Handhole/Pole #:");
                    RowColIndexes mhPoleOwnedByIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "MH/Pole Owned By:");
                    RowColIndexes otherSpliceMHIndexes = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, "Other Splice Points in Manhole:");

                    //create object using linq
                    locInfo = workSheetCoverLocInfo.AsEnumerable().Select(x => new LocationInfo
                    {
                        splicePoint = x.Table.Rows[spliceIndexes.rowIndex + 1][spliceIndexes.colIndex + 2].ToString().Trim()/* + x.Table.Rows[spliceIndexes.rowIndex + 1][spliceIndexes.colIndex + 1].ToString().Trim()*/,
                        release = x.Table.Rows[releaseIndexes.rowIndex + 1][releaseIndexes.colIndex + 3].ToString().Trim() + x.Table.Rows[releaseIndexes.rowIndex][releaseIndexes.colIndex + 2].ToString().Trim(),
                        address = x.Table.Rows[addressIndexes.rowIndex + 1][addressIndexes.colIndex].ToString().Trim(),
                        city = x.Table.Rows[cityStateIndexes.rowIndex + 1][cityStateIndexes.colIndex].ToString().Trim(),
                        floorNo = x.Table.Rows[floorNumberIndexes.rowIndex + 1][floorNumberIndexes.colIndex].ToString().Trim(),
                        room = x.Table.Rows[roomCageIndexes.rowIndex + 2][roomCageIndexes.colIndex + 1].ToString().Trim(),
                        buildingCode = x.Table.Rows[buildingCodeIndexes.rowIndex + 1][buildingCodeIndexes.colIndex].ToString().Trim(),
                        buildingName = x.Table.Rows[buildingNameIndexes.rowIndex + 1][buildingNameIndexes.colIndex].ToString().Trim(),
                        enclosure = x.Table.Rows[enclosureIndexes.rowIndex + 1][enclosureIndexes.colIndex].ToString().Trim(),
                        makeModel = x.Table.Rows[makeModelIndexes.rowIndex + 1][makeModelIndexes.colIndex].ToString().Trim(),
                        ospCables = x.Table.Rows[ospCablesIndexes.rowIndex + 1][ospCablesIndexes.colIndex].ToString().Trim(),
                        fiberEngineer = x.Table.Rows[fibEngineerIndexes.rowIndex + 2][fibEngineerIndexes.colIndex + 1].ToString().Trim(),
                        releaseDate = x.Table.Rows[feRelDateIndexes.rowIndex + 1][feRelDateIndexes.colIndex].ToString().Trim(),
                        ospPM = x.Table.Rows[ospPmIndexes.rowIndex + 1][ospPmIndexes.colIndex].ToString().Trim(),
                        //releaseNo = x.Table.Rows[releaseNoIndexes.rowIndex + 1][releaseNoIndexes.colIndex].ToString().Trim(),
                        subFloor = x.Table.Rows[subfloorIndexes.rowIndex + 1][subfloorIndexes.colIndex].ToString().Trim(),
                        splicePointNo = x.Table.Rows[spliceIndexes.rowIndex + 1][spliceIndexes.colIndex].ToString().Trim(),
                        floorCLLI = x.Table.Rows[floorCLLIIndexes.rowIndex + 1][floorCLLIIndexes.colIndex].ToString().Trim(),
                        //manhole = x.Table.Rows[manholeIndexes.rowIndex + 1][manholeIndexes.colIndex].ToString().Trim(),
                        //manholeHandholeNo = x.Table.Rows[manhandHoleNumberIndexes.rowIndex + 1][manhandHoleNumberIndexes.colIndex].ToString().Trim(),
                        //owner = x.Table.Rows[ownerIndexes.rowIndex + 1][ownerIndexes.colIndex].ToString().Trim(),
                        //manholeHandholeOwnedBy = x.Table.Rows[mhhhOwnerIndexes.rowIndex + 1][mhhhOwnerIndexes.colIndex].ToString().Trim(),
                        //mfnBackbone = x.Table.Rows[mfnBackboneIndexes.rowIndex + 1][mfnBackboneIndexes.colIndex].ToString().Trim(),
                        //facilitiesEngineer = x.Table.Rows[facEngineerIndexes.rowIndex + 1][facEngineerIndexes.colIndex].ToString().Trim(),
                        //otherSplicePointsInEnclosure = x.Table.Rows[otherSpliceIndexes.rowIndex + 1][otherSpliceIndexes.colIndex].ToString().Trim(),
                        //dirFacilitiesMgmt = x.Table.Rows[dirFacMgmtIndexes.rowIndex + 1][dirFacMgmtIndexes.colIndex].ToString().Trim(),
                        //dirFacilitiesMgmtDate = x.Table.Rows[dirFacMgmtDateIndexes.rowIndex + 1][dirFacMgmtDateIndexes.colIndex].ToString().Trim(),
                        //dirNetworkEngineering = x.Table.Rows[dirNetEngIndexes.rowIndex + 1][dirNetEngIndexes.colIndex].ToString().Trim(),
                        //dirNetworkEngineeringDate = x.Table.Rows[dirNetEngDateIndexes.rowIndex + 1][dirNetEngDateIndexes.colIndex].ToString().Trim(),
                        //vpNetworkEngineering = x.Table.Rows[vpNetEngIndexes.rowIndex + 1][vpNetEngDateIndexes.colIndex].ToString().Trim(),
                        //vpNetworkEngineeringDate = x.Table.Rows[vpNetEngDateIndexes.rowIndex + 1][vpNetEngDateIndexes.colIndex].ToString().Trim(),
                        //floor = x.Table.Rows[floorIndexes.rowIndex + 1][floorIndexes.colIndex].ToString().Trim(),
                        //racks = x.Table.Rows[rackIndexes.rowIndex + 1][rackIndexes.colIndex].ToString().Trim(),
                        //otherSplicePointsInNode = x.Table.Rows[otherSplicesNodeIndexes.rowIndex + 1][otherSplicesNodeIndexes.colIndex].ToString().Trim(),
                        //notes = x.Table.Rows[notesIndexes.rowIndex + 1][notesIndexes.colIndex].ToString().Trim(),
                        //customerInfo = x.Table.Rows[customerInfoIndexes.rowIndex + 1][customerInfoIndexes.colIndex].ToString().Trim(),
                        //handholeOrPoleNo = x.Table.Rows[hhPoleNumberIndexes.rowIndex + 1][hhPoleNumberIndexes.colIndex].ToString().Trim(),
                        //manholeOrPoleNo = x.Table.Rows[mhPoleOwnedByIndexes.rowIndex + 1][mhPoleOwnedByIndexes.colIndex].ToString().Trim(),
                        //otherSplicePointsInManhole = x.Table.Rows[otherSpliceMHIndexes.rowIndex + 1][otherSpliceMHIndexes.colIndex].ToString().Trim(),
                    }).First();
                }
            }

            return locInfo;
        }

        //private string GetFieldValue(System.Data.DataTable workSheetCoverLocInfo, string fieldName, string fieldLabel, int relativePos, int offset)
        //{
        //    string returnVal = "";

        //    var xy = GetTableRowColIndexesForExactMatch(workSheetCoverLocInfo, fieldLabel);

        //    switch (relativePos)
        //    {
        //        case 0:
        //            returnVal = workSheetCoverLocInfo.AsEnumerable().AsDataView().Table.Rows[xy.rowIndex + 1][xy.colIndex].ToString().Trim();
        //            break;
        //        case 1:
        //            returnVal = workSheetCoverLocInfo.AsEnumerable().AsDataView().Table.Rows[xy.rowIndex][xy.colIndex + 1].ToString().Trim();
        //            break;
        //        default:
        //            break;
        //    }
        //    return returnVal;
        //}

        private RowColIndexes GetTableRowColIndexesForExactMatch(System.Data.DataTable workSheetCoverLocInfo, string searchText)
        {
            RowColIndexes returnRCIndexes = new RowColIndexes();
            int rowIndex = -1; //return -1 if no match found

            var rowIndexArray = workSheetCoverLocInfo
             .Rows
             .Cast<DataRow>()
             .Where(r => r.ItemArray.Any(c => Regex.IsMatch(c.ToString().Trim(), searchText, RegexOptions.IgnoreCase)))
             .Select(r => r.Table.Rows.IndexOf(r)).ToArray();

            if (rowIndexArray.Length > 0)
            {
                //return the row index
                var rowCol = rowIndexArray[0];
                rowIndex = rowCol;
            }

            int colIndex = 0;
            if (rowIndex >= 0)
            {
                //loop through row until column index is found
                foreach (var dc in workSheetCoverLocInfo.Rows[rowIndex].ItemArray)
                {
                    if (dc != DBNull.Value)
                    {
                        //if (dc.ToString() == searchText)
                        if(Regex.IsMatch(dc.ToString().Trim(), searchText, RegexOptions.IgnoreCase))
                        {
                            break;
                        }
                    }
                    colIndex++;
                }
            } else
            {
                colIndex = -1; //return -1 if no match found
            }

            return new RowColIndexes { rowIndex = rowIndex, colIndex = colIndex };
        }


        private int GetTableRowIndexForContainsText(System.Data.DataTable workSheetCoverLocInfo, string searchText)
        {
            int returnRowIndex = -1; //return -1 if no match found
            var rowIndex = workSheetCoverLocInfo
             .Rows
             .Cast<DataRow>()
             //c.interiorColor color from css
             .Where(r => r.ItemArray.Any(c => Regex.IsMatch(c.ToString().Trim(), searchText, RegexOptions.IgnoreCase)))
             .Select(r => r.Table.Rows.IndexOf(r)).ToArray();

            if (rowIndex.Length > 0)
            {
                //return the row index
                returnRowIndex = rowIndex[0];
            }
            return returnRowIndex;
        }

        public void WriteToCSV(LocationInfo locInfo, string fileOutPath)
        {
            using (var writer = new StringWriter())
            {
                var csv = new CsvWriter(writer, new CsvSettings());
                //csv.Write("MyValue");                    // writing one value at a time
                //csv.Write(2, "N2");                      // or with explicit format
                //csv.WriteLine(DateTime.Now);             // or with automatic formatting
                //csv.WriteLine(1, 2, 3, 4, DateTime.Now);    // another line with many values at once
                csv.WriteLine(locInfo.splicePoint, locInfo.release, locInfo.address, locInfo.city, locInfo.floorNo, locInfo.room, locInfo.buildingCode, locInfo.buildingName, locInfo.enclosure, locInfo.makeModel, locInfo.ospCables, locInfo.fiberEngineer, locInfo.releaseDate, locInfo.ospPM, locInfo.releaseNo, locInfo.subFloor, locInfo.splicePointNo, locInfo.floorCLLI, locInfo.manhole, locInfo.manholeHandholeNo, locInfo.owner, locInfo.manholeHandholeOwnedBy, locInfo.manholeOwner, locInfo.mfnBackbone, locInfo.facilitiesEngineer, locInfo.otherSplicePointsInEnclosure, locInfo.dirFacilitiesMgmt, locInfo.dirFacilitiesMgmtDate, locInfo.dirNetworkEngineering, locInfo.dirNetworkEngineeringDate, locInfo.vpNetworkEngineering, locInfo.vpNetworkEngineeringDate, locInfo.floor, locInfo.racks, locInfo.otherSplicePointsInNode, locInfo.notes, locInfo.customerInfo, locInfo.handholeOrPoleNo, locInfo.manholeOrPoleNo, locInfo.otherSplicePointsInManhole);
                File.WriteAllText(fileOutPath, writer.ToString());
            }
        }
    }


}
    