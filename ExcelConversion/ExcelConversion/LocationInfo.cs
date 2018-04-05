namespace ExcelConversion
{
    public class LocationInfo
    {
        public string splicePoint { get; set; }
        public string release { get; set; }
        public string address { get; set; }
        public string city { get; set; }
        public string floorNo { get; set; }
        public string room { get; set; }
        public string buildingCode { get; set; }
        public string buildingName { get; set; }
        public string enclosure { get; set; }
        public string makeModel { get; set; }
        public string ospCables { get; set; }
        public string fiberEngineer { get; set; }
        public string releaseDate { get; set; }
        public string ospPM { get; set; }
        public string releaseNo { get; set; }
        public string subFloor { get; set; }
        public string splicePointNo { get; set; }
        public string floorCLLI { get; set; }
        public string manhole { get; set; }
        public string manholeHandholeNo { get; set; }
        public string owner { get; set; }
        public string manholeHandholeOwnedBy { get; set; }
        public string manholeOwner { get; set; }
        public string mfnBackbone { get; set; }
        public string facilitiesEngineer { get; set; }
        public string otherSplicePointsInEnclosure { get; set; }
        public string dirFacilitiesMgmt { get; set; }
        public string dirFacilitiesMgmtDate { get; set; }
        public string dirNetworkEngineering { get; set; }
        public string dirNetworkEngineeringDate { get; set; }
        public string vpNetworkEngineering { get; set; }
        public string vpNetworkEngineeringDate { get; set; }
        public string floor { get; set; }
        public string racks { get; set; }
        public string otherSplicePointsInNode { get; set; }
        public string notes { get; set; }
        public string customerInfo { get; set; }
        public string handholeOrPoleNo { get; set; }
        public string manholeOrPoleNo { get; set; }
        public string otherSplicePointsInManhole { get; set; }
    }

    public class RowColIndexes
    {
        public int rowIndex { get; set; }
        public int colIndex { get; set; }
    }

    public class MapVal {
        public string fieldName { get; set; }
        public string fieldLabel { get; set; }

        public int relativePos { get; set; }

        public int offset { get; set; }
    }

    public enum RelativePos
    {
       D = 0,
       R = 1,
       U = 2,
       L = 3
    }
}