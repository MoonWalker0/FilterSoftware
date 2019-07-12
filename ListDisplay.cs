using System.Data.SQLite;

namespace TelevendFilter
{
    class ListDisplay
    {
        public struct ListItem
        {
            public string ItemWorkerID { get; set; }
            public string ItemStickerID { get; set; } 
            public string ItemID { get; set; }
            public string ItemDate { get; set; }  
            public string ItemPurchase { get; set; }
            public string ItemProduct { get; set; }
            public string ItemMachine { get; set; }
            // public string ItemReload { get; set; }
        }

        public static ListItem FormListItem(SQLiteDataReader data)
        {
            return new ListItem
            {
                ItemWorkerID = data["WorkerID"].ToString(),
                ItemStickerID = data["StickerID"].ToString(),
                ItemID = data["AccountID"].ToString(), 
                ItemDate = data["Date"].ToString(), 
                ItemPurchase = data.GetFloat(data.GetOrdinal("Paid")).ToString("0.00"),
                ItemProduct = data["Product"].ToString(),
                ItemMachine = data["Machine"].ToString()
                //ItemReload = data.GetFloat(data.GetOrdinal("Reload")).ToString("0.00")
            };
        }
    }
}
