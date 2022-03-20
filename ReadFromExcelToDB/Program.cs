// See https://aka.ms/new-console-template for more information

using System.Data;
using System.Data.OleDb;
using System.Text;

Console.WriteLine("Hello, World!");
string path = @"C:\Users\themagicalplace\Documents\test_gamma.xlsx";
string connString = $"Provider= Microsoft.ACE.OLEDB.12.0;" + $"Data Source={path}" + ";Extended Properties='Excel 8.0;HDR=Yes'";
OleDbConnection oledbConn = new OleDbConnection(connString);
oledbConn.Open();
OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Sheet1$]", oledbConn);
// Create new OleDbDataAdapter
OleDbDataAdapter oleda = new OleDbDataAdapter();

oleda.SelectCommand = cmd;
DataSet ds = new DataSet();
oleda.Fill(ds, "Employees");


foreach(var m in ds.Tables[0].DefaultView)
{
    StringBuilder bids = new StringBuilder("");
    StringBuilder asks = new StringBuilder("");
    int i = 0;
    var row = ((System.Data.DataRowView)m).Row.ItemArray;

    //bids.Append($"{row[14]}".PadLeft(15));

    foreach (var r in row)
    {
             if(r.ToString() != "")
                 bids.Append($"{r}".PadLeft(15));
    }
    
    // while (i < 14)
    // {
    //     if(row[i].ToString() != "")
    //         bids.Append($"{((System.Data.DataRowView)m).Row.ItemArray[i]} ".PadLeft(15));
    //     i++;
    // }
    //
    // i = 0;
    // while (i < row.Length - 15)
    // {
    //     bids.Append($"{((System.Data.DataRowView)m).Row.ItemArray[i+15]} ".PadLeft(15));
    //     i++;
    // }
    
    Console.WriteLine(bids.ToString());
}
oledbConn.Close();