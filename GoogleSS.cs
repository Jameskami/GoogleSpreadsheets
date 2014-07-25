using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Google.GData.Client;
using Google.GData.Extensions;
using Google.GData.Spreadsheets;

namespace GoogleIntegration
{
    public class GoogleSpreadsheet
    {
        SpreadsheetsService service;
        SpreadsheetEntry spreadsheet;
        SpreadsheetFeed feed;
        AtomLink atomLink;
        WorksheetQuery workSheetQuery;
        WorksheetFeed workSheetFeed;
        WorksheetEntry worksheetEntry;
        AtomLink cellLink;
        SpreadsheetQuery query;
        CellQuery cellQuery;
        CellFeed cellFeed;

        List<CellAddress> cellAdresses = new List<CellAddress>();

        public GoogleSpreadsheet(string username, string password, int spreadSheetNumber)
        {
            service = new SpreadsheetsService("test-app");
            service.setUserCredentials(GoogleIntegration.Properties.Resources.username, GoogleIntegration.Properties.Resources.password);
            query = new SpreadsheetQuery();
            feed = service.Query(query);
            spreadsheet = (SpreadsheetEntry)feed.Entries[spreadSheetNumber];
            atomLink = spreadsheet.Links.FindService(GDataSpreadsheetsNameTable.WorksheetRel, null);
            workSheetQuery = new WorksheetQuery(atomLink.HRef.ToString());
            workSheetFeed = service.Query(workSheetQuery);
            worksheetEntry = (WorksheetEntry)workSheetFeed.Entries[spreadSheetNumber];
            cellLink = worksheetEntry.Links.FindService(GDataSpreadsheetsNameTable.CellRel, null);
            cellQuery = new CellQuery(cellLink.HRef.ToString());
            cellFeed = service.Query(cellQuery);
        }
        Dictionary<String, CellEntry> GetCellEntryMap(List<CellAddress> cellAddrs)
        {
            CellFeed batchRequest = new CellFeed(new Uri(cellFeed.Self), service);
            foreach (CellAddress cellId in cellAddrs)
            {
                CellEntry batchEntry = new CellEntry(cellId.Row, cellId.Col, cellId.IdString);
                batchEntry.Id = new AtomId(string.Format("{0}/{1}", cellFeed.Self, cellId.IdString));
                batchEntry.BatchData = new GDataBatchEntryData(cellId.IdString, GDataBatchOperationType.query);
                batchRequest.Entries.Add(batchEntry);
            }

            CellFeed queryBatchResponse = (CellFeed)service.Batch(batchRequest, new Uri(cellFeed.Batch));

            Dictionary<String, CellEntry> cellEntryMap = new Dictionary<String, CellEntry>();
            foreach (CellEntry entry in queryBatchResponse.Entries)
            {
                cellEntryMap.Add(entry.BatchData.Id, entry);
            }

            return cellEntryMap;
        }
        public void setCells(IEnumerable<string> list, SheetDimensions dimension)
        {
            string[] strs = list.ToArray();

            for (uint row = 1; row <= dimension.rows; ++row)
            {
                for (uint col = 1; col <= dimension.columns; ++col)
                {
                    cellAdresses.Add(new CellAddress(row, col));
                }
            }
            Dictionary<String, CellEntry> cellEntries = GetCellEntryMap(cellAdresses);
            CellFeed batchRequest = new CellFeed(cellQuery.Uri, service);
            int i = 0;
            foreach (CellAddress cellAddress in cellAdresses)
            {
                CellEntry batchEntry = cellEntries[cellAddress.IdString];
                batchEntry.InputValue = strs[i];
                batchEntry.BatchData = new GDataBatchEntryData(strs[i], GDataBatchOperationType.update);
                batchRequest.Entries.Add(batchEntry);
                i++;
            }
            CellFeed batchResponse = (CellFeed)service.Batch(batchRequest, new Uri(cellFeed.Batch));
        }
        public void ClearSheet()
        {
            foreach (CellEntry cell in cellFeed.Entries)
            {
                cell.InputValue = "";
                cell.Update();
            }
        }
        public List<string> GetCellListString()
        {
            var strs = new List<string>();
            foreach (CellEntry cell in cellFeed.Entries)
            {
                strs.Add(cell.Value);
            }
            return strs;
        }
        public List<int> GetCellListInt()
        {
            var ints = new List<int>();
            foreach (CellEntry cell in cellFeed.Entries)
            {
                ints.Add(int.Parse(cell.Value));
            }
            return ints;
        }
        public List<double> GetCellListDouble()
        {
            var doubles = new List<double>();
            foreach (CellEntry cell in cellFeed.Entries)
            {
                doubles.Add(double.Parse(cell.Value));
            }
            return doubles;
        }
        public List<string> ToStringList<T>(List<T> list)
        {
            var stringList = new List<string>();
            foreach (T i in list)
            {
                stringList.Add(i.ToString());
            }
            return stringList;
        }
    }

    internal struct CellAddress
    {
        public uint Row;
        public uint Col;
        public string IdString;

        public CellAddress(uint row, uint col)
        {
            this.Row = row;
            this.Col = col;
            this.IdString = string.Format("R{0}C{1}", row, col);
        }
    }
    public class SheetDimensions
    {
        public int rows = 1;
        public int columns = 1;
        public SheetDimensions(int rows = 1, int columns = 1)
        {
            if (rows > 0 && columns > 0)
            {
                this.rows = rows;
                this.columns = columns;
            }
        }
    }
}
