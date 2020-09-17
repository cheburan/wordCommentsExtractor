using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace WordsCommentsExtractor
{
	public class ExcelDocument
	{
		public string filePath;
		private string[] titles;
		bool exist = false;

		public ExcelDocument(string _path, string[] _titles)
		{
			filePath = _path;
			titles = _titles;

			if (File.Exists(_path))
			{
				// File exist
				exist = true;
				//Open file later
			}
			else
			{
				exist = false;
			}
		}

		public void Create()
		{
			if (exist)
			{
				//Open for writing
			} else
			{
				// Create a spreadsheet document by supplying the filepath.
				// By default, AutoSave = true, Editable = true, and Type = xlsx.
				SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
					Create(filePath, SpreadsheetDocumentType.Workbook);

				// Add a WorkbookPart to the document.
				WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
				workbookpart.Workbook = new Workbook();

				// Add a WorksheetPart to the WorkbookPart.
				WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
				worksheetPart.Worksheet = new Worksheet(new SheetData());

				// Add Sheets to the Workbook.
				Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
					AppendChild<Sheets>(new Sheets());

				//// Append a new worksheet and associate it with the workbook.
				//Sheet sheet = new Sheet()
				//{
				//	Id = spreadsheetDocument.WorkbookPart.
				//	GetIdOfPart(worksheetPart),
				//	SheetId = 1,
				//	Name = "mySheet"
				//};
				//sheets.Append(sheet);

				workbookpart.Workbook.Save();

				// Close the document.
				spreadsheetDocument.Close();
			}
		}

		// Given a document name, inserts a new worksheet.
		public void InsertNewBlankWorksheet(string NewSheetName)
		{
			// Open the document for editing.
			using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(filePath, true))
			{
				// Add a blank WorksheetPart.
				WorksheetPart newWorksheetPart = spreadSheet.WorkbookPart.AddNewPart<WorksheetPart>();
				newWorksheetPart.Worksheet = new Worksheet(new SheetData());

				Sheets sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>();
				string relationshipId = spreadSheet.WorkbookPart.GetIdOfPart(newWorksheetPart);

				// Get a unique ID for the new worksheet.
				uint sheetId = 1;
				if (sheets.Elements<Sheet>().Count() > 0)
				{
					sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
				}

				// Give the new worksheet a name.
				string sheetName = NewSheetName;

				// Append the new worksheet and associate it with the workbook.
				Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
				sheets.Append(sheet);
			}
		}

		// Given a document name and text, 
		// inserts a new worksheet and writes the text to specified cell of the new worksheet.
		public void InsertText(string sheetName, string[] cells, string[][] data)
		{
			// Open the document for editing.
			using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(filePath, true))
			{
				// Get the SharedStringTablePart. If it does not exist, create a new one.
				SharedStringTablePart shareStringPart;
				if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
				{
					shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
				}
				else
				{
					shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
				}

				// Insert a new worksheet.
				WorksheetPart worksheetPart = InsertWorksheet(spreadSheet.WorkbookPart, sheetName);

				//Insert Titles and Data
				for (int i = 0; i < data.Length; i++)
				{
					string cellName;
					for (int j = 0; j < data[i].Length; j++)
					{
						cellName = cells[j];
						// Insert the text into the SharedStringTablePart.
						int index = InsertSharedStringItem(data[i][j], shareStringPart);

						// Insert cell A1 into the new worksheet.
						Cell cell = InsertCellInWorksheet(cellName, (uint)i+1, worksheetPart);

						// Set the value of cell A1.
						cell.CellValue = new CellValue(index.ToString());
						cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

					}
					// Save the new worksheet.
					worksheetPart.Worksheet.Save();
				}
			}
		}

		// Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
		// and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
		private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
		{
			// If the part does not contain a SharedStringTable, create one.
			if (shareStringPart.SharedStringTable == null)
			{
				shareStringPart.SharedStringTable = new SharedStringTable();
			}

			int i = 0;

			// Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
			foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
			{
				if (item.InnerText == text)
				{
					return i;
				}

				i++;
			}

			// The text does not exist in the part. Create the SharedStringItem and return its index.
			shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
			shareStringPart.SharedStringTable.Save();

			return i;
		}

		// Given a WorkbookPart, inserts a new worksheet.
		private static WorksheetPart InsertWorksheet(WorkbookPart workbookPart, string NewSheetName)
		{
			// Add a new worksheet part to the workbook.
			WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
			newWorksheetPart.Worksheet = new Worksheet(new SheetData());
			newWorksheetPart.Worksheet.Save();

			Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
			string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

			// Get a unique ID for the new sheet.
			uint sheetId = 1;
			if (sheets.Elements<Sheet>().Count() > 0)
			{
				sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
			}

			string sheetName = NewSheetName;

			// Append the new worksheet and associate it with the workbook.
			Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
			sheets.Append(sheet);
			workbookPart.Workbook.Save();

			return newWorksheetPart;
		}

		// Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
		// If the cell already exists, returns it. 
		private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
		{
			Worksheet worksheet = worksheetPart.Worksheet;
			SheetData sheetData = worksheet.GetFirstChild<SheetData>();
			string cellReference = columnName + rowIndex;

			// If the worksheet does not contain a row with the specified row index, insert one.
			Row row;
			if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
			{
				row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
			}
			else
			{
				row = new Row() { RowIndex = rowIndex };
				sheetData.Append(row);
			}

			// If there is not a cell with the specified column name, insert one.  
			if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
			{
				return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
			}
			else
			{
				// Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
				Cell refCell = null;
				foreach (Cell cell in row.Elements<Cell>())
				{
					if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
					{
						refCell = cell;
						break;
					}
				}

				Cell newCell = new Cell() { CellReference = cellReference };
				row.InsertBefore(newCell, refCell);

				worksheet.Save();
				return newCell;
			}
		}
	}
}
