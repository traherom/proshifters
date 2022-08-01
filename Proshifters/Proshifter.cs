using System.Collections.Immutable;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Tuple = System.Tuple;

namespace Traherom.Proshifters;

record Month
{
    public string Name = "";

    public int StartColIdx = -1;

    // public int DayCount = -1;
    public List<string> Days = new();
};

record Person(string Name, List<string> Row)
{
    public Dictionary<string, int> ShiftCountsForMonth(Month month)
    {
        var shifts = new Dictionary<string, int>();
        foreach (var name in Proshifter.ValidShiftNames)
            shifts[name] = 0;

        for (var shiftColIdx = month.StartColIdx; shiftColIdx < month.StartColIdx + month.Days.Count; shiftColIdx++)
        {
            // No more data for this person?
            if (shiftColIdx >= Row.Count)
                break;

            // Cleanup
            var shiftIndicator = Row[shiftColIdx].Trim().ToUpperInvariant();
            shiftIndicator = shiftIndicator.Replace("↓", "");
            shiftIndicator = shiftIndicator.Replace("↑", "");
            shiftIndicator = shiftIndicator.Replace("D2", "D12");
            shiftIndicator = shiftIndicator.Replace("S2", "S12");
            shiftIndicator = shiftIndicator.Replace("M2", "M12");

            var isActualShift = Proshifter.ValidShiftNames.Contains(shiftIndicator);
            if (isActualShift)
                shifts[shiftIndicator]++;

            // Is a weekend?
            if (isActualShift && month.Days[shiftColIdx - month.StartColIdx] == "S")
                shifts[Proshifter.WeekendShiftName]++;
        }

        return shifts;
    }
}

public class Proshifter
{
    public static readonly string WeekendShiftName = "Weekend";

    public static readonly string[] ValidShiftNames = new[]
        { WeekendShiftName, "D", "D10", "D12", "S", "S10", "S12", "M", "M10", "M12", "FF", "EV", "FPC" };

    private List<List<string>>? _scheduleRawData;

    private List<string> MonthRow => _scheduleRawData?[0] ?? throw new NullReferenceException();
    private List<string> DayNameRow => _scheduleRawData?[1] ?? throw new NullReferenceException();
    private List<string> DayNumberRow => _scheduleRawData?[2] ?? throw new NullReferenceException();

    private List<string> NameColumn => _scheduleRawData?.Select(x => x.Take(7).LastOrDefault() ?? "").ToList() ??
                                       throw new NullReferenceException();

    private List<string> IsProColumn => _scheduleRawData?.Select(x => x.Take(4).LastOrDefault() ?? "").ToList() ??
                                        throw new NullReferenceException();

    private List<Month>? _months;
    private List<Person>? _people;

    public void Calculate(string schedulePath, string outputPath)
    {
        ReadWorksheet(schedulePath);
        GetMonthRanges();
        GetPeopleRows();
        WriteWorksheet(outputPath);
    }

    private void GetPeopleRows()
    {
        if (_scheduleRawData is null)
            throw new NullReferenceException();

        _people = new();

        var names = NameColumn;
        var isPros = IsProColumn;

        for (var rowIdx = 0; rowIdx < names.Count; rowIdx++)
        {
            var name = names[rowIdx];
            var isPro = isPros[rowIdx];

            // Skip rows which don't actually have names
            if (string.IsNullOrWhiteSpace(name))
                continue;

            // Skip rows for non-proshifters
            if (isPro.Trim().ToUpperInvariant() != "Y")
                continue;

            _people.Add(new(name, _scheduleRawData[rowIdx]));
        }
    }

    private void GetMonthRanges()
    {
        if (_scheduleRawData == null)
            throw new NullReferenceException("Schedule raw data not set yet");

        _months = new();
        int currMonthStart = -1;
        for (var colIdx = 0; colIdx < DayNumberRow.Count; colIdx++)
        {
            var cell = DayNumberRow[colIdx].Trim();

            // When we hit day "1", we've started a mew month. Save off the previous month
            // Alternatively, if we hit a blank cell, then we've reached the end of days
            // The only caveat to that is if we haven't actually started days yet, in which case
            // just keep on trucking
            if (cell is "" or "1")
            {
                if (currMonthStart >= 0)
                {
                    var dayCount = colIdx - currMonthStart;
                    _months.Add(
                        new()
                        {
                            Name = MonthRow[currMonthStart],
                            StartColIdx = currMonthStart,
                            // DayCount = dayCount,
                            Days = DayNameRow.Skip(currMonthStart).Take(dayCount)
                                .Select(x => x.Trim().ToUpperInvariant()).ToList(),
                        });
                }

                currMonthStart = colIdx;
            }

            // If we're on a blank, we don't know when the month starts yet
            if (cell == "")
                currMonthStart = -1;
        }
    }

    /**
     * Read in schedule into cell x row multidimensional array
     */
    private void ReadWorksheet(string schedulePath)
    {
        var simpleSheetData = new List<List<string>>();

        using var doc = SpreadsheetDocument.Open(schedulePath, false);

        // Local schedule sheet
        var workbookPart = doc.WorkbookPart;
        var allSheets = workbookPart.Workbook.GetFirstChild<Sheets>();
        var scheduleSheet = allSheets.Select(s => s as Sheet).First(s => s?.Name == "Schedule") ??
                            throw new NullReferenceException("'Schedule' sheet not found");
        var scheduleSheetData =
            (workbookPart.GetPartById(scheduleSheet.Id) as WorksheetPart)?.Worksheet.GetFirstChild<SheetData>() ??
            throw new NullReferenceException("'Schedule' sheet data not found");

        // statement to get the worksheet object by using the sheet id
        var currentSimpleRow = new List<string>();
        foreach (var currentRow in scheduleSheetData.Select(x => x as Row))
        {
            if (currentRow is null)
                continue;

            foreach (var currentCell in currentRow.Select(x => x as Cell))
            {
                if (currentCell is null)
                {
                    currentSimpleRow.Add(String.Empty);
                }
                else if (currentCell.DataType is null)
                {
                    currentSimpleRow.Add(currentCell.InnerText ?? currentCell.InnerXml ?? String.Empty);
                }
                else if (currentCell.DataType == CellValues.SharedString)
                {
                    if (Int32.TryParse(currentCell.InnerText, out var id))
                    {
                        var item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>()
                            .ElementAt(id);
                        currentSimpleRow.Add(item?.Text?.Text ?? item?.InnerText ?? item?.InnerXml ?? "");
                    }
                    else
                    {
                        currentSimpleRow.Add(String.Empty);
                    }
                }
                else
                {
                    currentSimpleRow.Add(currentCell.InnerText ?? currentCell.InnerXml ?? String.Empty);
                }
            }

            simpleSheetData.Add(currentSimpleRow);
            currentSimpleRow = new List<string>();
        }

        _scheduleRawData = simpleSheetData.Select(x => x.ToList()).ToList();
    }

    private void WriteWorksheet(string outputPath)
    {
        if (_people is null)
            throw new NullReferenceException();
        if (_months is null)
            throw new NullReferenceException();

        foreach (var person in _people)
        {
            Console.WriteLine($"{person.Name}:");
            foreach (var month in _months)
            {
                var shifts = person.ShiftCountsForMonth(month);
                var dayShiftCount = shifts.ContainsKey("D") ? shifts["D"] : 0;
                Console.WriteLine("\tD: {0}", dayShiftCount);
            }
        }

        using var document = SpreadsheetDocument.Create(outputPath, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(sheetData);

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        var sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };

        sheets.Append(sheet);

        // Build header rows. Months are spaced out
        var monthRow = new Row();
        var shiftRow = new Row();
        monthRow.AppendChild(new Cell());
        shiftRow.AppendChild(new Cell { DataType = CellValues.String, CellValue = new("Name") });
        foreach (var month in _months)
        {
            // Add month name, then skip a bunch of cells when going through the shifts
            var cell = new Cell();
            cell.DataType = CellValues.String;
            cell.CellValue = new CellValue(month.Name);
            monthRow.AppendChild(cell);
            for (var i = 0; i < ValidShiftNames.Length - 1; i++)
                monthRow.AppendChild(new Cell());

            // Add each shift type in order for each month
            foreach (var shiftType in ValidShiftNames)
            {
                shiftRow.AppendChild(new Cell { DataType = CellValues.String, CellValue = new(shiftType) });
            }

            var mergeCells = new MergeCells();
            //append a MergeCell to the mergeCells for each set of merged cells
            mergeCells.Append(new MergeCell() { Reference = new StringValue("A1:F1") });
            // worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetData>().First());
            // monthRow.AppendChild(mergeCells);


            // monthRow.InsertAt(cell, 5);
        }

        sheetData.AppendChild(monthRow);
        sheetData.AppendChild(shiftRow);

        // Now put in each person
        foreach (var person in _people)
        {
            var personRow = new Row();
            personRow.AppendChild(new Cell { DataType = CellValues.String, CellValue = new(person.Name) });

            foreach (var month in _months)
            {
                var shiftsWorkedByPerson = person.ShiftCountsForMonth(month);
                
                foreach (var shiftType in ValidShiftNames)
                {
                    var count = shiftsWorkedByPerson[shiftType];
                    personRow.AppendChild(new Cell { DataType = CellValues.Number, CellValue = new(count.ToString()) });
                }
            }

            sheetData.AppendChild(personRow);
        }

        // All done with that mess, thank god
        workbookPart.Workbook.Save();
    }
}