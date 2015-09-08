package ro.anproca.read;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by Alex Proca <alex.proca@gmail.com> on 31/07/15.
 */
public class ExcelSheet {

    private Map<Integer, Map<Short, Cell>> table = new HashMap<>();
    private List<Row> rows = new ArrayList<>();

    public ExcelSheet(Workbook wb, int position) throws IOException {
        this(wb.getSheetAt(position));
    }

    public ExcelSheet(Workbook wb, String sheetName) throws IOException {
        this(wb.getSheet(sheetName));
    }

    public ExcelSheet(Sheet sheet) throws IOException {
        for (Row row : sheet) {
            rows.add(row.getRowNum(), row);
            for (Cell cell : row) {
                CellReference cellRef = new CellReference(cell);
                addInTable(cellRef.getRow(), cellRef.getCol(), cell);
            }
        }
    }

    private Cell addInTable(Integer row, Short column, Cell value) {
        Map<Short, Cell> rowMap = table.get(row);

        if (rowMap == null) {
            rowMap = new HashMap<>();
            table.put(row, rowMap);
        }

        return rowMap.put(column, value);
    }

    private Cell getCellFromTable(Integer row, Short column) {

        if (table.get(row) == null) {
            return null;
        }

        return table.get(row).get(column);
    }

    public Row getRow(int line) {
        return rows.get(line);
    }

    public int getNumberOfRows() {
        return rows.size();
    }

    public Cell getCell(String column, int row) {
        return getCell(column + row);
    }

    public Cell getCell(String stringReference) {
        CellReference reference = new CellReference(stringReference);
        return getCellFromTable(reference.getRow(), reference.getCol());
    }
}
