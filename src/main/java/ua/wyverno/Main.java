package ua.wyverno;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileWriter;
import java.io.IOException;
import java.net.URI;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Iterator;

/**
 * Hello world!
 *
 */
public class Main
{
    private static final Logger logger = LoggerFactory.getLogger(Main.class);
    private static final Path pathFolderCSV = Paths.get("CSV");

    public static void main(String[] args) throws IOException {
        logger.info("Start downloading table");
        Workbook workbook = new XSSFWorkbook(URI.create(args[0]).toURL().openStream()); // URL to Source XLSX File
        logger.info("Finish download table");
        convertXLSXToCSV(workbook);
        logger.info("Finish converting table");
    }

    private static String escapeSpecialCharacters(String cellValue) {
        String escapedValue = cellValue;
        if (cellValue.contains(",") || cellValue.contains("\"") || cellValue.contains("\n")) {
            escapedValue = cellValue.replace("\"", "\"\"");
            escapedValue = "\"" + escapedValue + "\"";
        }
        return escapedValue;
    }

    public static void convertXLSXToCSV(Workbook workbook) throws IOException {
        if (pathFolderCSV.toFile().exists()) FileUtils.cleanDirectory(pathFolderCSV.toFile());

        Iterator<Sheet> sheetIterator = workbook.sheetIterator();

        DataFormatter formatter = new DataFormatter();
        while (sheetIterator.hasNext()) {
            if (pathFolderCSV.toFile().exists() || pathFolderCSV.toFile().mkdirs()) {
                Sheet sheet = sheetIterator.next();
                Path fileCsvPath = pathFolderCSV.resolve(sheet.getSheetName() + ".csv");

                try (FileWriter csvWriter = new FileWriter(fileCsvPath.toFile())) {
                    logger.info("Convert sheet to csv: {}", sheet.getSheetName());
                    int lastRowWithValue = lastRowWithValue(sheet)+1;
                    for (int i = 0; i < lastRowWithValue; i++) {
                        Row row = sheet.getRow(i);
                        int lastCellNumWithValue = lastCellNumWithValue(row)+1;
                        StringBuilder csvRowBuilder = new StringBuilder();
                        for (int j = 0; j < lastCellNumWithValue; j++) {
                            Cell cell = row.getCell(j);
                            csvRowBuilder.append(escapeSpecialCharacters(formatter.formatCellValue(cell)));
                            if (j < lastCellNumWithValue-1) csvRowBuilder.append(",");
                        }
                        csvRowBuilder.append("\n");
                        csvWriter.write(csvRowBuilder.toString());
                    }

                }
            }
        }
    }

    private static int lastRowWithValue(Sheet sheet) {
        int result = 0;
        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            if (lastCellNumWithValue(row) > 0) {
                result = row.getRowNum();
            }
        }
        return result;
    }

    private static int lastCellNumWithValue(Row row) {
        int result = 0;
        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                result = cell.getColumnIndex();
            }
        }
        return result;
    }
}
