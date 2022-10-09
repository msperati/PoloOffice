import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;

public class ExcelGenerator {

    private static final String MAIN_DIRECTORY = "..\\PoloOffice\\";

    private static final String OUTPUT_NAME_END = " da pagare.xlsx";

    public static void main(String[] args) throws IOException, InvalidFormatException {
        System.out.println("Searching for input file");
        final File folder = new File(MAIN_DIRECTORY);
        String path = null;
        for (final File f : folder.listFiles()) {
            if (f.getName().endsWith(".xlsx") && !f.getName().contains(OUTPUT_NAME_END)) {
                path = f.getName();
                break;
            }
        }
        if (path != null) {
            System.out.println("Starting reading file " + path);
            readInput(MAIN_DIRECTORY + path);
        }
    }

    private static void readInput(String path) throws IOException, InvalidFormatException {
        File inputFile = new File(path);

        /** INPUT */
        Workbook inputWorkbook = new XSSFWorkbook(inputFile);
        Sheet inputSheet = inputWorkbook.getSheetAt(0);

        /** OUTPUT */
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(inputSheet.getSheetName() + " OUTPUT");

        Map<String, Integer> map = new LinkedHashMap<>();
        List<Integer> indexOfGiornateDaPagare = new LinkedList<>();

        for (int i = 0; i <= inputSheet.getLastRowNum(); i++) {
            Row inputRow = inputSheet.getRow(i);
            if (i == 0) {
                Row row = sheet.createRow(0);
                copyRow(inputRow, row, true);
            } else {
                Cell cell = inputRow.getCell(0);
                String comune = cell.getStringCellValue().toUpperCase();
                if (!map.containsKey(comune)) {
                    map.put(comune, 1);
                } else {
                    map.put(comune, map.get(comune) + 1);
                    if (map.get(comune) >= 3) {
                        indexOfGiornateDaPagare.add(i);
                    }
                }
            }
        }

        for (Integer index : indexOfGiornateDaPagare) {
            Row inputRow = inputSheet.getRow(index);
            Row newRow = sheet.createRow(sheet.getPhysicalNumberOfRows());
            copyRow(inputRow, newRow, false);
        }

        String fileName = path.substring(path.lastIndexOf("\\") + 1, path.length() - 5);

        int columnsNum = sheet.getRow(0).getLastCellNum();
        for (int c = 0; c < columnsNum; c++) {
            sheet.autoSizeColumn(c);
        }
        /** Write the output to a file */
        FileOutputStream fileOut = new FileOutputStream(fileName + OUTPUT_NAME_END);
        workbook.write(fileOut);
        fileOut.close();

        System.out.println("Closing the output file");

        /** Closing the workbook */
        workbook.close();

        System.out.println("Output file closed");
    }

    private static void copyRow(Row oldRow, Row newRow, boolean isHeader) {
        for (int c = 0; c < oldRow.getPhysicalNumberOfCells(); c++) {
            Cell cell = newRow.createCell(c);
            if (oldRow.getCell(c).getCellType().equals(CellType.STRING)) {
                cell.setCellValue(oldRow.getCell(c).getStringCellValue().toUpperCase());
            } else {
                Date date = oldRow.getCell(c).getDateCellValue();
                try {
                    DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
                    String cellValue = df.format(date);
                    cell.setCellValue(cellValue);
                } catch (Exception e) {
                    System.out.println("Impossibile recuperare la data per " + date);
                    System.out.println(e.getMessage());
                }
            }

            if (isHeader) {
                Workbook workbook = newRow.getSheet().getWorkbook();

                /** Create a Font for styling header cells */
                Font headerFont = workbook.createFont();
                headerFont.setBold(true);
                headerFont.setFontHeightInPoints((short) 14);
                headerFont.setColor(IndexedColors.WHITE.index);

                /** Create a CellStyle with the font */
                CellStyle headerCellStyle = workbook.createCellStyle();
                headerCellStyle.setFont(headerFont);
                headerCellStyle.setFillForegroundColor(IndexedColors.GREEN.index);
                headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

                cell.setCellStyle(headerCellStyle);
            }

        }
    }

}
