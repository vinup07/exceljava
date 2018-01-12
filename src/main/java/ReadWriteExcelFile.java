import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ReadWriteExcelFile {

    private static final String FILE_NAME = "C:\\Workspace\\Barracuda OMS\\Order Book Migration\\Test.xlsx";
    private List<String> headers = new ArrayList<String>();
    private Map<String, String> orderBook = new HashMap<String, String>();

    public void readXLSXFile() {
        try {

            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();

            int rowCount = 1;
            boolean isfirstRow = true;
            while (iterator.hasNext()) {
                Row currentRow = iterator.next();
                if (isfirstRow) {
                    isfirstRow = false;
                    System.out.println("FIRSTRow");
                    Iterator<Cell> cellIterator = currentRow.iterator();
                    int colCnt=0;
                    while (cellIterator.hasNext()) {
                        //getCellTypeEnum shown as deprecated for version 3.15
                        //getCellTypeEnum ill be renamed to getCellType starting from version 4.0
                        Cell currentCell = cellIterator.next();
                        if (currentCell.getCellTypeEnum() == CellType.STRING) {
//                            System.out.print(currentCell.getStringCellValue() + "--");

                            headers.add(currentCell.getStringCellValue());
                            colCnt++;
                        } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
//                            System.out.print(currentCell.getNumericCellValue() + "--");
                        }
                    }
                    System.out.println("No of Columns:"+colCnt);
                    for (String header : headers) {
                        System.out.println("--" + header);
                    }
                }else {
                    System.out.println("secondRow");
                    Iterator<Cell> cellIterator = currentRow.iterator();
                    for (String header : headers) {
                        System.out.println("Header:"+header);
                        String value = "Empty";
                        Cell currentCell = cellIterator.next();
                        if (currentCell.getCellTypeEnum() == CellType.BLANK) {
                            value = "Empty";
                            System.out.println("Blank Type" + value);
                        }else if (currentCell.getCellTypeEnum() == CellType.STRING) {
                                value = currentCell.getStringCellValue();
                                System.out.println("String Type" + value);
                        } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                                value = Double.toString(currentCell.getNumericCellValue());
                                System.out.println("Num Type" + value);
                        }

                        System.out.println("Add to Orderbook - Key : " + header + " Value : " + value);
                        orderBook.put(header, value);
                    }

                    for (Map.Entry<String, String> entry : orderBook.entrySet()) {
                        System.out.println("Key : " + entry.getKey() + " Value : " + entry.getValue());
                    }
                }
            }

            } catch(FileNotFoundException e){
                e.printStackTrace();
            } catch(IOException e){
                e.printStackTrace();
            }
        }

        public void readExcelColumn() {

            try {
                FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
                Workbook workbook = new XSSFWorkbook(excelFile);
                Sheet sheet = workbook.getSheetAt(0);
//            Iterator<Row> iterator = datatypeSheet.iterator();
//                int colIndex = 1;
                for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    if (row != null) {
                        for (int colIndex = 0; colIndex < 3; colIndex++) {
                            Cell cell = row.getCell(colIndex);
                            String value = "Empty";
                            if (cell != null) {
                                // Found column and there is value in the cell.
                                if (cell.getCellTypeEnum() == CellType.BLANK) {
                                    value = "Empty";
//                                    System.out.println("Blank Type" + value);
                                } else if (cell.getCellTypeEnum() == CellType.STRING) {
                                    value = cell.getStringCellValue();
//                                    System.out.println("String Type" + value);
                                } else if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                                    value = Double.toString(cell.getNumericCellValue());
//                                    System.out.println("Num Type" + value);
                                }
                                System.out.println("Row:" + rowIndex + " Col:" + colIndex + " Value:" + value);
                            }
                        }
                    }
                }
        } catch(FileNotFoundException e){
        e.printStackTrace();
        } catch(IOException e){
        e.printStackTrace();
        }
    }
}
