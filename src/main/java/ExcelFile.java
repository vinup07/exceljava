import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;
public class ExcelFile {

    private Workbook workbook;
    private Sheet sheet;
    private List<String> listOfHeaders = new ArrayList<String>();
    private int totalNumberofColumns = 0;

    public ExcelFile(String fileName) {
        try {

            FileInputStream excelFile = new FileInputStream(new File(fileName));
            workbook = new XSSFWorkbook(excelFile);
            sheet = workbook.getSheetAt(0);
            createAMapofHeaderstoColumnIndex();
        }catch(FileNotFoundException e){
            e.printStackTrace();
        } catch(IOException e){
            e.printStackTrace();
        }
    }

    private void createAMapofHeaderstoColumnIndex(){
        int rowIndex = 0;
        String header = null;
        Row currentRow = sheet.getRow(rowIndex);
        totalNumberofColumns = currentRow.getPhysicalNumberOfCells();
//        System.out.println("First Cell Num: " + currentRow.getFirstCellNum());
//        System.out.println("last Cell Num: " + currentRow.getLastCellNum());
//        System.out.println("Total Number of physical ceels" + currentRow.getPhysicalNumberOfCells());
        if(currentRow!= null){
            Iterator<Cell> cellIterator = currentRow.iterator();
            while (cellIterator.hasNext()) {
                Cell currentCell = cellIterator.next();
                if(currentCell != null){
//                    numberofColumns++;
                    if (currentCell.getCellTypeEnum() == CellType.BLANK) {
                        header = "";
//                        System.out.println("Blank Type" + header);
                    }else if (currentCell.getCellTypeEnum() == CellType.STRING) {
                        header = currentCell.getStringCellValue();
//                        System.out.println("String Type" + header);
                    } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                        header = Double.toString(currentCell.getNumericCellValue());
//                        System.out.println("Num Type" + header);
                    }
                }
                listOfHeaders.add(header);
            }
        }
//        System.out.println("Number of Columns: " + listOfHeaders.size());
//        for (String head : listOfHeaders) {
//            System.out.println("--" + head);
//        }
    }

    public int getTotalNumberofColumns(){
        System.out.println("Total Number of Columns: "+ totalNumberofColumns);
        return totalNumberofColumns;

    }

    public int getTotalNumberofRowswithoutHeader(){
        System.out.println("Total Number of Rows: "+ sheet.getLastRowNum());
        return sheet.getLastRowNum();

    }

    public int getColumnIndexOfAColumn(String columnHeader){
        return listOfHeaders.indexOf(columnHeader);
    }

    public double getNumericValueOfaCell(int rowIndex, String columnHeader){
        double cellValue = Double.MAX_VALUE;
        Row currentRow = sheet.getRow(rowIndex);
        if (null == currentRow) {
            System.out.println("Error:Retrieving Row");
        }else {
            int columnIndex = getColumnIndexOfAColumn(columnHeader);
            Cell currentCell = currentRow.getCell(columnIndex);
            if (currentCell == null) {
                System.out.println("Error:Retrieving Column");
            }
            else {
                if (currentCell.getCellTypeEnum() == CellType.BLANK) {
                    cellValue = Double.MAX_VALUE;
                        System.out.println("Error: Blank Cell");
                }else if (currentCell.getCellTypeEnum() == CellType.STRING) {
                    cellValue = Double.MAX_VALUE;
                    System.out.println("Error: Incorrect Data Type in the input file");
                } else {
                    cellValue = currentCell.getNumericCellValue();
                }
            }
        }
        return cellValue;
    }
}
