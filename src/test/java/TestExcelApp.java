import org.junit.Test;

public class TestExcelApp {
    private static final String PROD_FILE_NAME = "C:\\Workspace\\Barracuda OMS\\Order Book Migration\\Prod.xlsx";
    private static final String TEST_FILE_NAME = "C:\\Workspace\\Barracuda OMS\\Order Book Migration\\Test.xlsx";
    private ExcelFile prodExcelFile;
    private ExcelFile testExcelFile;
    private int numberofRowsinProd = 0;
    @Test
    public void testExcel() {

//        ExcelApp app = new ExcelApp();
//        app.readExcel();
        prodExcelFile = new ExcelFile(PROD_FILE_NAME);
        testExcelFile = new ExcelFile(TEST_FILE_NAME);
       numberofRowsinProd = prodExcelFile.getTotalNumberofRowswithoutHeader();
        this.compareColumn("Rate");


//        excelFile.getTotalNumberofColumns();
//        excelFile.getTotalNumberofRowswithoutHeader();
//        for(int id=1;id<=6;id++) {
//            int colIndex = excelFile.getColumnIndexOfAColumn("Col"+Integer.toString(id));
//            System.out.println("Column Index: " + colIndex);
//            double cellValue = excelFile.getNumericValueOfaCell(id,"Col"+Integer.toString(id));
//            System.out.println("Cell Value of Row: " + id + "is" + cellValue);
//        }
//        excelFile.createAMapofHeaderstoColumnIndex();
//        excelFile.readXLSXFile();
//        excelFile.readExcelColumn();
    }

    public void compareColumn(String column){
        int colmnIndexofRate = prodExcelFile.getColumnIndexOfAColumn(column);
        System.out.println("Col Number: " + (colmnIndexofRate+1));
        if(colmnIndexofRate >= 0){
            for(int rowIndex=1;rowIndex <= numberofRowsinProd;rowIndex++){
                double prodCellValue = prodExcelFile.getNumericValueOfaCell(rowIndex,column);
                double testCellValue = testExcelFile.getNumericValueOfaCell(rowIndex,column);
                if((Double.MAX_VALUE != prodCellValue) && (Double.MAX_VALUE != testCellValue )){
                    if(compareRateColumn(prodCellValue,testCellValue)){
                        System.out.println("Row Number: " + (rowIndex+1) + " Prod Value" + prodCellValue + " Test Value" + testCellValue + "Result:PASS");
                    }else{
                        System.out.println("Row Number: " + (rowIndex+1) + " Prod Value" + prodCellValue + " Test Value" + testCellValue + "Result:FAIL");
                    }
                }else{
                    System.out.println("Row Number: " + (rowIndex+1) + " Prod Value" + prodCellValue + " Test Value" + testCellValue + "Result:Invalid");
                }

            }
        }else{
            System.out.println("Invalid Column Index for Rate " + colmnIndexofRate);
        }
    }

    public boolean compareRateColumn(double pRate, double tRate){
        double value = tRate - 1.0;
        if (pRate == value){
            return true;
        } else{
            return false;
        }
    }

}
