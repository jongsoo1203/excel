import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

public class ExcelReport {
    private int rows; // number of rows in the sheet like 1, 2, 3, etc.
    private int cols; // number of columns in the sheet like A, B, C, etc.
    private Cell cell; // cell object like A3, B4, etc.
    private Date date; // date object like Jan 2023, Feb 2023, etc.
    private static String formattedDate; // formatted date like Jan-2023, Feb-2023, etc.
    private String name; // name of the employee.
    private int sheetNumber; // the sheet number.
    private String[] companyNameList; // 건명
    private List<Double> workRatio; // 업무비율

    public ExcelReport() {
        this.rows = 0;
        this.cols = 0;
        this.cell = null;
        this.date = null;
        this.formattedDate = "";
        this.name = "";
        this.sheetNumber = 0;
        this.companyNameList = null;
        this.workRatio = null;
    }
    // 업무비율 가져오기
    public List<Double> workRatio(Sheet sheet) {
        List<Double> workRatio = new ArrayList<>();
        this.cell = sheet.getRow(6).getCell(7);
        // Check the cell type before retrieving the value
        while (cell.getCellType() != CellType.BLANK) {
            if (cell.getCellType() == CellType.FORMULA) {
                // Handle numeric value
                workRatio.add(cell.getNumericCellValue());
            }
            cell = sheet.getRow(cell.getRowIndex() + 1).getCell(7);
        }
        return workRatio;
    }
    // 건명 가져오기
    public String[] getCompanyNameList(Sheet sheet) {
        List<String> companyNameList = new ArrayList<>();
        this.cell = sheet.getRow(6).getCell(0);
        // get the name next to "이름" cell.
        while (!cell.getStringCellValue().equals("")) {
            companyNameList.add(cell.getStringCellValue());
            cell = sheet.getRow(cell.getRowIndex() + 1).getCell(0);
        }
        // Convert the ArrayList to an array
        String[] companyNameArray = companyNameList.toArray(new String[0]);

        return companyNameArray;
    }
    // 이름 가져오기
    public String getName(Sheet sheet) {
        // get the name next to "이름" cell.
        this.cell = sheet.getRow(3).getCell(6);
        // if the cell type is string, then get the name from the cell.
        if (cell.getCellType() == CellType.STRING) {
            name = cell.getStringCellValue();
        }

        return name;
    }

    // 날짜 가져오기
    public String getDate(Sheet sheet) {
        // get the date from the cell A3.
        this.cell = sheet.getRow(2).getCell(0);

        // if the cell type is numeric, then get the date from the cell.
        if (cell.getCellType() == CellType.NUMERIC) {
            date = cell.getDateCellValue();
        } else if (cell.getCellType() == CellType.STRING) {
            // if the cell type is string, then get the date from the cell.
            String dateString = cell.getStringCellValue();
            // parse the string date to date object.
            // parse is used to convert string to date object.
            try {
                // date format: yyyy-MM-dd
                date = new SimpleDateFormat("yyyy-MM-dd").parse(dateString);
            } catch (Exception e) {
                // if the date format is not yyyy-MM-dd, then print the error.
                e.printStackTrace();
            }
        }
        SimpleDateFormat dateFormat = new SimpleDateFormat("MMM-yyyy");
        formattedDate = dateFormat.format(date);
        return formattedDate;
    }

    // 행 크기 가져오기
    public int getRowSize(Sheet sheet) {
        return sheet.getLastRowNum() + 1;
    }
    // 열 크기 가져오기
    public int getColSize(Sheet sheet) {
        XSSFRow firstRow = (XSSFRow) sheet.getRow(0);
        return firstRow == null ? 0 : firstRow.getLastCellNum();
    }
    // 엑셀 파일 읽기
    public void readExcelFile(String filePath) throws IOException {
        FileInputStream file = new FileInputStream(filePath);
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);

            rows = getRowSize(sheet);
            cols = getColSize(sheet);

            // return all the rows
            Iterator rowiterator = sheet.iterator();

            while(rowiterator.hasNext()) {
                // retrun the first row
                XSSFRow row = (XSSFRow) rowiterator.next();
                // in that row get all the cells.
                Iterator cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    XSSFCell cell = (XSSFCell) cellIterator.next();
                    switch(cell.getCellType()) {
                        case STRING: System.out.print(cell.getStringCellValue()); break;
                        case NUMERIC: System.out.print(cell.getNumericCellValue()); break;
                        case BOOLEAN: System.out.print(cell.getBooleanCellValue()); break;
                        case FORMULA: System.out.print(cell.getNumericCellValue()); break;
                    }
                    System.out.print(" | ");
                }
                System.out.println();
            }

            formattedDate = getDate(sheet);
            name = getName(sheet);
            companyNameList = getCompanyNameList(sheet);
            workRatio = workRatio(sheet);
            System.out.println("Date: " + formattedDate);
            System.out.println("rows: " + rows);
            System.out.println("cols: " + cols);
            System.out.println("Name: " + getName(sheet));
            System.out.println("Company Name List: " + Arrays.toString(companyNameList));
            System.out.println("Work Ratio: " + workRatio(sheet));
            generateExcelReport(formattedDate);
            workbook.close();
            file.close();
        }
    }

    public void generateExcelReport(String formattedDate) {
        String fileName = formattedDate + ".xlsx"; // 새로운 파일 이름
        String filePath = "C:\\Users\\user\\Desktop\\"; // 새로운 파일 들어갈 곳

        try {
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet newSheet = workbook.createSheet("Sheet1");

            Row row = newSheet.createRow(0);
            Cell cell = row.createCell(0);
            cell.setCellValue("인건비 원가반영 테이블");

            // create a new col for each company name in the second row.
            row = newSheet.createRow(1);
            for(int i = 0; i < companyNameList.length; i++) {
                cell = row.createCell(i + 2);
                cell.setCellValue(companyNameList[i]);
            }

            row = newSheet.createRow(2);
            for(int i = 0; i < workRatio.size(); i++) {
                cell = row.createCell(i + 2);
                cell.setCellValue(workRatio.get(i));
            }

            cell = row.createCell(0);
            cell.setCellValue(name);

            FileOutputStream fileOutputStream = new FileOutputStream(filePath + fileName);
            workbook.write(fileOutputStream);
            fileOutputStream.close();

            System.out.println("New Excel file created: " + filePath + fileName);

            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Failed to create the Excel file.");
        }
    }

    public static void main(String[] args) {
        String path = "C:\\Users\\user\\Desktop\\test2.xlsx";
        ExcelReport reader = new ExcelReport();
        // Check if at least one argument is provided
        try {
            int sheetNumber = 0;
            reader.readExcelFile(path);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
