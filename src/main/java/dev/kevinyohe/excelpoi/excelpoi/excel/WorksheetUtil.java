package dev.kevinyohe.excelpoi.excelpoi.excel;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.WorkbookUtil;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class WorksheetUtil {


    public Workbook getNewWorkbook(String filename, String sheetName) throws IOException {
        Workbook wb = new HSSFWorkbook();
        try  (OutputStream fileOut = new FileOutputStream(filename)) {
            wb.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
        Sheet sheet1 = wb.createSheet(WorkbookUtil.createSafeSheetName(sheetName));
        OutputStream fileOut = new FileOutputStream("workbook.xls");
        wb.write(fileOut);
        return wb;
    }


    public void writeData() {
        Workbook wb = new HSSFWorkbook();
//Workbook wb = new XSSFWorkbook();
        CreationHelper createHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet("new sheet");
// Create a row and put some cells in it. Rows are 0 based.
        Row row = sheet.createRow(0);
// Create a cell and put a value in it.
        Cell cell = row.createCell(0);
        cell.setCellValue(1);
// Or do it on one line.
        row.createCell(1).setCellValue(1.2);
        row.createCell(2).setCellValue(
                createHelper.createRichTextString("This is a string"));
        row.createCell(3).setCellValue(true);
// Write the output to a file

        cell = row.createCell(4);
        cell.setCellValue("http://mless.com");
        Hyperlink link = (Hyperlink)createHelper.createHyperlink(HyperlinkType.URL);
        link.setAddress("http://www.tutorialspoint.com/");
        cell.setHyperlink(link);
        try (OutputStream fileOut = new FileOutputStream("workbook.xls")) {
            wb.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void printData() {

    }
}
