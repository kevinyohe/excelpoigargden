package dev.kevinyohe.excelpoi.excelpoi.excel;

import org.junit.jupiter.api.Test;

import java.io.FileNotFoundException;

import static org.junit.jupiter.api.Assertions.*;

class WorksheetUtilTest {

    @Test
    void getNewWorkbook() throws Exception {
        WorksheetUtil wsu = new WorksheetUtil();
        wsu.getNewWorkbook("workbook.xls", "new one");

    }

    @Test
    void writeData() {
        WorksheetUtil wsu = new WorksheetUtil();
        wsu.writeData();
    }
}