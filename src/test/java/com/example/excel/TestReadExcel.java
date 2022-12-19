package com.example.excel;

import com.example.excel.relatoryXLSX.ReadExcel;
import org.junit.jupiter.api.Test;

public class TestReadExcel {

    @Test
    void ReadExcel() {
        var readFileExcel = new ReadExcel();
        var clients = readFileExcel.readFile("clients.xlsx");
    }
}
