package com.example.excel;

import com.example.excel.model.Client;
import com.example.excel.relatoryXLSX.CreateExcel;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class TestCreateExcel {

    private final List<Client> clients = new ArrayList<>();

    @BeforeEach
    public void setup() {
        clients.add(new Client(1, "Abner Amos de Souza", "abner@email.com", "11973851774"));
        clients.add(new Client(2, "Juliana Menezes Silva", "juliana@email.com", "11932026167"));
        clients.add(new Client(3, "Anna Beatriz de Jesus", "anna@email.com", "11981756810"));
    }

    @Test
    void createFile() throws IOException {
        var createFileExcel = new CreateExcel();
        createFileExcel.createFile("clients.xlsx", clients);
    }
}
