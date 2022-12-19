package com.example.excel.relatoryXLSX;

import com.example.excel.model.Client;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Slf4j
public class ReadExcel {

    public List<Client> readFile(final String nameFile) {
        log.info("Lendo arquivo {}", nameFile);
        List<Client> clients = new ArrayList<>();

        // Recebe o arquivo
        try (FileInputStream excelFile = new FileInputStream(nameFile)) {
            var workbook = new XSSFWorkbook(excelFile);     // Faz a abertura do arquivo
            var firstTab = workbook.getSheetAt(0);    // Faz a leitura da primeira Aba

            int countLine = 0;
            for (Row line : firstTab) {
                if (++countLine == 1) continue;     // Pula a primeira linha ao qual contém o cabeçalho
                Client client = Client.builder()
                        .id((int) line.getCell(0).getNumericCellValue())
                        .name(line.getCell(1).getStringCellValue())
                        .email(line.getCell(2).getStringCellValue())
                        .contact(line.getCell(3).getStringCellValue())
                        .build();
                clients.add(client);
                log.info("Lendo Cliente {}", client);
            }

        } catch (FileNotFoundException e) {
            log.error("Arquivo nao encontrado {}", nameFile);
        } catch (IOException e) {
            log.error("Erro ao processar o arquivo {}", nameFile);
        }
        log.info("Total de clientes lidos {}", clients.size());
        return clients;
    }
}
