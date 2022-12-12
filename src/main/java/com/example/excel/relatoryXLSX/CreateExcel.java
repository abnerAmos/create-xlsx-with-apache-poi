package com.example.excel.relatoryXLSX;

import com.example.excel.model.Client;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

@Slf4j
public class CreateExcel {

    // Método que recebe o Nome do Arquivo e a Lista de clientes que serão inseridos no Excel
    public void createFile(final String fileName, final List<Client> client) throws IOException { // Excessao para FileOutpuStream e XSSDWorkbook
        log.info("Gerando o arquivo {}" + fileName);

        try (
                XSSFWorkbook workbook = new XSSFWorkbook(); // Cria o Arquivo do Excel
                FileOutputStream outputStream = new FileOutputStream(fileName)) { // Faz a saída do Arquivo Excel
            XSSFSheet spreadsheet = workbook.createSheet("Lista de Clientes"); // Gera a Planilha e nomeia (spreadsheet = planilha)
            int countNumberLine = 0; // Contador de linhas

            addHeader(spreadsheet, countNumberLine++); // recebe a Planilha e Numero de Linhas

            for (Client e : client) {   // For para iteração dos dados do cliente - Uma linha para cada cliente
                XSSFRow linha = spreadsheet.createRow(countNumberLine++);   //
                addCell(linha, 0, e.getId());
                addCell(linha, 1, e.getName());
                addCell(linha, 2, e.getEmail());
                addCell(linha, 3, e.getContact());
        }
            workbook.write(outputStream);   // Faz a gravação dos dados na Tabela
        } catch (FileNotFoundException e) {
            log.error("Arquivo não encontrado: {}", fileName);  // Erro caso não encontre o arquivo
        } catch (IOException e) {
            log.error("Erro ao processar o arquivo: {} ", fileName);    // Erro caso não processe o arquivo
        }
        log.info("Arquivo gerado com sucesso!");
    }

        private void addHeader(XSSFSheet spreadsheet, int countNumberLine) { // recebe a Planilha e Numero de Linhas
            XSSFRow line = spreadsheet.createRow(countNumberLine); // Cria Colunas na Tabela
            addCell(line, 0, "Id");
            addCell(line, 1, "Name");
            addCell(line, 2, "Email");
            addCell(line, 3, "Contact");  // Adiciona os "Titulos" em cada coluna
        }

        private void addCell(Row line, int column, String valor) { // Métodos que adiciona campos com valor String
            Cell cell = line.createCell(column);    // Informa qual coluna esta posicionado
            cell.setCellValue(valor);               // Informa o valor da coluna/linha
    }

        private void addCell(Row line, int column, int valor) { // Métodos que adiciona campo com valor String
            Cell cell = line.createCell(column);    // Informa qual coluna esta posicionado
            cell.setCellValue(valor);               // Informa o valor da coluna/linha
        }
}
