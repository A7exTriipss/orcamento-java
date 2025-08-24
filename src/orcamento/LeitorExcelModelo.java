package orcamento;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class LeitorExcelModelo {
    public static void main(String[] args) {
        String caminhoArquivo = "orçamento.xlsx"; // Renomeie para o nome do seu arquivo

        try (FileInputStream fis = new FileInputStream(caminhoArquivo);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // primeira aba
            System.out.println("Lendo conteúdo da planilha...");

            for (int rowIndex = 0; rowIndex <= 40; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row != null) {
                    System.out.print("Linha " + (rowIndex + 1) + ": ");
                    for (int colIndex = 0; colIndex <= 10; colIndex++) {
                        Cell cell = row.getCell(colIndex);
                        if (cell != null) {
                            System.out.print(getValorCelula(cell) + " | ");
                        } else {
                            System.out.print("vazio | ");
                        }
                    }
                    System.out.println();
                }
            }

        } catch (IOException e) {
            System.out.println("Erro ao ler Excel: " + e.getMessage());
        }
    }

    private static String getValorCelula(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return " ";
        }
    }
}