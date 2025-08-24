package orcamento;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExportadorExcel {
	
	public static void exportarOrcamento(Orcamento orcamento, String caminhoArquivo) {
		
		Workbook workbook = new XSSFWorkbook();
		
		Sheet sheet = workbook.createSheet("Orçamento");

        int rowNum = 0;

        // Dados do cliente
        sheet.createRow(rowNum++).createCell(0).setCellValue("Orçamento");
        sheet.createRow(rowNum++).createCell(0).setCellValue("Cliente: " + orcamento.getCliente().getNome());
        sheet.createRow(rowNum++).createCell(0).setCellValue("Email: " + orcamento.getCliente().getEmail());
        sheet.createRow(rowNum++).createCell(0).setCellValue("Telefone: " + orcamento.getCliente().getTelefone());

        rowNum++; // Linha em branco

        // Cabeçalho da tabela
        Row cabecalho = sheet.createRow(rowNum++);
        cabecalho.createCell(0).setCellValue("Descrição");
        cabecalho.createCell(1).setCellValue("Quantidade");
        cabecalho.createCell(2).setCellValue("Valor Unitário");
        cabecalho.createCell(3).setCellValue("Subtotal");

        // Conteúdo dos serviços
        for (Servico s : orcamento.getServicos()) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(s.getDescricao());
            row.createCell(1).setCellValue(s.getQuantidade());
            row.createCell(2).setCellValue(s.getPreco());
            row.createCell(3).setCellValue(s.getQuantidade() * s.getPreco());
        }

        rowNum++; // Espaço antes do total

        // Total
        Row totalRow = sheet.createRow(rowNum++);
        totalRow.createCell(2).setCellValue("TOTAL:");
        totalRow.createCell(3).setCellValue(orcamento.getTotalFinal());

        // Ajuste automático de largura
        for (int i = 0; i <= 3; i++) {
            sheet.autoSizeColumn(i);
        }

        // Salvar o arquivo
        try (FileOutputStream out = new FileOutputStream(caminhoArquivo)) {
            workbook.write(out);
            workbook.close();
            System.out.println("Excel gerado com sucesso: " + caminhoArquivo);
        } catch (IOException e) {
            System.out.println("Erro ao gerar Excel: " + e.getMessage());
        }
		
	}
	

}
