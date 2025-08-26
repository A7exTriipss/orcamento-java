package orcamento;

import java.io.*;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class EscritorExcelModelo {
	
	public void preencher(Orcamento orcamento, String caminhoModelo, String caminhoSaida) {
		try(FileInputStream fis = new FileInputStream(caminhoModelo);
				Workbook workbook = new XSSFWorkbook(fis)){
			Sheet sheet = workbook.getSheetAt(0);
			
			// Preencher Area do cliente
			
			sheet.getRow(15).getCell(1).setCellValue(orcamento.getCliente().getNome());
			sheet.getRow(16).getCell(1).setCellValue(orcamento.getCliente().getTelefone());
			sheet.getRow(16).getCell(5).setCellValue(orcamento.getCliente().getEmail());
			
			
			
			
			//Preencher Serviços
			
			int linhaInicialServico = 22;
			int maximoServicosModelo = 2;
			int linhaFinalModelo = linhaInicialServico + maximoServicosModelo -1;
			List<Servico> servicos = orcamento.getServicos();
			
			// Antes de preencher, empurra a linha do total
			EmpurrarLinhasParaBaixo.empurraLinhasParaBaixo(sheet, linhaInicialServico, servicos.size(), 24, 42);
			
			
			
			
			
			int LinhaAtual = linhaInicialServico;
			for(int i = 0 ; i < servicos.size() ; i++) {
				
				if (LinhaAtual > linhaFinalModelo) {
					//criar nova linha copiando a última linha do modelo
					AdicionarLinhas.copiarUltimaLinha(sheet, linhaFinalModelo, LinhaAtual);
					
				}
					
				
				
				Servico s = servicos.get(i);
				Row linha = sheet.getRow(linhaInicialServico + i);
				if (linha == null) {
					linha = sheet.createRow(linhaInicialServico +i);
					
				}
				
				linha.getCell(0).setCellValue(s.getDescricao());
				
				LinhaAtual++;
			}

			
			//Preencher Total
			
			sheet.getRow(22 + servicos.size()).getCell(8).setCellValue(orcamento.getTotalFinal());
			
			try(FileOutputStream fos = new FileOutputStream(caminhoSaida)){
				workbook.write(fos);
				System.out.println("Arquivo salvo em " + caminhoSaida);
			}
			

			
			
			
		} catch (IOException e) {
			System.out.println("Erro ao preencher Excel" + e.getMessage());
		}
	}
		
	
	
	
	

}
