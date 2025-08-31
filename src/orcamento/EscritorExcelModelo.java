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
			int linhaFinalServico = linhaInicialServico + maximoServicosModelo -1;
			int linhaTotal = 24;
			int linhaFimBloco = 42;
			List<Servico> servicos = orcamento.getServicos();
			
			// Antes de preencher, empurra a linha do total
			EmpurrarLinhasParaBaixo.empurraLinhasParaBaixo(sheet, linhaInicialServico, servicos.size(), linhaTotal, linhaFimBloco);
			
			
			
			//Criar/adicionar linhas de serviço extras usando AdicionarLinhas
			
			AdicionarLinhas adicionador = new AdicionarLinhas(sheet);
			int linhaAtual = linhaInicialServico;
			for(int i = 0 ; i < servicos.size() ; i++) {
				
				//Se linhaAtual estiver além do modelo, criar nova linha copiando a última do modelo
				if (linhaAtual > linhaFinalServico) {
					adicionador.copiarLinhaParaPosicao(linhaFinalServico, linhaAtual);
					
				}
					
				
				//Preencher os dados do serviço
				Servico s = servicos.get(i);
				Row linha = sheet.getRow(linhaAtual);
				if (linha == null) {
					linha = sheet.createRow(linhaAtual);
					
				}
				
				linha.getCell(0).setCellValue(s.getDescricao());
				
				// Preencher QTD VALOR...
				
				linhaAtual++;
			}

			
			//Preencher Total
			
			Row totalRow = sheet.getRow(linhaInicialServico + servicos.size());
			
			if(totalRow == null) {
				totalRow = sheet.createRow(linhaInicialServico + servicos.size());	
			}
			totalRow.getCell(8).setCellValue(orcamento.getTotalFinal());
			
			
			try(FileOutputStream fos = new FileOutputStream(caminhoSaida)){
				workbook.write(fos);
				System.out.println("Arquivo salvo em " + caminhoSaida);
			}
			

			
			
			
		} catch (IOException e) {
			System.out.println("Erro ao preencher Excel" + e.getMessage());
		}
	}
		
	
	
	
	

}
