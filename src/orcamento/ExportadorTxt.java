package orcamento;

import java.io.FileWriter;
import java.io.IOException;

public class ExportadorTxt {
	
	public static void exportarOrcamento(Orcamento orcamento, String caminhoArquivo) {
		try {
			
			FileWriter writer = new FileWriter(caminhoArquivo);
			
			writer.write("Orçamento\n");
			writer.write("Cliente: " + orcamento.getCliente().getNome() + "\n");
			writer.write("Telefone: " + orcamento.getCliente().getTelefone() + "\n");
			writer.write("Email: " + orcamento.getCliente().getEmail() + "\n");
			
			writer.write("\nServiços: ");
			for (Servico s: orcamento.getServicos()) {
				writer.write("- " + s.getDescricao() + " (" + s.getQuantidade() + " x R$" + s.getPreco() + ")\n");
			}
			
			writer.write("Valor total final: " + orcamento.getTotalFinal() + "\n");
			
			writer.close();
			
			System.out.println("Orçamento exportado com sucesso para: " + caminhoArquivo);
			
			
		} catch(IOException e) {
			System.out.println("Erro ao exportar orçamento: " + e.getMessage());
		}
	}

}
