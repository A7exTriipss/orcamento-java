package orcamento;

public class Main {
	
	
	public static void main(String[] args) {
		
		Cliente alex = new Cliente("Alex Tavares", "a7ex98@gmail.com", "+32466238958");
		
		Servico servico1 = new Servico("Instalar componentes eletronicos");
		servico1.setQuantidade(10);
		servico1.setPreco(200);
		
		Servico servico2 = new Servico("Instalar programas");
		servico2.setQuantidade(4);
		servico2.setPreco(400);
		
		Servico servico3 = new Servico("Limpeza do sistema");
		servico3.setQuantidade(1);
		servico3.setPreco(100);
		Servico servico4 = new Servico("Limpeza do sistema");
		servico4.setQuantidade(1);
		servico4.setPreco(300);
		Servico servico5 = new Servico("Limpeza do sistema");
		servico5.setQuantidade(1);
		servico5.setPreco(100);
		Servico servico6 = new Servico("Limpeza do sistema");
		servico6.setQuantidade(1);
		servico6.setPreco(100);
		Servico servico7 = new Servico("Limpeza do sistema");
		servico7.setQuantidade(1);
		servico7.setPreco(100);
		
		Orcamento orcamento1 = new Orcamento(alex, "Orçamento Alex");
		orcamento1.adicionarServico(servico1);
		orcamento1.adicionarServico(servico2);
		orcamento1.adicionarServico(servico3);
		orcamento1.adicionarServico(servico4);
		orcamento1.adicionarServico(servico5);
		orcamento1.adicionarServico(servico6); 
		orcamento1.adicionarServico(servico7);
		
		//orcamento1.setValorManual(850);
		
		System.out.println(orcamento1.getTotalFinal());
		System.out.println(orcamento1.getNomeDoOrcamento());
		
		//Imprimindo Serviços
		System.out.println("Orçamento para " + orcamento1.getCliente().getNome());
		System.out.println("Serviço: ");
		for (Servico s : orcamento1.getServicos()) {
			System.out.println("- " + s.getDescricao() + " (" + s.getQuantidade() + " x R$" + s.getPreco() + ")");
		}
		
		//Exportar para txt
		
		ExportadorTxt.exportarOrcamento(orcamento1, orcamento1.getNomeDoOrcamento() + ".txt");
		ExportadorExcel.exportarOrcamento(orcamento1, orcamento1.getNomeDoOrcamento() + ".xlsx");
		
		//Preenchedor do Excel Modelo
		
		
		EscritorExcelModelo escritor = new EscritorExcelModelo();
		escritor.preencher(orcamento1, "orcamento.xlsx", "orcamento" + orcamento1.getCliente().getNome() + ".xlsx");
		
	}

}
