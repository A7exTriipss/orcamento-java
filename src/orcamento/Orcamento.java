package orcamento;

import java.util.ArrayList;
import java.util.List;

public class Orcamento {
	
	private String nomeDoOrcamento;
	private Cliente cliente;
	private List<Servico> servicos;
	private double valorManual;
	
	
	
	public Orcamento(Cliente cliente, String nomeDoOrcamento) {
		this.cliente = cliente;
		this.nomeDoOrcamento = nomeDoOrcamento;
		servicos = new ArrayList<Servico>();
	}
	
	public Cliente getCliente() {
		return cliente;
	}
	
	public void setCliente(Cliente cliente) {
		this.cliente = cliente;
	}
	
	public List<Servico> getServicos() {
		return servicos;
	}
	
	public void adicionarServico(Servico servico) {
		servicos.add(servico);
	}
	
	public double getTotalCalculado() {
		return servicos.stream().mapToDouble(Servico::getPreco).sum();
	}
	
	public double getValorManual() {
		return valorManual;
	}
	
	public void setValorManual(double valorManual) {
		this.valorManual = valorManual;
	}
	
	public double getTotalFinal() {
		return valorManual > 0 ? valorManual : getTotalCalculado();
	}

	public String getNomeDoOrcamento() {
		return nomeDoOrcamento;
	}

	public void setNomeDoOrcamento(String nomeDoOrcamento) {
		this.nomeDoOrcamento = nomeDoOrcamento;
	}
	
	
	
	
	
	
	

}
