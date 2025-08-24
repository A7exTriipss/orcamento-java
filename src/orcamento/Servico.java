package orcamento;

public class Servico {

	private String descricao;
	private double quantidade;
	private double quantidadePorM2;
	private double preco;
	
	
	
	public Servico(String descricao) {
		this.descricao = descricao;
	}
	public String getDescricao() {
		return descricao;
	}
	public void setDescricao(String descricao) {
		this.descricao = descricao;
	}
	public double getQuantidade() {
		return quantidade;
	}
	public void setQuantidade(double quantidade) {
		this.quantidade = quantidade;
	}
	
	public double getQuantidadePorM2() {
		return quantidadePorM2;
	}
	public void setQuantidadePorM2(double quantidadePorM2) {
		this.quantidadePorM2 = quantidadePorM2;
	}
	public double getPreco() {
		return preco;
	}
	public void setPreco(double preco) {
		this.preco = preco;
	}
	
	
	
}
