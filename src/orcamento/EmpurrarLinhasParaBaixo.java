package orcamento;

import org.apache.poi.ss.usermodel.Sheet;

public class EmpurrarLinhasParaBaixo {

	
	public static void empurraLinhasParaBaixo(Sheet sheet, int linhaInicio, int qtdServicos, int linhaTotal, int linhaFimBloco) {
	    // Quantas linhas já existem disponíveis entre o início da tabela e o total
	    int espacoDisponivel = linhaTotal - linhaInicio;

	    if (qtdServicos > espacoDisponivel) {
	        int linhasExtras = qtdServicos - espacoDisponivel;

	        // Empurra todo o bloco (TOTAL + OBS + ASSINATURA)
	        sheet.shiftRows(linhaTotal, linhaFimBloco, linhasExtras, true, false);
	    }
	}
			
		
		
	
}
