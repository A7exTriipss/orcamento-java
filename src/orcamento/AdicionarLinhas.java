 	package orcamento;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class AdicionarLinhas{
	
	private final Sheet sheet;
	
	public AdicionarLinhas(Sheet sheet) {
		this.sheet = sheet;
	}
	
	/** 
	 * Copia a linha de referencia para uma posição especifica, empurrando linhas abaixo se necessá
	 * @param linhaReferencia índice da linha de referencia
	 * @param destino índice da linha de destino
	 * @return a nova Row criada
	 */
	
	public Row copiarLinhaParaPosicao(int linhaReferencia, int destino) {
		
		Row linhaReferenciaRow = sheet.getRow(linhaReferencia);
		Row novaLinha = sheet.createRow(destino);
		
		Workbook workbook = sheet.getWorkbook();
		
		//Copiar altura
		
		UtilExcel.copiarAltura(linhaReferenciaRow, novaLinha);
		
		//Copiar células
		
		for(int i = 0; i < linhaReferenciaRow.getLastCellNum(); i++) {
			
			Cell sourceCell = linhaReferenciaRow.getCell(i);
			Cell targetCell = novaLinha.createCell(i);
			
			if(sourceCell != null) {
				UtilExcel.copiarEstilo(workbook, sourceCell, targetCell);
				UtilExcel.copiarValorOuFormula(sourceCell, targetCell);
			}
			
		}
		
		//Copiar mesclagens
		
		UtilExcel.copiarMesclagens(sheet, linhaReferencia, destino);
		
		return novaLinha;
		
	}
}
