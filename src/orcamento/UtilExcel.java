package orcamento;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class UtilExcel {
	
	/*Copiar mesclagens da linha de origem para a linha de destino */
	public static void copiarMesclagens(Sheet sheet, int sourceRowNum, int targetRowNum) {
		
		for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress merged = sheet.getMergedRegion(i);
            if (merged.getFirstRow() == sourceRowNum) {
                CellRangeAddress newMerged = new CellRangeAddress(
                        targetRowNum,
                        targetRowNum + (merged.getLastRow() - merged.getFirstRow()),
                        merged.getFirstColumn(),
                        merged.getLastColumn()
                );

                boolean conflito = false;
                for (int j = 0; j < sheet.getNumMergedRegions(); j++) {
                    if (sheet.getMergedRegion(j).intersects(newMerged)) {
                        conflito = true;
                        break;
                    }
                }

                if (!conflito) sheet.addMergedRegion(newMerged);
            }
		
		}
		
		/* Copiar valor ou fórmula da célula de origem para a célula de destino */
		
		
		
	}
	
	public static void copiarValorOuFormula (Cell sourceCell, Cell targetCell) {
		
		if (sourceCell == null || targetCell == null) {
			return;
		}
		
		switch (sourceCell.getCellType()) {
		case STRING:
			targetCell.setCellValue(sourceCell.getStringCellValue());
			break;
		case NUMERIC:
			targetCell.setCellValue(sourceCell.getNumericCellValue());
			break;
		case BOOLEAN:
			targetCell.setCellValue(sourceCell.getBooleanCellValue());
			break;
		case FORMULA:
			targetCell.setCellValue(sourceCell.getCellFormula());
			break;
		default:
			break;
		
		
		
		}
	
	
	}
	
	/* Copiar estilo da célula e origem para a célula de destino usando o Workbook fornecido */
	public static void copiarEstilo(Workbook workbook, Cell sourceCell, Cell targetCell) {
		
		if (sourceCell == null || targetCell == null) 
			return;
		
		CellStyle newStyle = workbook.createCellStyle();
		
		newStyle.cloneStyleFrom(sourceCell.getCellStyle());
		targetCell.setCellStyle(newStyle);
		
	}
	
	/* Copiar a altura da linha de origem para a linha de destino */
	
	public static void copiarAltura(Row sourceRow, Row targetRow) {
		
		if (sourceRow == null || targetRow == null) 
			return;
		
		targetRow.setHeight(sourceRow.getHeight());
	}
	
	/* Copiar uma linha inteira (células + estilos) */
	public static void copiarLinha(Sheet sheet, Row sourceRow, Row targetRow) {
		Workbook workbook = sheet.getWorkbook();
		if(sourceRow == null || targetRow == null) {
			return;
		}
		
		for(int i = sourceRow.getFirstCellNum(); i < sourceRow.getLastCellNum(); i++) {
			Cell sourceCell = sourceRow.getCell(i);
			Cell targetCell = targetRow.createCell(i);
			
			if(sourceCell != null) {
				/* Copiar estilo*/
				copiarEstilo(workbook, sourceCell, targetCell);
				
				/*Copiar valor ou formula*/
				copiarValorOuFormula(sourceCell, targetCell);
			}
			
			
		}
		
		/* Copiar mesclagem*/
		
		copiarMesclagens(sheet, sourceRow.getRowNum(), targetRow.getRowNum());
		
	}
	/* Copiar varias linhas a partir de uma linha de origem*/
	
	public static void copiarLinhasComEstilo(Sheet sheet, Row sourceRow, int quantidade) {
		if(sourceRow == null || quantidade <= 0) {
			return;
		}
		
		int comeco = sourceRow.getRowNum() + 1;
		for(int i = 0; i < quantidade; i++) {
			Row novaLinha = sheet.createRow(comeco + i);
			copiarLinha(sheet, sourceRow, novaLinha);
			
		}
		
	
	}
	
	public static void copiarLinhasComEstiloNoMeio(Sheet sheet, Row sourceRow, int quantidade, Row inicioBlocoParaEmpurrar) {
		if(quantidade <= 0 || sourceRow == null) {
			return;
		}
		int comecoReferencia = sourceRow.getRowNum();
		int comecoBloco = inicioBlocoParaEmpurrar.getRowNum();
		/*Empurrar o bloco de linhas para baixo*/
		
		EmpurrarLinhasParaBaixo.empurraLinhasParaBaixo(sheet, comecoReferencia, quantidade, comecoBloco, sheet.getLastRowNum());
		
		
		/*copia as linhas com estilo e valor*/
		for(int i = 0; i < quantidade; i++) {
			Row novaLinha = sheet.createRow(comecoReferencia + i);
			copiarLinha(sheet,sourceRow, novaLinha);
		}
	}
	
	
	
}
