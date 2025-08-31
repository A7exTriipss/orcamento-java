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
	
	
	
	
	
}
