 	package orcamento;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

public class AdicionarLinhas{
	
	public static void copiarUltimaLinha(Sheet sheet, int sourceRowNum, int destinationRowNum) {
		
		Row sourceRow = sheet.getRow(sourceRowNum);
		Row newRow = sheet.getRow(destinationRowNum);
		
		
		// Copiar a altura 
		if (newRow == null) {
		    newRow = sheet.createRow(destinationRowNum);
		}
		
		newRow.setHeight(sourceRow.getHeight());
		
		
		// Copiar Celulas
		
		
		for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
			
			Cell oldCell = sourceRow.getCell(i);
			Cell newCell = newRow.createCell(i);
			
			
			// Copiar estilo
			
			newCell.setCellStyle(oldCell.getCellStyle());
			
			
			// Copiar valor ou formula
			
			
			switch (oldCell.getCellType()) {
			case STRING:
				newCell.setCellValue(oldCell.getStringCellValue());
				break;
			case NUMERIC:
				newCell.setCellValue(oldCell.getNumericCellValue());
				break;
			case BOOLEAN:
				newCell.setCellValue(oldCell.getBooleanCellValue());
				break;
			case FORMULA:
				newCell.setCellValue(oldCell.getCellFormula());
				break;
			default:
				break;
			
			
			
			}
			
		}
		
		// Copiar mesclagens, se houver
		
		for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress merged = sheet.getMergedRegion(i);
            if (merged.getFirstRow() == sourceRowNum) {
                CellRangeAddress newMerged = new CellRangeAddress(
                        destinationRowNum,
                        destinationRowNum + (merged.getLastRow() - merged.getFirstRow()),
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
		
	}

}
