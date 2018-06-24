import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.util.ArrayList;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellAddress;


public class Main {
	
	private static ArrayList<Integer> rowNumsOfValidTransferCells = new ArrayList<>();

	public static void main(String[] args) throws InvalidFormatException, IOException {


		File analysisSpreadsheet = new File("analysisSpreadsheet.xlsx");
		File exportedSpreadsheet = new File("exportedexcelquickbooks.xlsx");

		FileInputStream inputAnalysis = new FileInputStream(analysisSpreadsheet);
		FileInputStream inputExported = new FileInputStream(exportedSpreadsheet);
		Workbook analysisWB = WorkbookFactory.create(inputAnalysis);
		Workbook exportedWB = WorkbookFactory.create(inputExported);
		Sheet analysisSheet = analysisWB.getSheetAt(0);
		Sheet exportedSheet = exportedWB.getSheetAt(0);


		int rowStartAnalysis = 7;
		int rowStartExported = 7;
		int rowEndAnalysis = 8;
		int rowEndExported = 0;

		for (int rowNum = rowStartAnalysis; rowNum <= analysisSheet.getLastRowNum(); rowNum++) {
			Row r = analysisSheet.getRow(rowNum);
			if(isRowEmpty(r)) {
				break;
			} else {
				rowEndAnalysis++;
			}
		}

		for(int rowNum=rowStartExported; rowNum < exportedSheet.getLastRowNum(); rowNum++) {
			try {
				Row r = exportedSheet.getRow(rowNum);
				Cell c = r.getCell(9);

				if(c.toString().isEmpty()) {break;} 
				else {rowEndExported++;}

			} catch (NullPointerException e) {
				break;
			}
		} rowEndExported += 7;



		int exportedNumOfRows = numOfRowsToInsertFromExport(analysisSheet, exportedSheet, rowEndAnalysis, rowEndExported);

		analysisSheet.shiftRows(rowEndAnalysis, analysisSheet.getLastRowNum(), exportedNumOfRows);
		for(int i=0; i < exportedNumOfRows + 2; i++) {
			analysisSheet.createRow(329 + i);
		}
		
//		for(int i=0; i<rowNumsOfValidTransferCells.length; i++) {
//			System.out.println(rowNumsOfValidTransferCells[i]);
//		}
		
		
		
		
		
		inputAnalysis.close();
		inputExported.close();

		analysisWB.write(new FileOutputStream(analysisSpreadsheet));
		analysisWB.close();
		exportedWB.write(new FileOutputStream(exportedSpreadsheet));
		exportedWB.close();


		System.out.println(rowEndExported + " :: exported rows");
		System.out.println(rowEndAnalysis + " :: analysis rows");
		System.out.println(exportedNumOfRows + " :: added rows");







	}
	
	public static void transferCells(Sheet analysis, Sheet transfer, ArrayList<Integer> rowsToTransfer) {
		int exportDateCol = 1;
		int exportTransactionTypeCol = 2;
		int exportCheckNumCol = 3;
		int exportNameCol = 4;
		int exportMemoCol = 5;
		int exportAmountCol = 8;
		for (int i=0; i < rowsToTransfer.size(); i++ ) {
			Cell dateCellsFromExport = transfer.getRow(rowsToTransfer.get(i)).getCell(exportDateCol);
			
		}
		
		
	}

	/**
	 * 
	 * @param row
	 * @return
	 */
	public static boolean isRowEmpty(Row row) {
		if (row == null) {
			return true;
		}
		if (row.getLastCellNum() <= 0) {
			return true;
		}
		for(int cellNum = row.getFirstCellNum(); cellNum < row.getLastCellNum(); cellNum++) {
			Cell cell = row.getCell(cellNum);
			if(cell != null && cell.getCellTypeEnum() != CellType.BLANK && !cell.toString().isEmpty()) {
				return false;
			}
		}return true;


	}
		
	
	/**
	 * 
	 * @param analysis
	 * @param export
	 * @param rowEndAnalysis
	 * @param rowEndExported
	 * @return
	 */
	public static int numOfRowsToInsertFromExport(Sheet analysis, Sheet export, int rowEndAnalysis, int rowEndExported) {
		//Need to figure out how many rows to insert into the analysis from the exported sheet
		int rowCount =0;
		for(int rowNumExport=7; rowNumExport < rowEndExported; rowNumExport++) {
			for(int rowNumAnalysis= rowEndAnalysis-2; rowNumAnalysis > 300; rowNumAnalysis--) {
				try {
					Cell currentExportCell = export.getRow(rowNumExport).getCell(9);
					Cell currentAnalysisCell = analysis.getRow(rowNumAnalysis).getCell(8);
													
					String exportedValue = String.format("%.2f", currentExportCell.getNumericCellValue()); 
					String analysisValue = String.format("%.2f", currentAnalysisCell.getNumericCellValue());

					if(!(exportedValue.equals(analysisValue))) {
						System.out.println("Export row num: " + rowNumExport);
						System.out.println("Analysis row num: " + rowNumAnalysis);
						
						rowNumsOfValidTransferCells.add(rowNumExport);

						rowCount++;
						break;
					} else {
						break;
					}
				} catch (Exception exception) {
					System.out.println(rowNumExport + " ::: " + rowNumAnalysis);
					exception.printStackTrace();
				}

			}
		}return rowCount;
	}
}
