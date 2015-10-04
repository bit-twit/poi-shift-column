package org.bittwit.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelOpener {
	private String sourceFile;
	private String outputFile;
	private XSSFWorkbook workbook;

	public ExcelOpener(String inputFile, String outputFile) {
		this.sourceFile = inputFile;
		this.outputFile = outputFile;
	}

	public void open() throws InvalidFormatException, IOException {
		workbook = new XSSFWorkbook(new File(this.sourceFile));
		writeToFile(workbook, this.outputFile);
		workbook.close();

		// open output file to edit
		workbook = new XSSFWorkbook(new FileInputStream(this.outputFile));
	}

	public int getNumberOfRows(int sheetIndex) {
		assert workbook != null;

		int sheetNumber = workbook.getNumberOfSheets();

		System.out.println("Found " + sheetNumber + " sheets.");

		if (sheetIndex >= sheetNumber) {
			throw new RuntimeException("Sheet index " + sheetIndex
					+ " invalid, we have " + sheetNumber + " sheets");
		}

		Sheet sheet = workbook.getSheetAt(sheetIndex);

		int rowNum = sheet.getLastRowNum() + 1;

		System.out.println("Found " + rowNum + " rows.");

		return rowNum;
	}

	public void insertNewColumnBefore(int sheetIndex, int columnIndex) {
		assert workbook != null;

		FormulaEvaluator evaluator = workbook.getCreationHelper()
				.createFormulaEvaluator();
		evaluator.clearAllCachedResultValues();

		Sheet sheet = workbook.getSheetAt(sheetIndex);
		int nrRows = getNumberOfRows(sheetIndex);
		int nrCols = getNrColumns(sheetIndex);
		System.out.println("Inserting new column at " + columnIndex);

		for (int row = 0; row < nrRows; row++) {
			Row r = sheet.getRow(row);

			if (r == null) {
				continue;
			}

			// shift to right
			for (int col = nrCols; col > columnIndex; col--) {
				Cell rightCell = r.getCell(col);
				if (rightCell != null) {
					r.removeCell(rightCell);
				}

				Cell leftCell = r.getCell(col - 1);

				if (leftCell != null) {
					Cell newCell = r.createCell(col, leftCell.getCellType());
					cloneCell(newCell, leftCell);
					if (newCell.getCellType() == Cell.CELL_TYPE_FORMULA) {
						evaluator.notifySetFormula(newCell);
						CellValue cellValue = evaluator.evaluate(newCell);
						evaluator.evaluateFormulaCell(newCell);
						System.out.println(cellValue);
					}
				}
			}

			// delete old column
			int cellType = Cell.CELL_TYPE_BLANK;

			Cell currentEmptyWeekCell = r.getCell(columnIndex);
			if (currentEmptyWeekCell != null) {
//				cellType = currentEmptyWeekCell.getCellType();
				r.removeCell(currentEmptyWeekCell);
			}

			// create new column
			r.createCell(columnIndex, cellType);
		}

		// Adjust the column widths
		for (int col = nrCols; col > columnIndex; col--) {
			sheet.setColumnWidth(col, sheet.getColumnWidth(col - 1));
		}

/*
		Row Specialrow = sheet.getRow(46);
		String formula = "SUM(AP16:AP46)";
		Cell cellFormula = Specialrow.createCell(nrCols - 1);
		cellFormula.setCellType(XSSFCell.CELL_TYPE_FORMULA);
		cellFormula.setCellFormula(formula);
*/
		XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
	}

	/*
	 * Takes an existing Cell and merges all the styles and forumla into the new
	 * one
	 */
	private static void cloneCell(Cell cNew, Cell cOld) {
		cNew.setCellComment(cOld.getCellComment());
		cNew.setCellStyle(cOld.getCellStyle());

		switch (cOld.getCellType()) {
		case Cell.CELL_TYPE_BOOLEAN: {
			cNew.setCellValue(cOld.getBooleanCellValue());
			break;
		}
		case Cell.CELL_TYPE_NUMERIC: {
			cNew.setCellValue(cOld.getNumericCellValue());
			break;
		}
		case Cell.CELL_TYPE_STRING: {
			cNew.setCellValue(cOld.getStringCellValue());
			break;
		}
		case Cell.CELL_TYPE_ERROR: {
			cNew.setCellValue(cOld.getErrorCellValue());
			break;
		}
		case Cell.CELL_TYPE_FORMULA: {
			cNew.setCellFormula(cOld.getCellFormula());
			break;
		}
		}
	}

	public int getNrColumns(int sheetIndex) {
		assert workbook != null;

		Sheet sheet = workbook.getSheetAt(sheetIndex);

		// get header row
		Row headerRow = sheet.getRow(0);
		int nrCol = headerRow.getLastCellNum();

		// while
		// (!headerRow.getCell(nrCol++).getStringCellValue().equals(LAST_COLUMN_HEADER));

		// while (nrCol <= headerRow.getPhysicalNumberOfCells()) {
		// Cell c = headerRow.getCell(nrCol);
		// nrCol++;
		//
		// if (c!= null && c.getCellType() == Cell.CELL_TYPE_STRING) {
		// if (c.getStringCellValue().equals(LAST_COLUMN_HEADER)) {
		// break;
		// }
		// }
		// }

		System.out.println("Found " + nrCol + " columns.");
		return nrCol;

	}

	private static void writeToFile(Workbook workbook, String fileName)
			throws IOException {
		FileOutputStream fileOut = new FileOutputStream(fileName);
		workbook.write(fileOut);
		fileOut.close();
	}

	public void close() {
		assert workbook != null;

		try {
			writeToFile(workbook, this.outputFile);
			workbook.close();
		}
		catch (IOException e) {
			e.printStackTrace();
		}
	}

}
