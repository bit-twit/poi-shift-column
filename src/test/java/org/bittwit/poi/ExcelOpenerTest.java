package org.bittwit.poi;

import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Assert;
import org.junit.Test;

public class ExcelOpenerTest {

	@Test
	public void testGetNumberOfRows () throws InvalidFormatException, IOException {
		System.out.println("testGetNumberOfRows()");

		ExcelOpener op = new ExcelOpener("test.xlsx", "out.xlsx");
		op.open();

		int nrRowsSheet0 = op.getNumberOfRows(0);
		Assert.assertTrue(nrRowsSheet0 > 0);

		op.close();
	}

	@Test
	public void testInsertNewColumn () throws InvalidFormatException, IOException {
		System.out.println("testInsertNewColumn()");
		int sheetIndex = 0;
		int columnIndex = 1;

		ExcelOpener op = new ExcelOpener("test.xlsx", "out.xlsx");
		op.open();

		int nrColumnsBefore = op.getNrColumns(sheetIndex);
		op.insertNewColumnBefore(sheetIndex, columnIndex);
		int nrColumnsAfter = op.getNrColumns(sheetIndex);

		Assert.assertTrue(nrColumnsAfter == (nrColumnsBefore  + 1));

		op.close();
	}

	@Test
	public void testGetNrColumn () throws InvalidFormatException, IOException {
		System.out.println("testGetNrColumn()");

		ExcelOpener op = new ExcelOpener("test.xlsx", "out.xlsx");
		op.open();

		int sheetIndex = 0;
		int nrColumns = op.getNrColumns(sheetIndex);

		Assert.assertTrue(nrColumns == 2);

		op.close();
	}
}
