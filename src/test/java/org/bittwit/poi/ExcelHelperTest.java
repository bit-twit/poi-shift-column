package org.bittwit.poi;

import org.junit.Assert;

import org.junit.Test;

public class ExcelHelperTest {

    @Test
    public void testUpdateFormula () {
        String formula = "A1:A6";
        String result = ExcelHelper.updateFormula(formula, 0);
        System.out.println(result);
        Assert.assertTrue("B1:B6".equals(result));

        formula = "B1:B6";
        result = ExcelHelper.updateFormula(formula, 1);
        System.out.println(result);
        Assert.assertTrue("C1:C6".equals(result));

        formula = "AAA1:AAA6";
        result = ExcelHelper.updateFormula(formula, 702);
        System.out.println(result);
        Assert.assertTrue("AAB1:AAB6".equals(result));
    }

	@Test
	public void testGetReferenceForColumnIndex () {
        int value = 0;
		String referenceValue = ExcelHelper.getReferenceForColumnIndex(value);
        System.out.println(value + " : " + referenceValue);
		Assert.assertEquals("A", referenceValue);

        value = 1;
        referenceValue = ExcelHelper.getReferenceForColumnIndex(value);
        System.out.println(value + " : " + referenceValue);
        Assert.assertEquals("B", referenceValue);

        value = 23;
        referenceValue = ExcelHelper.getReferenceForColumnIndex(value);
        System.out.println(value + " : " + referenceValue);
		Assert.assertEquals("X", referenceValue);

        value = 24;
		referenceValue = ExcelHelper.getReferenceForColumnIndex(value);
        System.out.println(value + " : " + referenceValue);
		Assert.assertEquals("Y", referenceValue);

        value = 25;
		referenceValue = ExcelHelper.getReferenceForColumnIndex(value);
        System.out.println(value + " : " + referenceValue);
		Assert.assertEquals("Z", referenceValue);

        value = 26;
		referenceValue = ExcelHelper.getReferenceForColumnIndex(value);
        System.out.println(value + " : " + referenceValue);
		Assert.assertEquals("AA", referenceValue);

        value = 27;
		referenceValue = ExcelHelper.getReferenceForColumnIndex(value);
        System.out.println(value + " : " + referenceValue);
		Assert.assertEquals("AB", referenceValue);

        value = 728;
		referenceValue = ExcelHelper.getReferenceForColumnIndex(value);
        System.out.println(value + " : " + referenceValue);
		Assert.assertEquals("ABA", referenceValue);

        value = 729;
		referenceValue = ExcelHelper.getReferenceForColumnIndex(value);
        System.out.println(value + " : " + referenceValue);
		Assert.assertEquals("ABB", referenceValue);

        value = 18278;
		referenceValue = ExcelHelper.getReferenceForColumnIndex(value);
        System.out.println(value + " : " + referenceValue);
		Assert.assertEquals("AAAA", referenceValue);

        value = 459030;
        referenceValue = ExcelHelper.getReferenceForColumnIndex(value);
        System.out.println(value + " : " + referenceValue);
        Assert.assertEquals("ZCAA", referenceValue);
    }
}
