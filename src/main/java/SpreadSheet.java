/*
 * Copyright (c) ${year}, In Mind Computing AG. All rights reserved.
 * IMC PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 */
package main.java;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * This is a simple Java program that demonstrates the given use case to write
 * and read from excel. 1. The setCellValue handles numeric and formula inputs
 * 2. The getCellValue evaluates formula and returns value if cell type is
 * FORMULA. 3. Other edge cases to consider are formula involving divide by 0
 * etc., which is not yet handled below but can be included This can easily be
 * converted into a REST API GET and POST calls if needed.
 */
public class SpreadSheet
{

	public static void main(final String[] args)
	{
		final File currDir = new File(".");
		final String path = currDir.getAbsolutePath();
		final String fileLocation = path.substring(0, path.length() - 1) + "temp.xlsx";

		final SpreadSheet spreadSheet = new SpreadSheet();

		try
		{
			final FileOutputStream outputStream = new FileOutputStream(fileLocation);

			spreadSheet.getSheet().setColumnWidth(0, 6000);
			spreadSheet.getSheet().setColumnWidth(1, 4000);

			spreadSheet.setCellValue("A1", 13);
			spreadSheet.setCellValue("A2", 14);
			System.out.println("Value at A1 is:" + spreadSheet.getCellValue("A1"));

			spreadSheet.setCellValue("A3", "=A1+A2");
			System.out.println("Value at A3 is:" + spreadSheet.getCellValue("A3"));

			spreadSheet.setCellValue("A4", "=A1+A2+A3");
			System.out.println("Value at A4 is:" + spreadSheet.getCellValue("A4"));

			spreadSheet.getWorkbook().write(outputStream);
			spreadSheet.getWorkbook().close();
		}
		catch (final IOException e)
		{
			throw new RuntimeException(e);
		}

	}

	Sheet sheet;

	Workbook workbook;

	/**
	 * ${tags}
	 */
	public SpreadSheet()
	{
		super();

		workbook = new XSSFWorkbook();
		sheet = workbook.createSheet("Sample");

	}

	public int getCellValue(final String cellId)
	{
		final CellReference cr = new CellReference(cellId);
		final Row row = sheet.getRow(cr.getRow());
		final Cell cell = row.getCell(cr.getCol());

		switch (cell.getCellTypeEnum())
		{
		case NUMERIC:
			return (int) cell.getNumericCellValue();
		case FORMULA:
			final FormulaEvaluator formulaEval = workbook.getCreationHelper().createFormulaEvaluator();
			final CellValue c = formulaEval.evaluate(cell);
			return (int) c.getNumberValue();
		default:
			break;
		}

		return 0;

	}

	/**
	 * @return the sheet
	 */
	public Sheet getSheet()
	{
		return sheet;
	}

	/**
	 * @return the workbook
	 */
	public Workbook getWorkbook()
	{
		return workbook;
	}

	public void setCellValue(final String cellId, final Object value)
	{

		// final Row row = sheet.createRow(0);

		final CellReference cr = new CellReference(cellId);
		final Row row = sheet.createRow(cr.getRow());
		final Cell cell = row.createCell(cr.getCol());

		if (value instanceof String)
		{
			final String strValue = value.toString();
			if (strValue.contains("="))
			{
				final String formula = strValue.split("=")[1];
				cell.setCellFormula(formula);
			}
			else
				cell.setCellFormula((String) value);
		}
		else if (value instanceof Integer)
		{
			final int valueNew = ((Integer) value).intValue();
			cell.setCellValue(valueNew);
		}

	}

	/**
	 * @param ${param} to set
	 */
	public void setSheet(final Sheet sheet)
	{
		this.sheet = sheet;
	}

	/**
	 * @param ${param} to set
	 */
	public void setWorkbook(final Workbook workbook)
	{
		this.workbook = workbook;
	}

}
