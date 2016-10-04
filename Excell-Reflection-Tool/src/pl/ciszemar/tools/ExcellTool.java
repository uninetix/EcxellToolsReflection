package pl.ciszemar.tools;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;

public class ExcellTool {

	public <E> void writeObjectList(File file, ArrayList<E> lista) {
		HSSFWorkbook workbook = new HSSFWorkbook();
		Class<?> c = lista.get(0).getClass();
		HSSFSheet sheet = workbook.createSheet(c.getSimpleName());
		HSSFCellStyle style = workbook.createCellStyle();
		Row row = sheet.createRow(0);
		if (lista != null)
			if (lista.get(0) != null) {
				Field[] fields = c.getDeclaredFields();
				int cellnum = 0;
				style.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);
				style.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
				style.setBorderLeft(HSSFCellStyle.BORDER_MEDIUM);
				style.setBorderRight(HSSFCellStyle.BORDER_MEDIUM);
//				style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
				style.setFillForegroundColor(IndexedColors.BLUE.getIndex());
				style.setFillBackgroundColor(IndexedColors.RED.getIndex());
				for (Field item : fields) {
					Cell cell = row.createCell(cellnum++);
					cell.setCellValue(item.getName());
					cell.setCellStyle(style);
				}
			} else
				return;

		int rownum = 1;
		style = workbook.createCellStyle();
		style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style.setBorderRight(HSSFCellStyle.BORDER_THIN);

		for (E item : lista) {
			Class<?> clazz = item.getClass();
			Field[] fields = clazz.getDeclaredFields();
			int cellnum = 0;
			row = sheet.createRow(rownum);
			for (Field fieldItem : fields) {
				fieldItem.setAccessible(true);
				Field field;
				try {
					field = clazz.getDeclaredField(fieldItem.getName());
					field.setAccessible(true);
					Object fieldValue = field.get(item);
					Cell cell = row.createCell(cellnum);
					cell.setCellValue(fieldValue.toString());
					cell.setCellStyle(style);
				} catch (NoSuchFieldException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (SecurityException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IllegalArgumentException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IllegalAccessException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				cellnum++;
			}
			rownum++;
		}
		try {
			FileOutputStream out = new FileOutputStream(file);
			workbook.write(out);
			out.close();
			System.out.println("Excel written successfully..");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public <E> void writeObjectList(File file, ArrayList<E> lista, Object[] naglowek) {
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("KWM Dane monza");
		int cellnum = 0;
		Row row = sheet.createRow(0);
		HSSFCellStyle style = workbook.createCellStyle();
		style.setBorderBottom(HSSFCellStyle.BORDER_MEDIUM);
		style.setBorderTop(HSSFCellStyle.BORDER_MEDIUM);
		style.setBorderLeft(HSSFCellStyle.BORDER_MEDIUM);
		style.setBorderRight(HSSFCellStyle.BORDER_MEDIUM);
//		style.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
		style.setFillForegroundColor(IndexedColors.BLUE.getIndex());
		style.setFillBackgroundColor(IndexedColors.RED.getIndex());
		for (Object item : naglowek) {
			Cell cell = row.createCell(cellnum++);
			cell.setCellValue((String) item);
			cell.setCellStyle(style);
		}

		int rownum = 1;
		style = workbook.createCellStyle();
		style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style.setBorderRight(HSSFCellStyle.BORDER_THIN);

		for (E item : lista) {
			Class<?> clazz = item.getClass();
			Field[] fields = clazz.getDeclaredFields();
			cellnum = 0;
			row = sheet.createRow(rownum);
			for (Field fieldItem : fields) {
				fieldItem.setAccessible(true);
				Field field;
				try {
					field = clazz.getDeclaredField(fieldItem.getName());
					field.setAccessible(true);
					Object fieldValue = field.get(item);
					Cell cell = row.createCell(cellnum);
					cell.setCellValue(fieldValue.toString());
					cell.setCellStyle(style);
				} catch (NoSuchFieldException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (SecurityException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IllegalArgumentException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (IllegalAccessException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				cellnum++;
			}
			rownum++;
		}
		try {
			FileOutputStream out = new FileOutputStream(file);
			workbook.write(out);
			out.close();
			System.out.println("Excel written successfully..");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
