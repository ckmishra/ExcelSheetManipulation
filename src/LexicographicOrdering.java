import java.io.Closeable;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 
 * @author Chandan Mishra
 *  
 */
public class LexicographicOrdering {
	
	private static final String INPUT_FILE_PATH = "/Users/apple/Desktop/Assignment1/R_Exported_Data.xlsx";
	private static final String OUTPUT_FILE_PATH = "/Users/apple/Desktop/Assignment1/Output_1.xlsx";
	// main method 
	public static void main(String[] args) {
		processOutputDatafromRtoExcel();
	}
	/**
	 * Method for reading Excel and processing for sorting
	 */
	private static void processOutputDatafromRtoExcel() {
		FileInputStream fis= null;
		FileOutputStream fos=null;
		try {
			fis = new FileInputStream(INPUT_FILE_PATH);
			fos = new FileOutputStream(OUTPUT_FILE_PATH);

			// Using XSSF for xlsx format, for xls use HSSF
			Workbook workbook = new XSSFWorkbook(fis);
			Sheet sheet = workbook.getSheetAt(0);
			Iterator rowIterator = sheet.iterator();

			while (rowIterator.hasNext()) {
				Row row = (Row) rowIterator.next();
				Iterator cellIterator = row.cellIterator();
				Cell cell = (Cell) cellIterator.next();
				String cellVaue = cell.getStringCellValue();
				String arrayOfInt = cellVaue
						.substring(1, cellVaue.length() - 1);
				String[] splitedString = arrayOfInt.split(",");
				List<Integer> temp = new ArrayList<Integer>();
				for (int i = 0; i < splitedString.length; i++) {
					temp.add(Integer.valueOf(splitedString[i]));
				}
				
				// sort
				Collections.sort(temp);
				String sortedString = "{";

				for (int i = 0; i < temp.size(); i++) {
					sortedString = sortedString.concat(temp.get(i) + ",");
				}
				sortedString = sortedString.concat("}");
				sortedString = sortedString.replaceAll(",}", "}");
				cell.setCellValue(sortedString);
			}
			sortLexicographically(sheet);
			workbook.write(fos);
			
		} catch (IOException e) {
			System.out.println("Exception occured while processing excel");
			e.printStackTrace();
		} finally {
			close(fos);
			close(fis);
		}
	}
	
	private static void close(Closeable closable) {
	    if (closable != null) {
	        try {
				closable.close();
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	    }
	}

	// sorting based on support
	private static void sortLexicographically(Sheet s) {
		int len = s.getLastRowNum();
		for (int i = 0; i < len; i++) {
			for (int j = i + 1; j <= len; j++) {
				double support_ith = s.getRow(i).getCell(1)
						.getNumericCellValue();
				double support_jth = s.getRow(j).getCell(1)
						.getNumericCellValue();
				if (support_ith == support_jth) {
					String ithItemSets = s.getRow(i).getCell(0)
							.getStringCellValue();
					String jthItemSets = s.getRow(j).getCell(0)
							.getStringCellValue();
					if (swap(s, ithItemSets, jthItemSets)) {
						s.getRow(i).getCell(0).setCellValue(jthItemSets);
						s.getRow(j).getCell(0).setCellValue(ithItemSets);
					}
				}
			}
		}
	}
	// helper method
	private static boolean swap(Sheet s, String ithItemSets, String jthItemSets) {

		List<Integer> temp_i = new ArrayList<Integer>();
		List<Integer> temp_j = new ArrayList<Integer>();
		// ith
		String arrayOfInt_i = ithItemSets
				.substring(1, ithItemSets.length() - 1);
		String[] splitedString = arrayOfInt_i.split(",");
		for (int k = 0; k < splitedString.length; k++) {
			temp_i.add(Integer.valueOf(splitedString[k]));
		}
		// jth
		String arrayOfInt_j = jthItemSets
				.substring(1, jthItemSets.length() - 1);
		splitedString = arrayOfInt_j.split(",");
		for (int k = 0; k < splitedString.length; k++) {
			temp_j.add(Integer.valueOf(splitedString[k]));
		}
		int length_i = temp_i.size();
		int length_j = temp_j.size();
		int len = length_i > length_j ? length_j : length_i;
		boolean isSwapRequired = false;
		
		for (int i = 0; i < len; i++) {
			if (temp_i.get(i) > temp_j.get(i)) {
				isSwapRequired = true;
				break;
			} else if (temp_i.get(i) < temp_j.get(i)) {
				isSwapRequired = false;
				break;
			}
		}
		return isSwapRequired;
	}

}
