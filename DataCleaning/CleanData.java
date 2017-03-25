import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CleanData {

	public static void writeXLSXFile() throws IOException {
		try {
			FileInputStream file = new FileInputStream("/home/rajat/workspace1/FinicialModeling/src/instruments.xlsm");
			//FileInputStream file = new FileInputStream("/home/rajat/Desktop/instruments_demo.xlsx");
			System.out.println("Loading workbook");
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			System.out.println("workbook loaded");
			XSSFSheet sheet = workbook.getSheetAt(1);
			System.out.println("Sheet loaded");
			FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
			System.out.println("Evaluator loaded");

			Cell cell = null;
			int noOfRow = sheet.getLastRowNum();
			int noOfCol = sheet.getRow(2).getLastCellNum();
			//int noOfRow = 49;
			//int noOfCol = 3;
			System.out.println("No of rowws and column "+noOfRow+"  "+noOfCol);
			double previous = -1;
			double after = -1;
			int noOfBlankVal = 0;
			//System.exit(0);

			//hardcoding column value due to null pointer exception 426*2+2
			for(int col =2;col<noOfCol;col++){            	

				for(int row=3;row<=noOfRow;row++){    
					// Access the cell value            
					cell = sheet.getRow(row).getCell(col);
					if(cell == null){						
						//replace null if its in between 1st and last row
						//compute previous and after value if next value is not blank
						if(noOfBlankVal == 0){
							if(row > 3){
							switch (evaluator.evaluateFormulaCell(cell)) {
							case Cell.CELL_TYPE_NUMERIC:
								previous = sheet.getRow(row-1).getCell(col).getNumericCellValue();					        	
								break;
							case Cell.CELL_TYPE_STRING:
								//debug
								System.out.println("String cell at "+row+" "+col+ " value "+sheet.getRow(row-1).getCell(col).getStringCellValue());

								break;
							default: System.out.println("In default pos "+row+ "  "+col);  //debug
								previous = sheet.getRow(row-1).getCell(col).getNumericCellValue();	
								  System.out.println("default value "+previous);
								break;

							}
							//set after = previous  in case last row is empty
							after = previous;
							}
							//check for consecutive blank elements
							for (int i=row+1; i<=noOfRow;i++ ){								
								System.out.println("i and col value "+i+"  "+col);
								Cell tmpCell = sheet.getRow(i).getCell(col);
								noOfBlankVal++;
								if(tmpCell != null){									
									after = tmpCell.getNumericCellValue();   
									System.out.println("Before value "+previous+" After value : "+after);
									//System.out.println("after value : "+sheet.getRow(i).getCell(col).getNumericCellValue());
									break;
								}
							}
							if(row == 3){
								previous = after;
							}
							if(row == noOfRow){
								noOfBlankVal++;
							}
							System.out.println("No of consecuteve blank : "+noOfBlankVal+" at position "+row);
						}
						//set avg value till for all consecutive empty value
						if (noOfBlankVal >= 1){
							cell = sheet.getRow(row).createCell(col, Cell.CELL_TYPE_NUMERIC);
							cell.setCellValue((previous+after)/2);
							System.out.println("Replacing value with "+previous+"  "+after+" at index :"+row);
							noOfBlankVal--;
						}

					}

				}
			}          
			//
			FileOutputStream outFile =new FileOutputStream(new File("/home/rajat/Desktop/ins_modified.xlsx"));
			workbook.write(outFile);
			outFile.close();
			file.close();

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		writeXLSXFile();
	}

}