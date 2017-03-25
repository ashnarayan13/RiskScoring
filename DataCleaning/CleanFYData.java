import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CleanFYData {

	int noOfColToBeSearched = 5;
	int noOfRow = 0;
	int noOfCol = 0;
	boolean writeSuccess = false;

	XSSFWorkbook workbook;
	XSSFSheet sheet;
	FileInputStream file;
	FileOutputStream outFile;

	public void readAndwriteExcel(String action) throws IOException{

		if (action.equalsIgnoreCase("read")) {
			file = new FileInputStream("/home/rajat/Desktop/instruments.xlsm");
			System.out.println("Loading workbook");
			workbook = new XSSFWorkbook(file);
			System.out.println("workbook loaded");

		}else if (action.equalsIgnoreCase("write")){
			outFile =new FileOutputStream(new File("/home/rajat/Desktop/FY_modified.xlsx"));
			workbook.write(outFile);	
			outFile.close();
			file.close();
		}

	}

	private void cleanData() throws IOException{
		Cell cell = null;
		readAndwriteExcel("read");
		boolean foundClosevalue;

		for(int sht =3;sht<9;sht++){     // loop for 6 financial year
			sheet = workbook.getSheetAt(sht);
			noOfRow = sheet.getLastRowNum();
			noOfCol = sheet.getRow(3).getLastCellNum();
			System.out.println("No of rowws and column " + noOfRow + "  " + noOfCol +" at sheet no "+sht );
			for(int col =3;col<8;col++){ 	    						//loop for till total revenue
				for(int row = 3;row<noOfRow;row++){						// for all row present in the sheet
					//check for the null value for each column 

					cell = sheet.getRow(row).getCell(col);
					if(cell.getCellType() == Cell.CELL_TYPE_STRING ){

						if(cell.getStringCellValue().equals(new String("NULL"))){
							System.out.println(" found NULL at pos : "+row+" , "+col+" at sheet no : "+sht);
							int closeValueCol = col;
							foundClosevalue = false;
							while(!foundClosevalue && closeValueCol <= 6){       //replace column till Normalized Income After Taxes as total revenue may not be same with total debt	

								closeValueCol++;
								int nearRowIndex = getNearestValueInRow(sheet,row,closeValueCol,noOfRow); 
								if (nearRowIndex != -1){
									Cell closestValueCell = sheet.getRow(nearRowIndex).getCell(col);       //check corresponding closest value cell,if its null
									if(closestValueCell !=null && closestValueCell.getCellType() == Cell.CELL_TYPE_NUMERIC){										     //then check next column value's corresponding closest and so on 
										System.out.println("close value to be replace with "+closestValueCell.getNumericCellValue()+" and position "+nearRowIndex);
										cell.setCellValue(closestValueCell.getNumericCellValue());
										foundClosevalue = true;
									}
								}
							}
						}
					}

				}
			}
		}	
		//readAndwriteExcel("write");
		writeSuccess = true;
		if(writeSuccess){
			replaceNULLForTotalRevenue();
		}
		readAndwriteExcel("write");
		//file.close();
		//outFile.close();
	}

	// find the nearest possible match in a row, and return its row index
	private int getNearestValueInRow(XSSFSheet sheet, int row, int col, int noOfRow) {		
		double minVal = -1;
		int nearValueIndex = -1;


		Cell cell1 = sheet.getRow(row).getCell(col);						
		if(cell1!= null && cell1.getCellType() != Cell.CELL_TYPE_STRING){   											//if the next col is also blank,proceed to next col

			double currentVallue = cell1.getNumericCellValue();
			System.out.println("current value "+currentVallue);
			//get minimum value in row
			for(int j=3;j<noOfRow;j++){
				double tempVal = -1;
				if(j != row){
					Cell tmpCell = sheet.getRow(j).getCell(col);
					if(tmpCell != null && tmpCell.getCellType() != Cell.CELL_TYPE_STRING){
						tempVal = tmpCell.getNumericCellValue();					

						if(currentVallue > tempVal){
							tempVal = currentVallue - tempVal;
						}else{
							tempVal = tempVal - currentVallue;
						}
					}				
					//set minimum value 
					if(minVal>tempVal ){
						minVal = tempVal;
						nearValueIndex = j;
					}else if(minVal == -1 && tempVal > 0){
						minVal = tempVal;
						nearValueIndex = j;
					}
				}
			}
			System.out.println("Near value  row "+nearValueIndex +"and col "+col);
			System.out.println("the min value : "+minVal+" in row "+nearValueIndex+" and col "+col+" and the actual value at that col "+sheet.getRow(nearValueIndex).getCell(col).getNumericCellValue());
		}
		return nearValueIndex;
	}

	private void replaceNULLForTotalRevenue() throws IOException{

		boolean foundClosevalue;
		int colOfTotalRevenue = 7;				//get cell of total revenue
		for(int sht =3;sht<9;sht++){
			
			//workbook = new XSSFWorkbook(file);
			sheet = workbook.getSheetAt(sht);
			noOfRow = sheet.getLastRowNum();
			for(int row =3;row <noOfRow;row++){
				XSSFCell cell = sheet.getRow(row).getCell(colOfTotalRevenue);             
				if(cell.getCellType() == Cell.CELL_TYPE_STRING ){
					if(cell.getStringCellValue().equals(new String("NULL"))){
						System.out.println(" found NULL for total revenue at pos : "+row+" , at sheet no : "+sht);
						int closeValueCol = colOfTotalRevenue;
						foundClosevalue = false;
						while(!foundClosevalue && closeValueCol >= 3){      
							closeValueCol--;
							int nearRowIndex = getNearestValueInRow(sheet,row,closeValueCol,noOfRow); 	
							if (nearRowIndex != -1){
								Cell closestValueCell = sheet.getRow(nearRowIndex).getCell(colOfTotalRevenue);       //check corresponding closest value cell,if its null
								if(closestValueCell !=null && closestValueCell.getCellType() == Cell.CELL_TYPE_NUMERIC){										     //then check next column value's corresponding closest and so on 
									System.out.println("close value to be replace with "+closestValueCell.getNumericCellValue()+" and position "+nearRowIndex);
									cell.setCellValue(closestValueCell.getNumericCellValue());
									foundClosevalue = true;
								}
							}
						}
					}
				}

			}
		}
		
	}

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		CleanFYData cData = new CleanFYData();
		try {
			cData.cleanData();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
