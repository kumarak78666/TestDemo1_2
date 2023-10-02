package com.ibt.utilities;



	import java.io.FileInputStream;
	import java.io.FileNotFoundException;
	import java.io.FileOutputStream;
	import java.io.IOException;
	import java.io.OutputStream;
	import java.util.Calendar;

	import org.apache.poi.hssf.usermodel.HSSFCellStyle;
	import org.apache.poi.hssf.usermodel.HSSFDateUtil;
	import org.apache.poi.hssf.usermodel.HSSFWorkbook;
	import org.apache.poi.hssf.util.HSSFColor;
	import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
	import org.apache.poi.ss.usermodel.Cell;
	import org.apache.poi.ss.usermodel.Row;
	import org.apache.poi.ss.usermodel.Sheet;
	import org.apache.poi.ss.usermodel.Workbook;
	import org.apache.poi.ss.usermodel.WorkbookFactory;

	public class ExcelReader {

		static String excelFilePath;
		static Workbook workbook;
		static Sheet worksheet;
		static String sheet;
		static Row row;
		static Cell cell;
		static FileInputStream fis;
		static FileOutputStream fileOut;
		static int testCaseNameRowNo;
		public ExcelReader() {

		}

		public ExcelReader(String excelFilePath) throws InvalidFormatException,FileNotFoundException, IOException {
			try{
				ExcelReader.excelFilePath = excelFilePath;//this.excelFilePath = excelFilePath;
			//System.out.println(excelFilePath);
			workbook = WorkbookFactory.create(new FileInputStream(excelFilePath));	
			}catch (Exception e) {
				e.printStackTrace();

			}
		}
		
		//To retrieve No Of rows from .xls file's sheets.
		public static int getRowCount(String sheetName) {

			int rowCount = 0;
			try {
				
				sheet = sheetName;
				rowCount = workbook.getSheet(sheetName).getLastRowNum()+ 1;

			} catch (Exception e) {
				e.printStackTrace();

			}

			return rowCount;

		}
		
		//To retrieve No Of Columns from .xls file's sheets.
		public static int getColumnCount(String wsName){
			int sheetIndex=workbook.getSheetIndex(wsName);
			if(sheetIndex==-1)
				return 0;
			else{
				worksheet = workbook.getSheetAt(sheetIndex);
				int colCount=worksheet.getRow(0).getLastCellNum();			
				return colCount;
			}
		}

		//Method to get the cell data 
		@SuppressWarnings("static-access")
		public static String getCellData(String sheetName, int rowNumber, int columnNumber) {
			String cellValue = "";
			try {
				if(rowNumber <=0)
	                return "";
				if (workbook.getSheet(sheetName).getRow(rowNumber).getCell(columnNumber).getCellType() == workbook.getSheet(sheetName).getRow(rowNumber).getCell(columnNumber).CELL_TYPE_STRING) {

					cellValue = workbook.getSheet(sheetName).getRow(rowNumber)
							.getCell(columnNumber).getStringCellValue().trim();
				} else {

					cellValue = String.valueOf(
							(int) (workbook.getSheet(sheetName).getRow(rowNumber)
									.getCell(columnNumber).getNumericCellValue()))
							.trim();
				}
			} catch (Exception e) {

				return cellValue;
			}

			return cellValue;
		}
		
		// returns the data from a cell
	    public static String getCellData(String sheetName,String colName,int rowNum){
	        try{
	        	
	            if(rowNum <= 0)
	                return "";

	            int index = workbook.getSheetIndex(sheetName);
	            //System.out.println("index"+index);
	            int col_Num=-1;
	            if(index==-1)
	                return "";

	            worksheet = workbook.getSheetAt(index);
	            row=worksheet.getRow(0);
	           // System.out.println("Row="+row.getRowNum());
	          //  System.out.println("Last cell="+row.getLastCellNum());
	            if(row!=null){
		            for(int i=0;i<row.getLastCellNum();i++){
		                //System.out.println(row.getCell(i).getStringCellValue().trim());
		              if(row.getCell(i)!=null){	//if cell is not null i,e to handle Null Pointer Exception
		                if(row.getCell(i).getStringCellValue().trim().equalsIgnoreCase(colName.trim()))
		                    col_Num=i;
		                	//System.out.println(col_Num);
		              }
		            }
	            }
	            if(col_Num==-1)
	                return "";

	            worksheet = workbook.getSheetAt(index);
	            row = worksheet.getRow(rowNum-1);
	            if(row==null)
	                return "";
	            cell = row.getCell(col_Num);

	            if(cell==null)
	                return "";
	            //System.out.println(cell.getCellType());
	            if(cell.getCellType()==Cell.CELL_TYPE_STRING)
	                return cell.getStringCellValue();
	            else if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC || cell.getCellType()==Cell.CELL_TYPE_FORMULA ){

	                String cellText  = String.valueOf(cell.getNumericCellValue());
	                if (HSSFDateUtil.isCellDateFormatted(cell)) {
	                    // format in form of M/D/YY
	                    double d = cell.getNumericCellValue();

	                    Calendar cal =Calendar.getInstance();
	                    cal.setTime(HSSFDateUtil.getJavaDate(d));
	                    cellText =
	                            (String.valueOf(cal.get(Calendar.YEAR))).substring(2);
	                    cellText = cal.get(Calendar.DAY_OF_MONTH) + "/" +
	                            cal.get(Calendar.MONTH)+1 + "/" +
	                            cellText;

	                    //System.out.println(cellText);

	                }

	                return cellText;
	            }else if(cell.getCellType()==Cell.CELL_TYPE_BLANK)
	                return "";
	            else
	                return String.valueOf(cell.getBooleanCellValue());

	        }
	        catch(Exception e){

	            e.printStackTrace();
	            return "row "+rowNum+" or column "+colName +" does not exist in xls";
	        }
	    }

	    
	    //Method to set the cell value
		public static void setCellData(String sheetName, int rowNumber, int columnNumber,
				String value) {

			try {
				workbook.getSheet(sheetName).getRow(rowNumber)
						.createCell(columnNumber).setCellValue(value);
				OutputStream outputStream = new FileOutputStream(excelFilePath);
				workbook.write(outputStream);

				outputStream.close();

			} catch (Exception e) {
				e.printStackTrace();
			}

		}
		
		// returns true if data is set successfully else false
	    public static boolean setCellData(String sheetName,String colName,int rowNum, String data){
	        try{
	            fis = new FileInputStream(excelFilePath);
	            workbook = new HSSFWorkbook(fis);

	            if(rowNum<=0)
	                return false;

	            int index = workbook.getSheetIndex(sheetName);
	            int colNum=-1;
	            if(index==-1)
	                return false;

	            worksheet = workbook.getSheetAt(index);
	            row=worksheet.getRow(0);
	            if(row!=null){
	            	for(int i=0;i<row.getLastCellNum();i++){
	            		if(row.getCell(i)!=null){	//if cell is not null i,e to handle Null Pointer Exception
	            			//System.out.println(row.getCell(i).getStringCellValue().trim());
	            			if(row.getCell(i).getStringCellValue().trim().equalsIgnoreCase(colName))
	            				colNum=i;
	            		}	
	            	}
	            }
	            if(colNum==-1)
	                return false;

	            worksheet.autoSizeColumn(colNum);
	            row = worksheet.getRow(rowNum-1);
	            if (row == null)
	                row = worksheet.createRow(rowNum-1);

	            cell = row.getCell(colNum);
	            if (cell == null)
	                cell = row.createCell(colNum);

	            cell.setCellValue(data);

	            fileOut = new FileOutputStream(excelFilePath);

	            workbook.write(fileOut);

	            fileOut.close();

	        }
	        catch(Exception e){
	            e.printStackTrace();
	            return false;
	        }
	        return true;
	    }
	    
		// returns true if sheet is created successfully else false
	    public boolean addSheet(String  sheetname){

	        FileOutputStream fileOut;
	        try {
	            workbook.createSheet(sheetname);
	            fileOut = new FileOutputStream(excelFilePath);
	            workbook.write(fileOut);
	            fileOut.close();
	        } catch (Exception e) {
	            e.printStackTrace();
	            return false;
	        }
	        return true;
	    }
	    
	    
	    // returns true if sheet is removed successfully else false if sheet does not exist
	    public boolean removeSheet(String sheetName){
	        int index = workbook.getSheetIndex(sheetName);
	        if(index==-1)
	            return false;

	        
	        try {
	            workbook.removeSheetAt(index);
	            fileOut = new FileOutputStream(excelFilePath);
	            workbook.write(fileOut);
	            fileOut.close();
	        } catch (Exception e) {
	            e.printStackTrace();
	            return false;
	        }
	        return true;
	    }
	    
	    // find whether sheets exists
	    public boolean isSheetExist(String sheetName){
	        int index = workbook.getSheetIndex(sheetName);
	        if(index==-1){
	            index=workbook.getSheetIndex(sheetName.toUpperCase());
	            if(index==-1)
	                return false;
	            else
	                return true;
	        }
	        else
	            return true;
	    }
	    
	    
	    // returns true if column is created successfully
	    public boolean addColumn(String sheetName,String colName){
	        try{
	            fis = new FileInputStream(excelFilePath);
	            workbook = new HSSFWorkbook(fis);
	            int index = workbook.getSheetIndex(sheetName);
	            if(index==-1)
	                return false;

	            HSSFCellStyle style = (HSSFCellStyle) workbook.createCellStyle();
	            style.setFillForegroundColor(HSSFColor.GREY_40_PERCENT.index);
	            style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

	            worksheet=workbook.getSheetAt(index);

	            row = worksheet.getRow(0);
	            if (row == null)
	                row = worksheet.createRow(0);

	            //cell = row.getCell();
	            //if (cell == null)
	            //System.out.println(row.getLastCellNum());
	            if(row.getLastCellNum() == -1)
	                cell = row.createCell(0);
	            else
	                cell = row.createCell(row.getLastCellNum());

	            cell.setCellValue(colName);
	            cell.setCellStyle(style);

	            fileOut = new FileOutputStream(excelFilePath);
	            workbook.write(fileOut);
	            fileOut.close();

	        }catch(Exception e){
	            e.printStackTrace();
	            return false;
	        }

	        return true;


	    }
	    
	    public static String getCellData(String sheetName,String columnName){
	    	int rowNumber = 1;
	        int index = workbook.getSheetIndex(sheetName);
	        if(index==-1)
	            return "";

	        worksheet = workbook.getSheetAt(index);
	        row = worksheet.getRow(getTestCaseNameRowNo()); 
	        
	        //if Runmode=y then pic the data from that row y=rownum
	        //to get the row number i,e if RunMode=y then that row number will become rowNum to get the data
	        for(int i=getTestCaseNameRowNo();i<getRowCount(sheetName);i++){             	
	        	//System.out.println("Row count="+getRowCount(sheetName));            		
				//System.out.println(workbook.getSheet(sheetName).getRow(i).getCell(0));
	        	//check RunMode=y for first column, if found that will be rowNum to get the data
				if(workbook.getSheet(sheetName).getRow(i).getCell(0).getStringCellValue().trim().equalsIgnoreCase("y")){            		            				
					rowNumber = i;
					break;//if RunMode=y found break
				}
	        }
	        
	        String runMode = getCellText(worksheet, rowNumber, 0);
	        if(runMode == null || !runMode.equalsIgnoreCase("y")){     		            				
				return "";
			}
	  		return getCellText(worksheet, rowNumber, getColumnNumber(columnName));
	    }
	    
	    //to set the testcase name row in excel sheet because the next row after testcase name will become column heading
	    public static void setTestCaseNameRowNo(int rowNo){
	    	testCaseNameRowNo= rowNo;
	    }
	    
	    //returns the testcase name row from excel
	    public static int getTestCaseNameRowNo(){
	    	return testCaseNameRowNo;
	    }
	  	
	  	public static String getCellData(int rowNumber, String sheetName, String columnName) {
	  		int index = workbook.getSheetIndex(sheetName);
	        if(index == -1)
	            return "";
		
	        worksheet = workbook.getSheetAt(index);
	        String runMode = getCellText(worksheet, rowNumber, 0);
	        if(runMode == null || !runMode.equalsIgnoreCase("y")){     		            				
				return "";
			}
	  		return getCellText(worksheet, rowNumber, getColumnNumber(columnName));
	  		
	  	}
	  	
	  	private static int getColumnNumber(String columnName) {
	  		row = worksheet.getRow(getTestCaseNameRowNo()); 
	        
	        //to get the column number based on ColumnName passed as argument
	        if(row != null){
	            for(int i=0;i<row.getLastCellNum();i++){
	              //System.out.println(row.getCell(i).getStringCellValue().trim());
	              if(row.getCell(i)!=null){	//if cell is not null i,e to handle Null Pointer Exception
	                if(row.getCell(i).getStringCellValue().trim().equalsIgnoreCase(columnName.trim()))
	                    return i;
	              }
	            }
	        }
	        return -1;
	  	}
	  	
	  	private static String getCellText(Sheet worksheet, int rowNumber, int columnNumber) {
	  		try {
	  			if(columnNumber < 0 || rowNumber < 0) {
	  				return "";
	  			}
	  			
	  			row = worksheet.getRow(rowNumber);
		        if(row == null)
		             return "";
		        
		         cell = row.getCell(columnNumber);
		         if(cell == null)
		             return "";
		         if(cell.getCellType() == Cell.CELL_TYPE_STRING)
		             return cell.getStringCellValue();
		         else if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC || cell.getCellType()==Cell.CELL_TYPE_FORMULA ){
		             String cellText  = String.valueOf(cell.getNumericCellValue());
		             if (HSSFDateUtil.isCellDateFormatted(cell)) {
		                 // format in form of M/D/YY
		                 double d = cell.getNumericCellValue();
		
		                 Calendar cal =Calendar.getInstance();
		                 cal.setTime(HSSFDateUtil.getJavaDate(d));
		                 cellText =
		                         (String.valueOf(cal.get(Calendar.YEAR)));
		                 cellText = cal.get(Calendar.DAY_OF_MONTH) + "/" +
		                         cal.get(Calendar.MONTH)+1 + "/" +
		                         cellText;
		             }
		             return cellText;
		         }else if(cell.getCellType()==Cell.CELL_TYPE_BLANK)
		             return "";
		         else
		             return String.valueOf(cell.getBooleanCellValue());
	  		}catch(Exception e){
	  			e.printStackTrace();
	  			return "";
	  		}
	  	}
	}



