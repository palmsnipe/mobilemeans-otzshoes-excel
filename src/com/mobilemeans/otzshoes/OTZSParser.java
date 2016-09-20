package com.mobilemeans.otzshoes;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.mobilemeans.otzshoes.models.Company;
import com.mobilemeans.otzshoes.models.Shoe;
import com.mobilemeans.otzshoes.models.Stock;

public class OTZSParser {
	
	private File mFile;
	
	public OTZSParser(File file) {
		mFile = file;
	}
	
	public Company parseFile() {
		if (mFile == null) return null;
		
		Company company = null;
		
		try {
			FileInputStream file = new FileInputStream(mFile);
			
	        XSSFWorkbook workbook = new XSSFWorkbook(file);

	        XSSFSheet sheet = workbook.getSheetAt(0);
	        
	        company = parseCompany(sheet);
	        
	        workbook.close();
	        file.close();
	        
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		return company;
	}
	
	private Company parseCompany(XSSFSheet sheet) {
		Company company = new Company();
		
		HashMap<Integer, String> header = new HashMap<Integer, String>();
		int startTable = 0;
		
		List<Shoe> shoes = new ArrayList<Shoe>();
		company.setShoes(shoes);
		
        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) 
        {
            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            
            Shoe shoe = null;
             
            while (cellIterator.hasNext()) 
            {
                Cell cell = cellIterator.next();
                cell.setCellType(Cell.CELL_TYPE_STRING);

                if (row.getRowNum() == 1) {
                	if (cell.getColumnIndex() == 0) {
                        cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                		if (DateUtil.isCellDateFormatted(cell)) {
                			Date date = cell.getDateCellValue();
                			
                			SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                			company.setDate(sdf.format(date));

                		}
                	}
                	else if (cell.getColumnIndex() == 4) {
                		company.setContactPerson(cell.getStringCellValue());
                	}
                	
                }
                else if (row.getRowNum() == 3) {
                	if (cell.getColumnIndex() == 0) {
                		company.setCompanyName(cell.getStringCellValue());
                	}
                	else if (cell.getColumnIndex() == 4) {
                		company.setEmail(cell.getStringCellValue());
                	}
                }
                else if (row.getRowNum() == 5) {
                	if (cell.getColumnIndex() == 0) {
                		company.setVatnumber(cell.getStringCellValue());
                	}
                	else if (cell.getColumnIndex() == 4) {
                		company.setPhone(cell.getStringCellValue());
                	}
                }
                else if (row.getRowNum() == 7) {
                	if (cell.getColumnIndex() == 0) {
                		company.setStoreName(cell.getStringCellValue());
                	}
                }
                else if (row.getRowNum() >= 13 && cell.getColumnIndex() == 0 && cell.getStringCellValue().toLowerCase().contains("name")) {
                	startTable = row.getRowNum();
                }
                
                if (startTable > 0) {
                
	                if (row.getRowNum() == startTable) {
	                	if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
	                    	  if (!cell.getStringCellValue().isEmpty()) {
	                    		  header.put(cell.getColumnIndex(), cell.getStringCellValue().toLowerCase());
	                    	  }
	                    	  else
	                    		  header.put(cell.getColumnIndex(), ""); 
	                	}
	                }
	                else if (cell.getColumnIndex() == 0 && cell.getStringCellValue().contains("TOTAL")) {
	                	return company;
	                }
	                else {
	                	if (header.get(cell.getColumnIndex()).contains("name")) {
	                		String name = cell.getStringCellValue();
	            			shoe = new Shoe();
	            			shoe.setName(name.trim());
	                	}
	                	else if (header.get(cell.getColumnIndex()).equals("code")) {
	                		shoe.setCode(cell.getStringCellValue());
	                	}
	                	else if (isNumeric(header.get(cell.getColumnIndex()))) {
	                		int size = Integer.parseInt(header.get(cell.getColumnIndex()));
	                		String quantity = cell.getStringCellValue();
	                		if (quantity != null && quantity.length() > 0) {
	                			Stock stock = new Stock();
	                			
	                			stock.setSize(size);
	                			stock.setQuantity(Integer.parseInt(quantity));
	                			shoe.addStock(stock);
	                		}
	                		
	                	}
	                	else if (!cellIterator.hasNext()) {
	                		if (shoe != null && shoe.getCode() != null && shoe.getCode().trim().length() > 0) {
	                        	shoes.add(shoe);
	                        }
	                	}
	                	
	                }
                }
                
            }
            
        }
		
		return company;
	}
	
	public static boolean isNumeric(String str)
	{
	  return str.matches("-?\\d+(\\.\\d+)?");
	}
	
}
