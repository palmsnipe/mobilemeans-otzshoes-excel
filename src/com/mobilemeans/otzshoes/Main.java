package com.mobilemeans.otzshoes;

import java.awt.BorderLayout;
import java.awt.dnd.DropTarget;
import java.beans.PropertyChangeEvent;
import java.beans.PropertyChangeListener;
import java.io.File;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.ResourceBundle;

import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JProgressBar;
import javax.swing.SwingConstants;
import javax.swing.SwingWorker;
import javax.swing.UIManager;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.mobilemeans.otzshoes.DragDropListener.DragObserver;
import com.mobilemeans.otzshoes.models.Company;
import com.mobilemeans.otzshoes.models.Output;
import com.mobilemeans.otzshoes.models.Shoe;
import com.mobilemeans.otzshoes.models.Stock;


public class Main implements DragObserver {
	
	private static final String OUTPUT = "OTZSHOES - EXCEL RESULT.xlsx";
	
	private ResourceBundle mRes;
	
	private JFrame mFrame;
	private JProgressBar mProgressBar;
	
	private SwingWorker<Void, String> mWorker;
		
	public Main() {
		Locale locale = Locale.getDefault();
	    mRes = ResourceBundle.getBundle("com.mobilemeans.otzshoes/Resources", locale);
	    
		gui();
	}
	
	public void gui() {
		mFrame = new JFrame("OTZSHOES Excel");
	    mFrame.pack();
	    mFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

	    mFrame.setSize(500, 350);
	    
	    mFrame.setResizable(false);
	    
	    mFrame.setLocationRelativeTo(null);

	    JLabel myLabel = new JLabel(mRes.getString("dropLabel"), SwingConstants.CENTER);

	    DragDropListener myDragDropListener = new DragDropListener(this);

	    new DropTarget(myLabel, myDragDropListener);
	    
	    mProgressBar = new JProgressBar(0, 100);

	    mFrame.getContentPane().add(BorderLayout.CENTER, myLabel);
	    mFrame.getContentPane().add(BorderLayout.SOUTH, mProgressBar);
	    
	    mFrame.setVisible(true);
	}
	

	@Override
	public void folderDropped(final File folder) {
		
		mWorker = new SwingWorker<Void, String>() {
			
			@Override
			protected Void doInBackground() {
				List<Company> list = new ArrayList<>();
				List<String> errors = new ArrayList<String>();
				List<String> warnings = new ArrayList<String>();
				List<Output> output = new ArrayList<>();
				
				int totalShoes = 0;
				BigDecimal totalPriceVAT = BigDecimal.ZERO;
				int i = 0;
				
				File[] files = folder.listFiles();
				float divisor = files.length / 100.0f;
				for (final File fileEntry : files) {
			        if (!fileEntry.isDirectory() && !fileEntry.isHidden() && !fileEntry.getName().startsWith("~") && fileEntry.getName().endsWith(".xlsx")) {
			        	OTZSParser parser = new OTZSParser(fileEntry);
			        	
			        	Company company = parser.parseFile();
			        	
			        	parser = null;
			    		
			        	if (company != null) {
			        		list.add(company);
			        		
			        		for (Shoe shoe : company.getShoes()) {
			        			if (!shoe.getCode().matches("\\d+-\\d+")) {
			        				System.out.println("## " + fileEntry.getName() + " - " + shoe.getCode());
			        				errors.add(String.format("%s - The code \"%s\" is incorrect (not included in the results)", fileEntry.getName(), shoe.getCode()));
			        			}
			        			else {
			        				int stock = shoe.getTotalStock();
			        				totalShoes += stock;
			        				
	        						Double priceExVAT = shoe.getPriceExVAT();
	        						if (priceExVAT.compareTo(0.0) <= 0)
	        							errors.add(String.format("%s - The shoe price excluded VAT for the code \"%s\" is %.2f euro. (the total amount will be wrong!)", fileEntry.getName(), shoe.getCode(), priceExVAT));
	        						
	        						HashMap<String, String> errorPrice = new HashMap<>();
	        						totalPriceVAT = totalPriceVAT.add(new BigDecimal("" + priceExVAT).multiply(new BigDecimal(stock)));
				        			for (Stock size : shoe.getStock()) {
				        				boolean added = false;
				        				for (Output item : output) {
				        					if (item.getCode().equals(shoe.getCode()) && item.getSize() == size.getSize()) {
				        						item.setQuantity(item.getQuantity() + size.getQuantity());
				        						
				        						if (item.getName().length() < shoe.getName().length())
				        							item.setName(shoe.getName());
				        						
//				        						if (priceExVAT.compareTo(item.getPriceExVAT()) != 0) {
//				        							errorPrice.put(shoe.getCode(), String.format("%s - The shoe price excluded VAT for the code \"%s\" is not the same in an other file.", fileEntry.getName(), shoe.getCode()));
//				        						}
				        						
				        						added = true;
				        						break;
				        					}
				        				}
				        				if (!added) {
				        					Output newOutput = new Output();
				        					
				        					newOutput.setCode(shoe.getCode());
				        					newOutput.setName(shoe.getName());
				        					newOutput.setSize(size.getSize());
				        					newOutput.setQuantity(size.getQuantity());
				        					newOutput.setPriceExVAT(shoe.getPriceExVAT());
				        					
				        					output.add(newOutput);
				        				}
				        				
				        			}
				        			for (String value : errorPrice.values()) {
				        			    warnings.add(value);
				        			}
			        			}
			        		}
			        	}
			        	else {
			        		System.out.println("Couldn't parse the file: " + fileEntry.getName());
			        		errors.add(String.format("WARNING! the file \"%s\" couldn't be parsed", fileEntry.getName()));
			        	}
			        }
			        setProgress((int) (++i / divisor));
			    }
				
				// Sort the output
				Collections.sort(output, new Comparator<Output>() {
				    @Override
				    public int compare(Output o1, Output o2) {
				    	int result = 0;
				    	if (o1.getCode().equals(o2.getCode())) {
				    		if (o1.getSize() > o2.getSize())
				    			result = 1;
				    		else if (o1.getSize() < o2.getSize())
				    			result = -1;
				    		else
				    			result = 0;
				    	}
				    	else
				    		result = o1.getCode().compareTo(o2.getCode());
				    	
				    	return result;
				    }
				});
				
				// Create the Excel file
				XSSFWorkbook workbook = new XSSFWorkbook();
		        XSSFSheet sheet = workbook.createSheet("Quantities");
				CellStyle cellStyle = workbook.createCellStyle();
				CellStyle cellStyleCurrency = workbook.createCellStyle();
				CreationHelper createHelper = workbook.getCreationHelper();
				
				int numRowOther = 6;


				cellStyle.setAlignment(CellStyle.ALIGN_LEFT);
				cellStyleCurrency.setAlignment(CellStyle.ALIGN_LEFT);
				cellStyleCurrency.setDataFormat(createHelper.createDataFormat().getFormat("#,##0.00 €"));

		        sheet.setColumnWidth(0, 35 * 256);
		        sheet.setColumnWidth(1, 12 * 256);
		        sheet.setColumnWidth(4, 15 * 256);
		        sheet.setColumnWidth(numRowOther, 100 * 256);
				
		        Row row = sheet.createRow(0);
		        Cell cell = row.createCell(0);
		        cell.setCellType(Cell.CELL_TYPE_STRING);
		        cell.setCellValue("NAME");
		        cell = row.createCell(1);
		        cell.setCellType(Cell.CELL_TYPE_STRING);
		        cell.setCellValue("CODE");
		        cell = row.createCell(2);
		        cell.setCellType(Cell.CELL_TYPE_STRING);
		        cell.setCellValue("SIZE");
		        cell = row.createCell(3);
		        cell.setCellType(Cell.CELL_TYPE_STRING);
		        cell.setCellValue("QUANTITY");
		        cell = row.createCell(4);
		        cell.setCellType(Cell.CELL_TYPE_STRING);
		        cell.setCellValue("PRICE EXC. VAT");
		        
		        int line = 1;
				for (Output item : output) {
					row = sheet.createRow(line++);
			        cell = row.createCell(0);
			        cell.setCellType(Cell.CELL_TYPE_STRING);
			        cell.setCellValue(item.getName());
			        cell = row.createCell(1);
			        cell.setCellType(Cell.CELL_TYPE_STRING);
			        cell.setCellValue(item.getCode());
			        cell = row.createCell(2);
			        cell.setCellType(Cell.CELL_TYPE_NUMERIC);
			        cell.setCellValue(item.getSize());
			        cell = row.createCell(3);
			        cell.setCellType(Cell.CELL_TYPE_NUMERIC);
			        cell.setCellValue(item.getQuantity());
			        cell = row.createCell(4);
			        cell.setCellType(Cell.CELL_TYPE_NUMERIC);
					cell.setCellStyle(cellStyleCurrency);
			        cell.setCellValue(Double.parseDouble("" + item.getPriceExVAT()));
				}
				

				cellStyle.setAlignment(CellStyle.ALIGN_LEFT);
				// Total shoes
				line = 1;
				row = sheet.getRow(line++);
				cell = row.createCell(numRowOther);
		        cell.setCellType(Cell.CELL_TYPE_STRING);
		        cell.setCellValue("Total shoes");
				cell.setCellStyle(cellStyle);
	        	row = sheet.getRow(line++);
				cell = row.createCell(numRowOther);
				cell.setCellStyle(cellStyle);
		        cell.setCellType(Cell.CELL_TYPE_NUMERIC);
		        cell.setCellValue(totalShoes);
				
				// Total price ex VAT
				line += 1;
				row = sheet.getRow(line++);
				cell = row.createCell(numRowOther);
				cell.setCellStyle(cellStyle);
		        cell.setCellType(Cell.CELL_TYPE_STRING);
		        cell.setCellValue("Total price exc. VAT");
	        	row = sheet.getRow(line++);
				cell = row.createCell(numRowOther);
				cell.setCellType(Cell.CELL_TYPE_NUMERIC);
				cell.setCellStyle(cellStyleCurrency);
		        cell.setCellValue(totalPriceVAT.doubleValue());
				
				// Errors
				line += 2;
				row = sheet.getRow(line++);
				cell = row.createCell(numRowOther);
				cell.setCellStyle(cellStyle);
		        cell.setCellType(Cell.CELL_TYPE_STRING);
		        cell.setCellValue("Error messages");
		        if (errors.size() > 0) {
			        for (String error : errors) {
			        	row = sheet.getRow(line++);
						cell = row.createCell(numRowOther);
						cell.setCellStyle(cellStyle);
				        cell.setCellType(Cell.CELL_TYPE_STRING);
				        cell.setCellValue(error);
			        }
		        }
		        else {
		        	row = sheet.getRow(line++);
					cell = row.createCell(numRowOther);
					cell.setCellStyle(cellStyle);
			        cell.setCellType(Cell.CELL_TYPE_STRING);
			        cell.setCellValue(mRes.getString("noError"));
		        }
				
				// Warnings
				line += 2;
				row = sheet.getRow(line++);
				cell = row.createCell(numRowOther);
				cell.setCellStyle(cellStyle);
		        cell.setCellType(Cell.CELL_TYPE_STRING);
		        cell.setCellValue("Warning messages");
		        if (errors.size() > 0) {
			        for (String warning : warnings) {
			        	row = sheet.getRow(line++);
						cell = row.createCell(numRowOther);
						cell.setCellStyle(cellStyle);
				        cell.setCellType(Cell.CELL_TYPE_STRING);
				        cell.setCellValue(warning);
			        }
		        }
		        else {
		        	row = sheet.getRow(line++);
					cell = row.createCell(numRowOther);
					cell.setCellStyle(cellStyle);
			        cell.setCellType(Cell.CELL_TYPE_STRING);
			        cell.setCellValue(mRes.getString("noWarning"));
		        }
		        
				try {
					FileOutputStream os = new FileOutputStream(OUTPUT);
		            workbook.write(os);
		            
		            workbook.close();
		            os.close();
				} catch (Exception e) {
					e.printStackTrace();
				}
	           
	            // Show Error dialog
				if (errors.size() > 0) {
					JOptionPane.showMessageDialog(mFrame,
				    String.join("\r\n", errors),
				    "WARNING !!",
				    JOptionPane.ERROR_MESSAGE);
				}
				
				return null;
			}
			
			@Override
			protected void done() {
				super.done();
				
				JOptionPane.showMessageDialog(mFrame,
						mRes.getString("messageFileGenerated"),
					    "COMPLETED",
					    JOptionPane.PLAIN_MESSAGE);
			}
		};
		
		mWorker.addPropertyChangeListener(new PropertyChangeListener() {
			
			@Override
			public void propertyChange(PropertyChangeEvent evt) {
				if (evt.getPropertyName() == "progress") {
					mProgressBar.setValue((Integer) evt.getNewValue());
				}
			}
		});
		
		mWorker.execute();
	}
	
	public static void main(String[] args) {
		// take the menu bar off the jframe
		System.setProperty("apple.laf.useScreenMenuBar", "true");

		// set the name of the application menu item
		System.setProperty("com.apple.mrj.application.apple.menu.about.name", "OTZSHOES Excel");

		// set the look and feel
		try {
			UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
		} catch (Exception e) {
			e.printStackTrace();
		}

		new Main();
	}

}
