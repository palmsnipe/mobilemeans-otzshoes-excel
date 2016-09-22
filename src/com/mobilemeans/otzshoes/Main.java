package com.mobilemeans.otzshoes;

import java.awt.BorderLayout;
import java.awt.dnd.DropTarget;
import java.beans.PropertyChangeEvent;
import java.beans.PropertyChangeListener;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
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
				List<Output> output = new ArrayList<>();
				
				int i = 0;
				
				File[] files = folder.listFiles();
				float divisor = (float) files.length / 100.0f;
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
			        				errors.add(String.format("The code \"%s\" in \"%s\" is incorrect (not included in the results)", shoe.getCode(), fileEntry.getName()));
			        			}
			        			else {
				        			for (Stock size : shoe.getStock()) {
				        				boolean added = false;
				        				for (Output item : output) {
				        					if (item.getCode().equals(shoe.getCode()) && item.getSize() == size.getSize()) {
				        						item.setQuantity(item.getQuantity() + size.getQuantity());
				        						
				        						if (item.getName().length() < shoe.getName().length())
				        							item.setName(shoe.getName());
				        						
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
				        					
				        					output.add(newOutput);
				        				}
				        				
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
				}
				
				// Errors
				line = 1;
				row = sheet.getRow(line++);
				cell = row.createCell(5);
		        cell.setCellType(Cell.CELL_TYPE_STRING);
		        cell.setCellValue("Error messages");
		        if (errors.size() > 0) {
			        for (String error : errors) {
			        	row = sheet.getRow(line++);
						cell = row.createCell(5);
				        cell.setCellType(Cell.CELL_TYPE_STRING);
				        cell.setCellValue(error);
			        }
		        }
		        else {
		        	row = sheet.getRow(line++);
					cell = row.createCell(5);
			        cell.setCellType(Cell.CELL_TYPE_STRING);
			        cell.setCellValue(mRes.getString("noError"));
		        }
		        
				try {
					FileOutputStream os = new FileOutputStream(OUTPUT);
		            workbook.write(os);
		            
		            workbook.close();
		            os.close();
				} catch (FileNotFoundException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (Exception e) {
					// TODO Auto-generated catch block
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
				
//				mProgressBar.setIndeterminate(false);
				
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
