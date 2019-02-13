package main;

import java.awt.EventQueue;

import javax.swing.JFrame;
import java.awt.GridBagLayout;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
import java.awt.GridBagConstraints;
import javax.swing.JMenuBar;
import javax.swing.JMenu;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JFileChooser;
import java.awt.Insets;
import javax.annotation.Resource;
import javax.swing.JTextArea;
import java.awt.Color;
import javax.swing.JLabel;
import javax.swing.ImageIcon;
import java.awt.event.ActionListener;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Scanner;
import java.util.regex.Pattern;
import java.awt.event.ActionEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;

public class GUI {
	
	private static ArrayList<String> cities = new ArrayList<>();
	private JFrame frame;
	private static int blanks = 0;
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {	 
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					GUI window = new GUI();
					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public GUI() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame();
		frame.getContentPane().setBackground(Color.WHITE);
		frame.setBounds(100, 100, 765, 481);
		frame.setResizable(false);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		GridBagLayout gridBagLayout = new GridBagLayout();
		gridBagLayout.columnWidths = new int[]{0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0};
		gridBagLayout.rowHeights = new int[]{0, 0, 0, 0};
		gridBagLayout.columnWeights = new double[]{0.0, 0.0, 0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, Double.MIN_VALUE};
		gridBagLayout.rowWeights = new double[]{0.0, 1.0, 0.0, Double.MIN_VALUE};
		frame.getContentPane().setLayout(gridBagLayout);
		
		JLabel lblNewLabel = new JLabel("");
		lblNewLabel.addMouseListener(new MouseAdapter() {
			@Override
			public void mousePressed(MouseEvent e) {
				JOptionPane.showMessageDialog(null, "Did you know Adam crashed Fernando's bike the first day they met?");
			}
		});
		lblNewLabel.setIcon(new ImageIcon(GUI.class.getResource("/resource/MuleGUILogo.png")));
		GridBagConstraints gbc_lblNewLabel = new GridBagConstraints();
		gbc_lblNewLabel.anchor = GridBagConstraints.NORTH;
		gbc_lblNewLabel.insets = new Insets(0, 0, 5, 5);
		gbc_lblNewLabel.gridx = 1;
		gbc_lblNewLabel.gridy = 0;
		frame.getContentPane().add(lblNewLabel, gbc_lblNewLabel);
		
		JFileChooser fileChooser = new JFileChooser();
		fileChooser.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				if(arg0.getActionCommand().equals(javax.swing.JFileChooser.APPROVE_SELECTION))
				{
					String[] options = {"Yes", "No"};
					int n = JOptionPane.showOptionDialog(null,
							"The information will be extracted to an Excel Sheet, would you like to continue?",
							"Continue",
							JOptionPane.DEFAULT_OPTION,
							JOptionPane.QUESTION_MESSAGE,
							null,
							options,
							options[0]);
					if(n == 0)
					{
						try {
							exportToExcel(fileChooser.getSelectedFile());
							JOptionPane.showMessageDialog(null, "Complete!\nPlease look for file named ExtractedInfo.xls");
							System.out.println(blanks);
						} catch (FileNotFoundException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						} catch (IOException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						} catch (WriteException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
					}
					else if(n == 1)
					{
						JOptionPane.getRootFrame().dispose();
					}
					
				}
				else
				{
					JOptionPane.showMessageDialog(null, "No file chosen");
				}
			}
		});
		GridBagConstraints gbc_fileChooser = new GridBagConstraints();
		gbc_fileChooser.gridwidth = 7;
		gbc_fileChooser.insets = new Insets(0, 0, 5, 5);
		gbc_fileChooser.gridx = 6;
		gbc_fileChooser.gridy = 0;
		frame.getContentPane().add(fileChooser, gbc_fileChooser);
		
		JTextArea txtrWelcomeToMule = new JTextArea();
		txtrWelcomeToMule.setEditable(false);
		txtrWelcomeToMule.setLineWrap(true);
		txtrWelcomeToMule.setBackground(Color.WHITE);
		txtrWelcomeToMule.setText("Welcome to Mule001!\r\n\r\nPlease open the file using the file browser above.\r\nThen choose the correct output format and enjoy!");
		GridBagConstraints gbc_txtrWelcomeToMule = new GridBagConstraints();
		gbc_txtrWelcomeToMule.gridwidth = 12;
		gbc_txtrWelcomeToMule.insets = new Insets(0, 0, 5, 0);
		gbc_txtrWelcomeToMule.fill = GridBagConstraints.HORIZONTAL;
		gbc_txtrWelcomeToMule.gridx = 5;
		gbc_txtrWelcomeToMule.gridy = 1;
		frame.getContentPane().add(txtrWelcomeToMule, gbc_txtrWelcomeToMule);
		
		JMenuBar menuBar = new JMenuBar();
		frame.setJMenuBar(menuBar);
		
		JMenu mnNewMenu = new JMenu("Menu");
		menuBar.add(mnNewMenu);
		
		JMenuItem mntmNewMenuItem = new JMenuItem("Exit");
		mntmNewMenuItem.addMouseListener(new MouseAdapter() {
			@Override
			public void mousePressed(MouseEvent e) {
				System.exit(0);
			}
		});
		mnNewMenu.add(mntmNewMenuItem);
	}

	
	private void exportToExcel(File file) throws FileNotFoundException, IOException, WriteException 
	{
		isCityFL();
		interpretFile(file);
	}
	
	private void interpretFile(File file) throws FileNotFoundException, IOException, WriteException 
	{
		int tempInt = 0;
		Information infoClass = new Information();
		StringBuilder addy = new StringBuilder("");
		StringBuilder cNameStr = new StringBuilder("");
		StringBuilder addy2 = new StringBuilder("");
		StringBuilder nameStr = new StringBuilder("");
		StringBuilder addy3 = new StringBuilder("");
		StringBuilder nameStr2 = new StringBuilder("");
		
		// create/open an excel file
		int givenRow = 0;
        WritableWorkbook wworkbook;
        wworkbook = Workbook.createWorkbook(new File(file.getName() + ".xls"));
        
        //Sheet name
        WritableSheet wsheet = wworkbook.createSheet("Mailing_Data", 0);
        
        // column names
        Label label;
        label = new Label(0, 0, "Company");
        wsheet.addCell(label);
        label = new Label(1, 0, "Address Line 1");
        wsheet.addCell(label);
        label = new Label(2, 0, "City");
        wsheet.addCell(label);
        label = new Label(3, 0, "State");
        wsheet.addCell(label);
        label = new Label(4, 0, "Zip");
        wsheet.addCell(label);

        
        // set width of each cell to 150 chars
        for(int i = 0; i < 7; i++)
        {
        	wsheet.setColumnView(i, 100);
        }

		
		// create patterns for necessary information
		Pattern cName = Pattern.compile("[FLMNP]+[0-9]+[A-Z]+");
		Pattern date = Pattern.compile("[US]*[0-9]{4}[20][0-9]+");
		

		try (BufferedReader br = new BufferedReader(new FileReader(file))) {
		    String line;
		    String[] words;	
		    // read file
		    while ((line = br.readLine()) != null) {
		    	words = line.split("\\s+");
	    		if(cName.matcher(words[0]).find() && words[0].length() > 11)
	    		{
	    			// search for company name
	    			cNameStr.append(words[0].substring(12));
	    			tempInt = findAddressIndex(words);
	    			for(int i = 1; i < tempInt; i++)
	    			{
	    				cNameStr.append(" "+ words[i]);
	    			}
	    			
	    			// search for address
	    			for(int i = tempInt; i < findEndOfAddressIndex(words); i++)
	    			{
	    				addy.append(" "+ words[i]);
	    			}
	    			
	    			
	    			tempInt = findEndOfAddressIndex(words);
	    			
	    			// search second address, ignoring random dates
	    			for(int i = tempInt; i < words.length; i++)
	    			{
	    				if(!date.matcher(words[i]).find())
	    				{
	    					addy2.append(" " + words[i]);
	    					tempInt = i + 1;
	    				}
	    				else
	    				{
	    					tempInt = i + 1;
	    					break;
	    				}
	    			}
	    			
	    			// search for name, ignoring random strings
	    			for(int i = tempInt; i < words.length; i++)
	    			{
	    				if(!words[i].matches("[A-Z]{1}[0-9]+"))
	    				{
	    					nameStr.append(words[i] + " ");
	    				}
	    				else
	    				{
	    					tempInt = i + 1;
	    					break;
	    				}
	    			}
	    			
	    			// search for another address, ignoring random strings
	    			for(int i = tempInt; i < words.length; i++)
	    			{
	    				if(!words[i].matches("([A-Z]{2})*3[0-9]{4}"))
	    				{
	    					addy3.append(words[i] + " ");
	    				}
	    				else
	    				{
	    					addy3.append(words[i] + " ");
	    					tempInt = i + 1;
	    					break;
	    				}
	    			}
	    			
	    			// search for second name, ignoring random values
	    			for(int i = tempInt; i < words.length; i++)
	    			{
	    				if(!words[i].matches("[0-9]{2,}"))
	    				{
	    					nameStr2.append(words[i] + " ");
	    					tempInt = i + 1;
	    				}
	    				else
	    					break;
	    			}
	    			
	    			// store any uninterpreted information
	    			String rest = "";
	    			for(int i = tempInt; i < words.length; i++)
	    			{
	    				rest += words[i] + " ";
	    			}
	    			
	    			// fix name
	    			if(cNameStr.toString().contains("AFLAL") || cNameStr.toString().contains("ADOMP"))
	    			{
	    				cNameStr.delete(cNameStr.length() - 6, cNameStr.length());
	    			}
	    			
	    			// add information to its class
	    			infoClass.setAddress(addy.toString());
	    			infoClass.setAddress2(addy2.toString());
	    			infoClass.setAddress3(addy3.toString());
	    			infoClass.setCompName(cNameStr.toString());
	    			infoClass.setName(nameStr.toString().replaceAll("N*\\s+[A-Z]{2}\\s+", " ").replaceAll("FL\\s+", " "));
	    			infoClass.setCompName2(nameStr2.toString());
	    			infoClass.setOther(rest);
	    			
	    			// write to the excel file
	                Write_to_excel_file_directly(infoClass, ++givenRow, wsheet, wworkbook);
	                
	                // clear the strings
	    	    	cNameStr.delete(0,cNameStr.length());
	    	    	addy.delete(0, addy.length());
	    	    	addy2.delete(0, addy2.length());
	    	    	nameStr.delete(0, nameStr.length());
	    	    	addy3.delete(0, addy3.length());
	    	    	nameStr2.delete(0, nameStr2.length());		
	    		}
	    	}
	    }
		
        wworkbook.write();
        wworkbook.close();
	}

	private void Write_to_excel_file_directly(Information infoClass, int i, WritableSheet wsheet,
			WritableWorkbook wworkbook) throws RowsExceededException, WriteException 
	{
		// these strings should probably be changed to 
		// StringBuilder because strings are immutable
		String temp = "";
		String last = "";
		String[] name;
		String otherTemp = infoClass.getName().trim();
		String[] addy;
		String tempAdd = "";
		String zip = "";
		String city = "";
		String state = "Florida";
		String compName = "";
		String[] comp;
		
		// divide name by spaces and interpret necessary portions
		name = otherTemp.split("\\s+");
		
		if(name.length > 2)
		{
			last = name[0];
			for(int j = 1; j < name.length; j++)
			{
				temp += name[j] + " ";
			}
		}
		else if(name.length == 2)
		{
			temp = name[1];
			last = name[0];
		}
		else
			temp = name[0];
		
		// special case(s)
		if(last.contains("N01182019"))
		{
			temp = name[2];
			last = name[1];
		}
		
		// adjust address 
		if(infoClass.getAddress().length() < 2)
		{
			addy = infoClass.getAddress2().split("\\s+");
		}
		else
		{
			// get parts of address
			addy = infoClass.getAddress().split("\\s+");
		}
		
		// find zip code
		for(int j = 0; j < addy.length; j++)
		{
			if(addy[j].matches("F*L*3[0-9]{4}"))
			{
				zip = addy[j];
				
				if(zip.contains("FL"))
				{
					zip = zip.replace("FL", "");
				}
				
				tempAdd = infoClass.getAddress().replaceAll(addy[j], "");
				
			}
			
		}
		
		// make sure it is a valid Florida city
		for(String str : cities)
		{
			if(tempAdd.contains(str.toUpperCase()))
			{
				city = str.toUpperCase();
				addy = tempAdd.split("\\s+");
				tempAdd = "";
				for(int j = 0; j < addy.length - city.split("\\s+").length; j++)
				{
					tempAdd += addy[j] + " ";
				}
				
				if(city.contains("HOLLYWOOOD"))
				{
					city = "HOLLYWOOD";
				}
				break;
			}
		}
		
		// if we have no city, we can't assume its florida
		if(city == "")
		{
			state = "";
		}
		
		// edit company name
		comp = infoClass.getCompName().split("\\s+");
		
		// remove unnecessary strings
		for(int k = 0; k < comp.length; k++)
		{
			if(comp[k].contains("AFLAL") || comp[k].contains("ADOMNP") 
					|| comp[k].contains("AFORL") || comp[k].contains("AFORP"))
			{
				break;
			}
			else
			{
				compName += comp[k];
				if( k == comp.length - 1)
				{
					break;
				}
				else
				{
					compName += " ";
				}
				
			}
		}
		
		// keep track of misinterpreted information
		if(tempAdd.length() < 2)
		{
			blanks++;
		}
		
		// add information to excel sheet
		Label label;
		
        // Company Name
        label = new Label(0, i, compName);
        wsheet.addCell(label);
        // Address Line 1 
        label = new Label(1, i, tempAdd);
        wsheet.addCell(label);
        // City
        label = new Label(2, i, city);
        wsheet.addCell(label);
        // State
        label = new Label(3, i, state);
        wsheet.addCell(label);
        // Zip
        label = new Label(4, i, zip);
        wsheet.addCell(label);
		
	}

	// finds the index in which the address begins
	private int findAddressIndex(String[] words) {
		int retVal = 0;
		
		for(int i = 0; i < words.length; i++)
		{
			if(words[i].matches("[0-9]+") && !words[i+1].contains("LLC")
					&& !words[i+2].contains("LLC") && !words[i+3].contains("LLC"))
			{
				retVal = i;
				return retVal;
			}
		}
		
		return retVal;
	}
	
	// finds the index in which the address ends
	private int findEndOfAddressIndex(String[] words)
	{
		int retVal = 0;
		
		for(int i = 0; i < words.length; i++)
		{
			if(words[i].matches("[FL]*[3]{1}[0-9]{4}"))
			{
				retVal = i + 1;
				return retVal;
			}
		}
		
		return retVal;
	}

	// a class that holds all interpreted information
	private class Information
	{
		private String address;
		private String address2;
		private String address3;
		private String name;
		private String compName;
		private String compName2;
		private String other;
		
		public String getAddress() {
			return address;
		}
		public void setAddress(String address) {
			this.address = address;
		}
		public String getAddress2() {
			return address2;
		}
		public void setAddress2(String address2) {
			this.address2 = address2;
		}
		public String getAddress3() {
			return address3;
		}
		public void setAddress3(String address3) {
			this.address3 = address3;
		}
		public String getName() {
			return name;
		}
		public void setName(String name) {
			this.name = name;
		}
		public String getCompName() {
			return compName;
		}
		public void setCompName(String compName) {
			this.compName = compName;
		}
		public String getOther() {
			return other;
		}
		public void setOther(String other) {
			this.other = other;
		}
		public String getCompName2() {
			return compName2;
		}
		public void setCompName2(String compName2) {
			this.compName2 = compName2;
		}
	
	}
	
	// reads a text file that includes all the florida cities and adds them to a list 
	public static void isCityFL() throws FileNotFoundException
	{
		
		InputStream is = Resource.class.getResourceAsStream("/FloridaCities.txt");
		if(is == null)
		{
			return;
		}
		
		Scanner sc = new Scanner(is);
		while(sc.hasNextLine())
		{
			cities.add(sc.nextLine());
		}
		
		sc.close();

	}
}
