package com.capgemini.RenameFolders;

import java.awt.Color;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JTextField;

public class DirectoryPage implements ActionListener {
	
	private static final String APP_NAME = "Application-Listing.xlsx";
	private static final String APP_ITEM_NAME = "KAPView.xlsx";
	
	JFrame frame = new JFrame();
	JButton enterButton = new JButton("Enter");
	JTextField dirField = new JTextField();
	JLabel dirLabel = new JLabel("Project Directory : ");
	JLabel messageLabel = new JLabel("DTP Folder Rename");
	JLabel logLabel = new JLabel("Folder Renaming completed. Check your folders.");
	
	private Directory directory;
	
	public DirectoryPage(Directory directory) {
		this.directory = directory;
				
		messageLabel.setBounds(140, 20, 180, 50);
		messageLabel.setForeground(Color.WHITE);
		dirLabel.setBounds(40, 80, 150, 25);
		dirLabel.setForeground(Color.WHITE);
		dirField.setBounds(160, 80, 200, 25);
		enterButton.setBounds(160, 120, 75, 25);
		enterButton.addActionListener(this);
		enterButton.setFocusable(false);
		logLabel.setBounds(60, 160, 360, 50);
		logLabel.setForeground(Color.WHITE);
		logLabel.setVisible(false);
		
		frame.add(dirLabel);
		frame.add(messageLabel);
		frame.add(dirField);
		frame.add(enterButton);
		frame.add(logLabel);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.setSize(420, 280);
		frame.setLayout(null);
		frame.setVisible(true);
		frame.getContentPane().setBackground(Color.DARK_GRAY);

	}

	@Override
	public void actionPerformed(ActionEvent e) {
		
		if(e.getSource() == enterButton) {
			Directory directory = new Directory();
			directory.setDirectory(dirField.getText());
			
			String dirName = directory.getDirectory();
			
			ReadExcel.appFolderRenaming(dirName + "\\", APP_NAME, APP_ITEM_NAME);
			
			logLabel.setVisible(true);
		}
		
	}
	
}
