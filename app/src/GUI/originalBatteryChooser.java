package GUI;

import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;

import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import javax.swing.JTextField;
import javax.swing.JButton;
import javax.swing.JFileChooser;

public class originalBatteryChooser {

	JFrame frame;
	private JTextField textField;
	public static String originalBatteryFilePath = null;
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					originalBatteryChooser window = new originalBatteryChooser();
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
	public originalBatteryChooser() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame(MainFrame.titleName);
		frame.setBounds(100, 100, 763, 286);
		frame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		frame.getContentPane().setLayout(null);
		
		JLabel lblPleaseSelectThe = new JLabel("Please select the Original Battery file:");
		lblPleaseSelectThe.setFont(new Font("Tahoma", Font.PLAIN, 18));
		lblPleaseSelectThe.setBounds(66, 54, 350, 32);
		frame.getContentPane().add(lblPleaseSelectThe);
		
		textField = new JTextField();
		textField.setBounds(66, 92, 513, 32);
		frame.getContentPane().add(textField);
		textField.setColumns(10);
		
		JButton btnChoose = new JButton("Choose...");
		btnChoose.setBounds(594, 94, 115, 29);
		btnChoose.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser jfc=new JFileChooser();	
				jfc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);	//parameter for both files and directories
				jfc.setFileFilter(new javax.swing.filechooser.FileFilter() {	//create a file filter to select only ".xlsx" file
					public boolean accept(File f) {
						if(f.getName().endsWith(".xlsx") || f.isDirectory()) {
							return true;
						}
						return false;
					}
					public String getDescription() {
						return "Original Battery File(*.xlsx)";
					}
				});
				jfc.showDialog(new JLabel(), "Choose Original Battery File");
				if(jfc.getSelectedFile() != null) {		//choose procedure
					File file=jfc.getSelectedFile();	
					textField.setText(file.getAbsolutePath());
					originalBatteryFilePath = file.getAbsolutePath();
				}
			}
		});
		frame.getContentPane().add(btnChoose);
		
		JButton btnNext = new JButton("Next");
		btnNext.setBounds(594, 154, 115, 29);
		btnNext.addActionListener(new ActionListener() {	//add action
			public void actionPerformed(ActionEvent e) {
				EventQueue.invokeLater(new Runnable() {
					public void run() {
						try {
							formattedBatteryChooser window = new formattedBatteryChooser();
							window.frame.setVisible(true);
						} catch (Exception e) {
							JOptionPane.showMessageDialog(frame, "Failed!", "Warning!", JOptionPane.WARNING_MESSAGE);
							e.printStackTrace();
						} finally {
							frame.dispose();	//close the frame
						}
					}
				});
			}
		});
		frame.getContentPane().add(btnNext);
	}
}
