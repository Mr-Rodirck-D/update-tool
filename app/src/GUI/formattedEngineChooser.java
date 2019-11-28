package GUI;

import java.awt.EventQueue;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.SwingWorker;

import projectCode.updateEngine;

public class formattedEngineChooser {

	JFrame frame;
	private JTextField textField;
	public static String formattedEngineFilePath = null;
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					formattedEngineChooser window = new formattedEngineChooser();
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
	public formattedEngineChooser() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		processingFrame processingframe = new processingFrame();
		frame = new JFrame(MainFrame.titleName);
		frame.setBounds(100, 100, 763, 286);
		frame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		frame.getContentPane().setLayout(null);
		
		JLabel lblPleaseSelectThe = new JLabel("Please select the Formatted Engine file:");
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
				String DefaultPath = null;
				try {
					DefaultPath = formattedEngineChooser.class.getProtectionDomain().getCodeSource().getLocation().toURI().getPath();
					DefaultPath = java.net.URLDecoder.decode(DefaultPath, "utf-8");
					DefaultPath = DefaultPath.replace("%20", " ");
					DefaultPath = DefaultPath.substring(1, DefaultPath.lastIndexOf("/"));
				} catch (Exception e1) {
					JOptionPane.showMessageDialog(frame, "GUIFailed!\n Please check the existency of directory Engine!", "Warning!", JOptionPane.WARNING_MESSAGE);
					e1.printStackTrace();
				}
				JFileChooser jfc=new JFileChooser();
				jfc.setCurrentDirectory(new File(DefaultPath + "/Engine"));
				jfc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);	//parameter for both files and directories
				jfc.setFileFilter(new javax.swing.filechooser.FileFilter() {	//create a file filter to select only ".xlsx" file
					public boolean accept(File f) {
						if(f.getName().endsWith(".xlsx") || f.isDirectory()) {
							return true;
						}
						return false;
					}
					public String getDescription() {
						return "Formatted Engine File(*.xlsx)";
					}
				});
				jfc.showDialog(new JLabel(), "Choose Formatted Engine File");
				if(jfc.getSelectedFile() != null) {		//choose procedure
					File file=jfc.getSelectedFile();	
					textField.setText(file.getAbsolutePath());
					formattedEngineFilePath = file.getAbsolutePath();
				}
			}
		});
		frame.getContentPane().add(btnChoose);
		
		JButton btnStart = new JButton("Start");
		btnStart.setBounds(594, 154, 115, 29);
		frame.getContentPane().add(btnStart);
		btnStart.addActionListener(new ActionListener() {	//action
			public void actionPerformed(ActionEvent e) {
				EventQueue.invokeLater(new Runnable() {
					public void run() {
						try {
							processingframe.setVisible(true);
						} catch (Exception e2) {
							JOptionPane.showMessageDialog(frame, "GUIFailed!", "Warning!", JOptionPane.WARNING_MESSAGE);
							e2.printStackTrace();
						}
					}
				});
			}			
		});
		btnStart.addActionListener(new ActionListener() {	//action
			public void actionPerformed(ActionEvent e) {			
				SwingWorker<Void, Void> worker = new SwingWorker<Void, Void>(){
					protected Void doInBackground() throws Exception{
						try {
							String outputEngineFileNamePathandName = updateEngine.updateEngine(originalEngineChooser.originalEngingeFilePath, formattedEngineFilePath);
							processingframe.dispose();
							JOptionPane.showMessageDialog(frame, "Update succeed!\n" + (updateEngine.addRows + updateEngine.deleteRows) + " rows have been updated!\n" + updateEngine.addRows + " rows have been added!\n" + updateEngine.deleteRows + " rows have been deleted!", "Congratulation!", JOptionPane.INFORMATION_MESSAGE);	//message box 
							frame.dispose();
							
							//delete original file
							int isDelete = JOptionPane.showConfirmDialog(frame, "Are you going to delete the original file?", "Attention!!", JOptionPane.YES_NO_CANCEL_OPTION);	//confirm window for user to choose
							if(isDelete == JOptionPane.YES_OPTION) {	//if confirm
								try {	//delete and show message pane
									File originalFile = new File(originalEngineChooser.originalEngingeFilePath);
									originalFile.delete();
									JOptionPane.showMessageDialog(frame, "File has been deleted", "Congratulation!", JOptionPane.INFORMATION_MESSAGE);
								} catch (Exception e1) {
									JOptionPane.showMessageDialog(frame, "Delete Failed!", "Warning!", JOptionPane.WARNING_MESSAGE);
									e1.printStackTrace();
								}
							}
							
							//open created file
							int isOpen = JOptionPane.showConfirmDialog(frame, "Are you going to open the created file?", "Attention!!", JOptionPane.YES_NO_CANCEL_OPTION);	//confirm window for user to choose
							if(isOpen == JOptionPane.YES_OPTION) {	//if confirm
								try {	
									String cmd = "rundll32 url.dll FileProtocolHandler " +  outputEngineFileNamePathandName;
									Runtime.getRuntime().exec(cmd);	//open Access in work place 
								} catch (Exception e1) {
									JOptionPane.showMessageDialog(frame, "Open Failed!", "Warning!", JOptionPane.WARNING_MESSAGE);
									e1.printStackTrace();
								}
							}
						} catch (Exception e1) {
							processingframe.dispose();
							JOptionPane.showMessageDialog(frame, "Failed!" + e1.toString(), "Warning!", JOptionPane.WARNING_MESSAGE);
							e1.printStackTrace();
						} finally {
							frame.dispose();
						}
						return null;
					}
				};
				worker.execute();
			}
		});
	}

}
