package GUI;

import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JOptionPane;

import javax.swing.JButton;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

public class MainFrame {

	private JFrame frame;
	public static String titleName = "Update Tracking APP_version 2.1";
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					MainFrame window = new MainFrame();
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
	public MainFrame() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame(titleName);
		frame.setBounds(100, 100, 599, 337);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);
		
		JButton btnNewButton = new JButton("Update Engine Tracking Form");
		btnNewButton.setFont(new Font("Tahoma", Font.BOLD, 25));
		btnNewButton.setBounds(77, 38, 422, 88);
		btnNewButton.addActionListener(new ActionListener() {	//add action
			public void actionPerformed(ActionEvent e) {
				EventQueue.invokeLater(new Runnable() {
					public void run() {
						try {
							originalEngineChooser window = new originalEngineChooser();
							window.frame.setVisible(true);
						} catch (Exception e) {
							JOptionPane.showMessageDialog(frame, "Failed!", "Warning!", JOptionPane.WARNING_MESSAGE);
							e.printStackTrace();
						} 
//						finally {
//							frame.dispose();	//close the frame
//						}
					}
				});
			}
		});
		frame.getContentPane().add(btnNewButton);
		
		JButton btnUpdateBatteryTracking = new JButton("Update Battery Tracking Form");
		btnUpdateBatteryTracking.setFont(new Font("Tahoma", Font.BOLD, 25));
		btnUpdateBatteryTracking.setBounds(77, 157, 422, 88);
		btnUpdateBatteryTracking.addActionListener(new ActionListener() {	//add action
			public void actionPerformed(ActionEvent e) {
				EventQueue.invokeLater(new Runnable() {
					public void run() {
						try {
							originalBatteryChooser window = new originalBatteryChooser();
							window.frame.setVisible(true);
						} catch (Exception e) {
							JOptionPane.showMessageDialog(frame, "Failed!", "Warning!", JOptionPane.WARNING_MESSAGE);
							e.printStackTrace();
						} 
//						finally {
//							frame.dispose();	//close the frame
//						}
					}
				});
			}
		});
		frame.getContentPane().add(btnUpdateBatteryTracking);
	}
}
