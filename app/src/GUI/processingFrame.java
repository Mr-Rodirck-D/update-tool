package GUI;

import java.awt.BorderLayout;
import java.awt.EventQueue;
import java.awt.Font;

import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.SwingConstants;
import javax.swing.border.EmptyBorder;

public class processingFrame extends JFrame {

	private JPanel contentPane;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					processingFrame frame = new processingFrame();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public processingFrame() {
		setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		setBounds(100, 100, 495, 192);
		JPanel contentPane = new JPanel();
		this.setTitle("Processing...");
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		contentPane.setLayout(new BorderLayout(0, 0));
		setContentPane(contentPane);
		
		JLabel lblProcessingPleaseWait = new JLabel("Processing... Please wait...");
		lblProcessingPleaseWait.setFont(new Font("Tahoma", Font.PLAIN, 35));
		lblProcessingPleaseWait.setHorizontalAlignment(SwingConstants.CENTER);
		contentPane.add(lblProcessingPleaseWait, BorderLayout.CENTER);
	}

}
