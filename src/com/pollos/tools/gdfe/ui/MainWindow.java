package com.pollos.tools.gdfe.ui;

import java.awt.Dimension;
import java.awt.Font;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.UIManager;

import com.pollos.tools.gdfe.util.ExcelUtil;
import com.pollos.tools.gdfe.util.StringUtil;

public class MainWindow extends JFrame {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	
	private static final Dimension screenSize =Toolkit.getDefaultToolkit().getScreenSize();
	private static final int xSize = 600;
	private static final int ySize = 150;
	private static final int x = screenSize.width/2 - xSize/2;
	private static final int y = screenSize.height/2 - ySize/2;
	private static int leftPadding = 10;
	private static int topPadding = 20;
	private static int linePadding = 10;
	private static int labelWidth = 300;
	private static int labelHeight = 20;
	private static int inLinePadding = 5;
	private static int textWidth = 180;
	private static int largeTextWidth = 200;
	private static int textHeight = 20;
	private static int buttonWidth = 40;
	private static int largeButtonWidth = 100;
	private static int buttonHeight = 20;
	private static String pathLabelText = "请输入要处理的文件夹路径:";
	private static String pathButtonText = "...";
	private static String cellLabelText = "请输入要抓取的单元格（逗号分隔，如：A1,A2）：";
	private static String pathChooserTitle = "请选择文件夹：";
	private static String title = "Grab Data From Excel";
	private static String grabButtonText = "Run Now!";
	private static String grabButtonText2 = "Running...";
	private static String grabButtonText3 = "Done!";
	private static String grabButtonText4 = "Failed!";
	private boolean running = false;
	
	private int walkX = 0;
	private int walkY = 0;
	private JPanel pane = null;
	private JLabel label = null;
	private JTextField pathText = null;
	private JButton pathButton = null;
	private JTextField cellText = null;
	private JButton grabButton = null;
	
	public MainWindow(){
		try {
			UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
		} catch (Exception e) {
			e.printStackTrace();
		}
		setLocation(x, y);
		setSize(xSize, ySize);
		setTitle(title);
		setDefaultLookAndFeelDecorated(true);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		pane = new JPanel();
		add(pane);
		createPaneComponments(pane);
		setVisible(true);
	}
	private void createPaneComponments(JPanel pane) {
		pane.setLayout(null);

		walkX += leftPadding;
		walkY += topPadding;
		
		addPathLabel();
		walkX += inLinePadding;
		addPathText();
		walkX += inLinePadding;
		addPathSelect();

		walkX = leftPadding;
		walkY += labelHeight + linePadding;
		addCellStrLabel();
		walkX += inLinePadding;
		addCellStrText();
		walkX = leftPadding;
		walkY += labelHeight + linePadding;
		addGrabButton();
	}
	
	private void resetGrabButton(){
		if(grabButton != null && !grabButtonText.equals(grabButton.getText())){
			grabButton.setText(grabButtonText);
		}
	}
	private void addGrabButton() {
		grabButton = new JButton(grabButtonText);
		grabButton.setFont(new Font(Font.DIALOG, Font.ITALIC, 10));
		walkX = xSize/2 - largeButtonWidth/2;
		grabButton.setBounds(walkX, walkY, largeButtonWidth, buttonHeight);
        pane.add(grabButton);
        walkX += buttonWidth;
        grabButton.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				new Thread(new Runnable() {
					@Override
					public void run() {
						String path = pathText.getText();
						String cellPositionStr = cellText.getText();
						if(running || StringUtil.isNullOrEmpty(path) || StringUtil.isNullOrEmpty(cellPositionStr)){
							return;
						}
						try {
							grabButton.setText(grabButtonText2);
							running = true;
							ExcelUtil.grab(path, cellPositionStr);
							grabButton.setText(grabButtonText3);
						} catch (Exception e1){
							JOptionPane.showMessageDialog(null, e1.getMessage() + "\n抓取数据失败，请保证输入参数正确后重试！", "错误提示",JOptionPane.ERROR_MESSAGE);
							grabButton.setText(grabButtonText4);
							e1.printStackTrace();
						} finally{
							running = false;
						}
					}
				}).start();;
			}
		});		
	}
	private void addCellStrText() {
		cellText = new JTextField();
		cellText.setBounds(walkX, walkY, largeTextWidth, textHeight);
        pane.add(cellText);
        walkX += textWidth;		
        
//        pathText.getDocument().addDocumentListener(new DocumentListener() {
//			
//			@Override
//			public void removeUpdate(DocumentEvent e) {
//				
//			}
//			
//			@Override
//			public void insertUpdate(DocumentEvent e) {
//				
//			}
//			
//			@Override
//			public void changedUpdate(DocumentEvent e) {
//				resetGrabButton();
//			}
//		});
	}
	private void addCellStrLabel() {
		label = new JLabel(cellLabelText);
		label.setBounds(walkX, walkY, labelWidth , labelHeight);
		pane.add(label);
		walkX += labelWidth;	
	}
	private void addPathSelect() {
		pathButton = new JButton(pathButtonText);
		pathButton.setFont(new Font(Font.DIALOG, Font.PLAIN, 10));
		pathButton.setBounds(walkX, walkY, buttonWidth, buttonHeight);
        pane.add(pathButton);
        walkX += buttonWidth;
        
        pathButton.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				JFileChooser jfc = new JFileChooser();
				jfc.setDialogTitle(pathChooserTitle);
				jfc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				if(jfc.showOpenDialog(pane) == JFileChooser.APPROVE_OPTION){
					pathText.setText(jfc.getSelectedFile().getAbsolutePath());
					resetGrabButton();
				}
			}
		});		
	}
	private void addPathText() {
		pathText = new JTextField();
		pathText.setBounds(walkX, walkY, textWidth, textHeight);
        pane.add(pathText);
        walkX += textWidth;
        
//        pathText.getDocument().addDocumentListener(new DocumentListener() {
//			
//			@Override
//			public void removeUpdate(DocumentEvent e) {
//				
//			}
//			
//			@Override
//			public void insertUpdate(DocumentEvent e) {
//				
//			}
//			
//			@Override
//			public void changedUpdate(DocumentEvent e) {
//				resetGrabButton();
//			}
//		});
	}
	private void addPathLabel() {
		//label
		label = new JLabel(pathLabelText);
		label.setBounds(walkX, walkY, labelWidth, labelHeight);
		pane.add(label);
		walkX += labelWidth;		
	}
}
