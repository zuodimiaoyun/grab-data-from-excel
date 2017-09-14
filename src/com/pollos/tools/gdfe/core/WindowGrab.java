package com.pollos.tools.gdfe.core;

import java.io.File;
import java.io.PrintStream;

import javax.swing.JOptionPane;

import com.pollos.tools.gdfe.ui.MainWindow;

public class WindowGrab {
	private static boolean relocateLog = true;
	private static String logFile = "a.log";
	public static void main(String[] args){
		relocateLog();
		new MainWindow();
	}
	private static void relocateLog(){
		if(relocateLog){
			try{
				File file = new File(logFile);
				if(!file.exists()){
					file.createNewFile();
				}
				PrintStream p = new PrintStream(file);
				System.setErr(p);
				System.setOut(p);
			}catch(Exception e){
				JOptionPane.showMessageDialog(null, e.getMessage() + "\n启动错误，请联系作者", "错误提示",JOptionPane.ERROR_MESSAGE);
				e.printStackTrace();
			}
		}
	}
}
