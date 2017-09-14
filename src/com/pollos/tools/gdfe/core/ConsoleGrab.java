package com.pollos.tools.gdfe.core;

import java.util.Scanner;

import com.pollos.tools.gdfe.util.ExcelUtil;


public class ConsoleGrab {

	public static void main(String[] arg){
		Scanner scanner = new Scanner(System.in);
		try{
			System.out.println("请输入要处理的文件夹路径:");
			String path = scanner.next();
			System.out.println("请输入要抓取的单元格（逗号分隔，如：A1,A2）");
			String cellPositionStr = scanner.next();
			ExcelUtil.grab(path, cellPositionStr);
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
		scanner.next();
		scanner.close();
	}
}
