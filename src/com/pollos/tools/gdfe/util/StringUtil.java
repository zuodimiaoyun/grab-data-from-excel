package com.pollos.tools.gdfe.util;

public class StringUtil {
	public static boolean isNullOrEmpty(String str){
		if(str == null || str.trim().isEmpty()){
			return true;
		}
		return false;
	}
}
