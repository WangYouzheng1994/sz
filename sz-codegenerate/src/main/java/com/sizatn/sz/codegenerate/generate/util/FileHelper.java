package com.sizatn.sz.codegenerate.generate.util;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 文件处理器
 */
public class FileHelper {
	private static final Logger log = LoggerFactory.getLogger(FileHelper.class);

	public static List<File> searchAllNotIgnoreFile(File paramFile) throws IOException {
		ArrayList<File> localArrayList = new ArrayList<>();
		searchAllNotIgnoreFile(paramFile, localArrayList);
		Collections.sort(localArrayList, (paramAnonymousFile1, paramAnonymousFile2) -> paramAnonymousFile1
				.getAbsolutePath().compareTo(paramAnonymousFile2.getAbsolutePath()));
		return localArrayList;
	}

	public static void searchAllNotIgnoreFile(File paramFile, List<File> paramList) throws IOException {
		log.debug("---------dir------------path: " + paramFile.getPath() + " -- isHidden --: " + paramFile.isHidden()
				+ " -- isDirectory --: " + paramFile.isDirectory());
		if (!paramFile.isHidden() && paramFile.isDirectory() && !isIgnoreFile(paramFile)) {

			File[] arrayOfFile = paramFile.listFiles();
			for (int i = 0; i < arrayOfFile.length; i++) {
				searchAllNotIgnoreFile(arrayOfFile[i], paramList);
			}
		} else if (!isIgnoreEndWithFile(paramFile) && !isIgnoreFile(paramFile)) {

			paramList.add(paramFile);
		}
	}

	public static String getRelativePath(File paramFile1, File paramFile2) {
		if (paramFile1.equals(paramFile2)) {
			return "";
		}
		if (paramFile1.getParentFile() == null) {
			return paramFile2.getAbsolutePath().substring(paramFile1.getAbsolutePath().length());
		}
		return paramFile2.getAbsolutePath().substring(paramFile1.getAbsolutePath().length() + 1);
	}

	public static boolean isBinaryFile(File paramFile) {
		if (paramFile.isDirectory()) {
			return false;
		}
		return isBinaryFile(paramFile.getName());
	}

	public static boolean isBinaryFile(String paramString) {
		if (StringUtils.isBlank(getExtension(paramString))) {
			return false;
		}
		return true;
	}

	public static String getExtension(String paramString) {
		if (paramString == null) {
			return null;
		}
		int i = paramString.indexOf(".");
		if (i == -1) {
			return "";
		}
		return paramString.substring(i + 1);
	}

	public static File parentMkdir(String paramString) {
		if (paramString == null) {
			throw new IllegalArgumentException("file must be not null");
		}
		File localFile = new File(paramString);
		parentMkdir(localFile);
		return localFile;
	}

	public static void parentMkdir(File paramFile) {
		if (paramFile.getParentFile() != null) {
			paramFile.getParentFile().mkdirs();
		}
	}

	public static List<String> ignoreList = new ArrayList<>();
	public static List<String> ignoreEndWithList = new ArrayList<>();

	static {
		ignoreList.add(".svn");
		ignoreList.add("CVS");
		ignoreList.add(".cvsignore");
		ignoreList.add(".copyarea.db");
		ignoreList.add("SCCS");
		ignoreList.add("vssver.scc");
		ignoreList.add(".DS_Store");
		ignoreList.add(".git");
		ignoreList.add(".gitignore");
		ignoreEndWithList.add(".ftl");
	}

	private static boolean isIgnoreFile(File paramFile) {
		for (int i = 0; i < ignoreList.size(); i++) {
			if (paramFile.getName().equals(ignoreList.get(i))) {
				return true;
			}
		}
		return false;
	}

	private static boolean isIgnoreEndWithFile(File paramFile) {
		for (int i = 0; i < ignoreEndWithList.size(); i++) {
			if (paramFile.getName().endsWith(ignoreEndWithList.get(i))) {
				return true;
			}
		}
		return false;
	}


	public static void main(String[] args) {
		File file = new File("C://1.txt");

		System.out.println(file.exists());
		System.out.println(file.getAbsolutePath());

	}
}
