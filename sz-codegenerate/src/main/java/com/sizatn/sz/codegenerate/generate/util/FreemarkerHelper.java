package com.sizatn.sz.codegenerate.generate.util;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.StringReader;
import java.io.StringWriter;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.StringTokenizer;

import freemarker.cache.FileTemplateLoader;
import freemarker.cache.MultiTemplateLoader;
import freemarker.cache.TemplateLoader;
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;

public class FreemarkerHelper {
	public static Configuration newFreeMarkerConfiguration(List<File> paramList, String paramString1,
			String paramString2) throws IOException {
		Configuration localConfiguration = new Configuration();

		FileTemplateLoader[] arrayOfFileTemplateLoader = new FileTemplateLoader[paramList.size()];
		for (int i = 0; i < paramList.size(); i++) {
			arrayOfFileTemplateLoader[i] = new FileTemplateLoader(paramList.get(i));
		}
		MultiTemplateLoader localMultiTemplateLoader = new MultiTemplateLoader(
				(TemplateLoader[]) arrayOfFileTemplateLoader);

		localConfiguration.setTemplateLoader((TemplateLoader) localMultiTemplateLoader);
		localConfiguration.setNumberFormat("###############");
		localConfiguration.setBooleanFormat("true,false");
		localConfiguration.setDefaultEncoding(paramString1);

		return localConfiguration;
	}

	public static List<String> getParentPaths(String paramString1, String paramString2) {
		String[] arrayOfString = tokenizeToStringArray(paramString1, "\\/");
		ArrayList<String> localArrayList = new ArrayList<String>();
		localArrayList.add(paramString2);
		localArrayList.add(File.separator + paramString2);
		String str = "";
		for (int i = 0; i < arrayOfString.length; i++) {

			str = str + File.separator + arrayOfString[i];
			localArrayList.add(str + File.separator + paramString2);
		}
		return localArrayList;
	}

	public static String[] tokenizeToStringArray(String paramString1, String paramString2) {
		if (paramString1 == null) {
			return new String[0];
		}
		StringTokenizer localStringTokenizer = new StringTokenizer(paramString1, paramString2);
		ArrayList<String> localArrayList = new ArrayList<String>();
		while (localStringTokenizer.hasMoreElements()) {

			Object localObject = localStringTokenizer.nextElement();
			localArrayList.add(localObject.toString());
		}
		return localArrayList.<String>toArray(new String[localArrayList.size()]);
	}

	/**
	 *
	 * @param paramString
	 * @param paramMap
	 * @param paramConfiguration
	 * @return
	 */
	public static String processTemplateString(String paramString, Map<String, Object> paramMap,
			Configuration paramConfiguration) {
		StringWriter localStringWriter = new StringWriter();

		try {
			Template localTemplate = new Template("templateString...", new StringReader(paramString),
					paramConfiguration);
			localTemplate.process(paramMap, localStringWriter);
			return localStringWriter.toString();
		} catch (Exception localException) {

			throw new IllegalStateException("cannot process templateString:" + paramString + " cause:" + localException,
					localException);
		}
	}

	public static void processTemplate(Template paramTemplate, Map<String, Object> paramMap, File paramFile,
			String paramString) throws IOException, TemplateException {
		BufferedWriter localBufferedWriter = new BufferedWriter(
				new OutputStreamWriter(new FileOutputStream(paramFile), paramString));

		paramMap.put("Format", new SimpleFormat());
		paramTemplate.process(paramMap, localBufferedWriter);
		localBufferedWriter.close();
	}
}
