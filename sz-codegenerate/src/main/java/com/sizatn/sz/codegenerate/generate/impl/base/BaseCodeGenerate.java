package com.sizatn.sz.codegenerate.generate.impl.base;

import com.sizatn.sz.codegenerate.config.CodeConfigProperties;
import com.sizatn.sz.codegenerate.generate.base.CreateFileConfig;
import com.sizatn.sz.codegenerate.generate.util.FileHelper;
import com.sizatn.sz.codegenerate.generate.util.FreemarkerHelper;
import freemarker.template.Configuration;
import freemarker.template.Template;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * 代码生成器基类
 * 核心方法为 #generateFileCommon()
 */
public class BaseCodeGenerate {
	private static final Logger log = LoggerFactory.getLogger(BaseCodeGenerate.class);
	protected static String sourceEncoding = "UTF-8";

	protected void generateFileCommon(CreateFileConfig createFileConfig, String projectPath,
			Map<String, Object> templateData) throws Exception {
		log.debug("--------projectPath--------" + projectPath);
		for (int i = 0; i < createFileConfig.getTemplateRootDirs().size(); i++) {
			File templateRootDir = createFileConfig.getTemplateRootDirs().get(i);
			scanTemplatesAndProcess(projectPath, templateRootDir, templateData, createFileConfig);
		}
	}

	/**
	 * 扫描指定路径下的模板文件，并且将传递的模板参数进行替换
	 *
	 * @param projectPath
	 * @param templateRootDir
	 * @param templateData
	 * @param createFileConfig
	 * @throws Exception
	 */
	protected void scanTemplatesAndProcess(String projectPath, File templateRootDir, Map<String, Object> templateData,
			CreateFileConfig createFileConfig) throws Exception {
		if (templateRootDir == null) {
			throw new IllegalStateException("'templateRootDir' must be not null");
		}
		log.debug("-------------------load template from templateRootDir = '" + templateRootDir.getAbsolutePath()
				+ "' outJavaRootDir:"
				+ (new File(CodeConfigProperties.sourceRootPackage.replace(".", File.separator))).getAbsolutePath()
				+ "' outWebappRootDir:"
				+ (new File(CodeConfigProperties.webRootPackage.replace(".", File.separator))).getAbsolutePath());

		List<File> localList = FileHelper.searchAllNotIgnoreFile(templateRootDir);
		log.debug("----srcFiles----size-----------" + localList.size());
		log.debug("----srcFiles----list------------" + localList.toString());
		for (int i = 0; i < localList.size(); i++) {
			File srcFile = localList.get(i);
			executeGenerate(projectPath, templateRootDir, templateData, srcFile, createFileConfig);
		}
	}

	protected void executeGenerate(String projectPath, File templateRootDir, Map<String, Object> templateData,
			File srcFile, CreateFileConfig createFileConfig) throws Exception {
		log.debug("-------templateRootDir--" + templateRootDir.getPath());
		log.debug("-------srcFile--" + srcFile.getPath());

		// 获取模板文件路径
		String templateFile = FileHelper.getRelativePath(templateRootDir, srcFile);
		try {
			log.debug("-------templateFile--" + templateFile);
			//
			String outputFilepath = proceeForOutputFilepath(templateData, templateFile, createFileConfig);
			log.debug("-------outputFilepath--" + outputFilepath);
			if (outputFilepath.startsWith("java")) {
				String path = projectPath + File.separator
						+ CodeConfigProperties.sourceRootPackage.replace(".", File.separator);

				String soure = path;
				outputFilepath = outputFilepath.substring("java".length());
				outputFilepath = soure + outputFilepath;
				log.debug("-------java----outputFilepath--" + outputFilepath);
				generateNewFileOrInsertIntoFile(templateFile, outputFilepath, templateData, createFileConfig);
			} else if (outputFilepath.startsWith("webapp")) {
				String path = projectPath + File.separator
						+ CodeConfigProperties.webRootPackage.replace(".", File.separator);

				String soure = path;
				outputFilepath = outputFilepath.substring("webapp".length());
				outputFilepath = soure + outputFilepath;
				log.debug("-------webapp---outputFilepath---" + outputFilepath);
				generateNewFileOrInsertIntoFile(templateFile, outputFilepath, templateData, createFileConfig);
			}
		} catch (Exception e) {
			e.printStackTrace();
			log.error(e.toString());
		}
	}

	protected void generateNewFileOrInsertIntoFile(String templateFile, String outputFilepath,
			Map<String, Object> templateModel, CreateFileConfig createFileConfig) throws Exception {
		if (outputFilepath.endsWith("i")) {
			outputFilepath = outputFilepath.substring(0, outputFilepath.length() - 1);
		}
		Template template = getFreeMarkerTemplate(templateFile, createFileConfig);
		template.setOutputEncoding(sourceEncoding);
		File absoluteOutputFilePath = FileHelper.parentMkdir(outputFilepath);
		log.info("[generate]\t template:" + templateFile + " ==> " + outputFilepath);
		FreemarkerHelper.processTemplate(template, templateModel, absoluteOutputFilePath, sourceEncoding);
		if (isCutFile(absoluteOutputFilePath)) {
			splitFile(absoluteOutputFilePath, "#segment#");
		}
	}

	protected Template getFreeMarkerTemplate(String templateName, CreateFileConfig createFileConfig)
			throws IOException {
		return FreemarkerHelper
				.newFreeMarkerConfiguration(createFileConfig.getTemplateRootDirs(), sourceEncoding, templateName)
				.getTemplate(templateName);
	}

	protected boolean isCutFile(File file) {
		if (file.getName().startsWith("[1-n]")) {
			return true;
		}
		return false;
	}

	protected static void splitFile(File paramFile, String paramString) {
		InputStreamReader localInputStreamReader = null;
		BufferedReader localBufferedReader = null;
		ArrayList<OutputStreamWriter> localArrayList = new ArrayList<OutputStreamWriter>();
		try {
			localInputStreamReader = new InputStreamReader(new FileInputStream(paramFile), "UTF-8");
			localBufferedReader = new BufferedReader(localInputStreamReader);

			int m = 0;
			OutputStreamWriter localOutputStreamWriter = null;
			String str1;
			while ((str1 = localBufferedReader.readLine()) != null) {
				if (str1.trim().length() > 0 && str1.startsWith(paramString)) {
					String str2 = str1.substring(paramString.length());
					String str3 = paramFile.getParentFile().getAbsolutePath();
					str2 = str3 + File.separator + str2;
					log.debug("[generate]\t split file:" + paramFile.getAbsolutePath() + " ==> " + str2);

					localOutputStreamWriter = new OutputStreamWriter(new FileOutputStream(str2), "UTF-8");
					localArrayList.add(localOutputStreamWriter);
					m = 1;
					continue;
				}
				if (m != 0) {
					log.debug("row : " + str1);
					localOutputStreamWriter.append(str1 + "\r\n");
				}
			}
			for (int n = 0; n < localArrayList.size(); n++) {
				((Writer) localArrayList.get(n)).close();
			}
			localBufferedReader.close();
			localInputStreamReader.close();

			log.debug("[generate]\t delete file:" + paramFile.getAbsolutePath());
			forceDelete(paramFile);

			return;
		} catch (FileNotFoundException localFileNotFoundException) {
			localFileNotFoundException.printStackTrace();
		} catch (IOException localIOException3) {
			localIOException3.printStackTrace();
		} finally {
			try {
				if (localBufferedReader != null) {
					localBufferedReader.close();
				}
				if (localInputStreamReader != null) {
					localInputStreamReader.close();
				}
				if (localArrayList.size() > 0) {
					for (int i1 = 0; i1 < localArrayList.size(); i1++) {
						if (localArrayList.get(i1) != null) {
							((Writer) localArrayList.get(i1)).close();
						}
					}
				}
			} catch (IOException localIOException5) {
				localIOException5.printStackTrace();
			}
		}
	}

	/**
	 * 输出代码文件=
	 *
	 * @param filePathModel 模板替换参数
	 * @param templateFile 模板路径文件
	 * @param createFileConfig 模板配置对象  用以freemaker
	 * @return
	 * @throws Exception
	 */
	protected static String proceeForOutputFilepath(Map<String, Object> filePathModel, String templateFile,
			CreateFileConfig createFileConfig) throws Exception {
		String outputFilePath = templateFile;

		int testExpressionIndex = -1;
		if ((testExpressionIndex = templateFile.indexOf('@')) != -1) {
			outputFilePath = templateFile.substring(0, testExpressionIndex);
			String testExpressionKey = templateFile.substring(testExpressionIndex + 1);
			Object expressionValue = filePathModel.get(testExpressionKey);
			if (expressionValue == null) {
				System.err.println("[not-generate] WARN: test expression is null by key:[" + testExpressionKey
						+ "] on template:[" + templateFile + "]");
				return null;
			}
			if (!"true".equals(String.valueOf(expressionValue))) {
				log.error("[not-generate]\t test expression '@" + testExpressionKey + "' is false,template:"
						+ templateFile);
				return null;
			}
		}
		Configuration conf = FreemarkerHelper.newFreeMarkerConfiguration(createFileConfig.getTemplateRootDirs(),
				sourceEncoding, "/");
		outputFilePath = FreemarkerHelper.processTemplateString(outputFilePath, filePathModel, conf);
		// 扩展名
		String extName = outputFilePath.substring(outputFilePath.lastIndexOf("."));
		// 文件名
		String fileName = outputFilePath.substring(0, outputFilePath.lastIndexOf(".")).replace(".", File.separator);
		//
		outputFilePath = fileName + extName;
		return outputFilePath;
	}

	protected static boolean forceDelete(File f) {
		boolean result = false;
		int tryCount = 0;
		while (!result && tryCount++ < 10) {
			System.gc();
			result = f.delete();
		}
		return result;
	}

	protected static String a(String paramString1, String paramString2) {
		int i = 1;
		int j = 1;

		while (true) {
			int k = (paramString1.indexOf(paramString2) == 0) ? 1 : 0;
			int m = (paramString1.lastIndexOf(paramString2) + 1 == paramString1.length())
					? paramString1.lastIndexOf(paramString2)
					: paramString1.length();
			paramString1 = paramString1.substring(k, m);
			i = (paramString1.indexOf(paramString2) == 0) ? 1 : 0;
			j = (paramString1.lastIndexOf(paramString2) + 1 == paramString1.length()) ? 1 : 0;
			if (i == 0 && j == 0)
				return paramString1;
		}
	}
}
