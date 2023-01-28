package com.sizatn.sz.codegenerate.generate.base;

import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 模板配置类
 */
public class CreateFileConfig {
	private static final Logger log = LoggerFactory.getLogger(CreateFileConfig.class);
	/**
	 * 模板所在路径
	 */
	private String templatePath;
	private List<File> templateRootDirs = new ArrayList<>();

	private String stylePath;

	public CreateFileConfig(String templatePath) {
		log.info("----templatePath-----------------" + templatePath);
		this.templatePath = templatePath;
	}

	private void setTemplateRootDir(File templateRootDir) {
		setTemplateRootDirs(new File[] { templateRootDir });
	}

	private void setTemplateRootDirs(File... templateRootDirs) {
		this.templateRootDirs = Arrays.asList(templateRootDirs);
	}

	public String getStylePath() {
		return this.stylePath;
	}

	public void setStylePath(String stylePath) {
		this.stylePath = stylePath;
	}

	public List<File> getTemplateRootDirs() {
		String str = getClass().getResource(this.templatePath).getFile();
		str = str.replaceAll("%20", " ");
		log.debug("-------classpath-------" + str);
		setTemplateRootDir(new File(str));
		return this.templateRootDirs;
	}

	public void setTemplateRootDirs(List<File> paramList) {
		this.templateRootDirs = paramList;
	}

	public String toString() {
		StringBuilder localStringBuilder = new StringBuilder();
		localStringBuilder.append("{\"templateRootDirs\":\"");
		localStringBuilder.append(this.templateRootDirs);
		localStringBuilder.append("\",\"stylePath\":\"");
		localStringBuilder.append(this.stylePath);
		localStringBuilder.append("\"} ");
		return localStringBuilder.toString();
	}
}