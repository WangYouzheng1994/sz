package com.sizatn.sz.codegenerate.config;

import java.util.ResourceBundle;

/**
 * 代码生成器 配置
 */
public class CodeConfigProperties {
	private static final ResourceBundle database_bundle = ResourceBundle.getBundle("codegenerate/codegenerate_database");
	private static final ResourceBundle config_bundle = ResourceBundle.getBundle("codegenerate/codegenerate_config");
	/**
	 * 数据库类型
	 */
	public static String databaseType = "mysql";
	/**
	 * 数据库驱动全路径类名
	 */
	public static String diverName = "com.mysql.jdbc.Driver";
	/**
	 * 数据库url
	 */
	public static String url = "jdbc:mysql://localhost:3306/test?useUnicode=true&characterEncoding=UTF-8";
	/**
	 * 数据库账号
	 */
	public static String username = "root";
	/**
	 * 数据库密码
	 */
	public static String password = "mysql";
	/**
	 * 数据库 database
	 */
	public static String databaseName = "test";
	/**
	 * 代码生成的项目地址：根目录
	 */
	public static String projectPath = "c:/workspace/test";
	/**
	 * 基础包名前缀
	 */
	public static String bussiPackage = "com.sizatn";
	/**
	 * 模板地址，当前classpass的相对路径
	 */
	public static String templatePath = "/codegenerate/code-template/";
	/**
	 *
	 */
	public static boolean DB_FILED_CONVERT = true;
	public static String PAGE_FIELD_REQUIRED_NUM = "4";
	public static String PAGE_FIELD_SEARCH_NUM = "3";
	public static String sourceRootPackage;
	public static String q = "1";
	public static String webRootPackage;

	static {
		diverName = getDIVER_NAME();
		url = getURL();
		username = getUSERNAME();
		password = getPASSWORD();
		databaseName = getDATABASE_NAME();

		bussiPackage = getBussiPackage();
		templatePath = getTemplatePath();
		projectPath = getProject_path();

		DB_TABLE_ID = get_generate_table_id();
		DB_FILED_CONVERT = getDB_FILED_CONVERT();

		PAGE_FILTER_FIELDS = get_generate_ui_filter_fields();

		PAGE_FIELD_SEARCH_NUM = get_ui_search_filed_num();
		if (url.indexOf("mysql") >= 0 || url.indexOf("MYSQL") >= 0) {
			databaseType = "mysql";
		} else if (url.indexOf("dm") >= 0 || url.indexOf("DM") >= 0) {
			databaseType = "dm";
		} else if (url.indexOf("oracle") >= 0 || url.indexOf("ORACLE") >= 0) {
			databaseType = "oracle";
		} else if (url.indexOf("postgresql") >= 0 || url.indexOf("POSTGRESQL") >= 0) {
			databaseType = "postgresql";
		} else if (url.indexOf("sqlserver") >= 0 || url.indexOf("sqlserver") >= 0) {
			databaseType = "sqlserver";
		}

		sourceRootPackage = getSourceRootPackage().replace(".", "/");
		webRootPackage = getWebRootPackage().replace(".", "/");
	}

	/**
	 * 表格主键字段名称
	 */
	public static String DB_TABLE_ID;

	/**
	 * 需要过滤的字段名称
	 */
	public static String PAGE_FILTER_FIELDS;

	public static final String getDIVER_NAME() {
		return database_bundle.getString("diver_name");
	}

	public static final String getURL() {
		return database_bundle.getString("url");
	}

	public static final String getUSERNAME() {
		return database_bundle.getString("username");
	}

	public static final String getPASSWORD() {
		return database_bundle.getString("password");
	}

	public static final String getDATABASE_NAME() {
		return database_bundle.getString("database_name");
	}

	public static final boolean getDB_FILED_CONVERT() {
		String s = config_bundle.getString("db_filed_convert");
		if (s.toString().equals("false")) {
			return false;
		}
		return true;
	}

	private static String getBussiPackage() {
		return config_bundle.getString("bussi_package");
	}

	private static String getTemplatePath() {
		return config_bundle.getString("templatepath");
	}

	public static final String getSourceRootPackage() {
		return config_bundle.getString("source_root_package");
	}

	public static final String getWebRootPackage() {
		return config_bundle.getString("webroot_package");
	}

	public static final String get_generate_table_id() {
		return config_bundle.getString("db_table_id");
	}

	public static final String get_generate_ui_filter_fields() {
		return config_bundle.getString("page_filter_fields");
	}

	public static final String get_ui_search_filed_num() {
		return config_bundle.getString("page_search_filed_num");
	}

	public static final String get_ui_field_required_num() {
		return config_bundle.getString("page_field_required_num");
	}

	public static String getProject_path() {
		String projp = config_bundle.getString("project_path");
		if (projp != null && !"".equals(projp)) {
			projectPath = projp;
		}
		return projectPath;
	}

	public static void setProject_path(String paramString) {
		projectPath = paramString;
	}

	public static void setTemplatePath(String paramString) {
		templatePath = paramString;
	}

}