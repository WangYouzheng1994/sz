package com.sizatn.sz.codegenerate.generate.impl;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.sizatn.sz.codegenerate.config.CodeConfigProperties;
import com.sizatn.sz.codegenerate.database.DbReadTableUtil;
import com.sizatn.sz.codegenerate.generate.IGenerate;
import com.sizatn.sz.codegenerate.generate.base.CreateFileConfig;
import com.sizatn.sz.codegenerate.generate.impl.base.BaseCodeGenerate;
import com.sizatn.sz.codegenerate.generate.pojo.ColumnVo;
import com.sizatn.sz.codegenerate.generate.pojo.TableVo;
import com.sizatn.sz.codegenerate.generate.util.NonceUtils;

public class CodeGenerateOne extends BaseCodeGenerate implements IGenerate {
	private static final Logger log = LoggerFactory.getLogger(CodeGenerateOne.class);

	private TableVo tableVo;
	private List<ColumnVo> columns;
	private List<ColumnVo> originalColumns;

	public CodeGenerateOne(TableVo tableVo) {
		this.tableVo = tableVo;
	}

	public CodeGenerateOne(TableVo tableVo, List<ColumnVo> columns, List<ColumnVo> originalColumns) {
		this.tableVo = tableVo;
		this.columns = columns;
		this.originalColumns = originalColumns;
	}

	public Map<String, Object> initData() throws Exception {
		HashMap<Object, Object> localHashMap = new HashMap<>();

		// service层 包
		localHashMap.put("bussiPackage", CodeConfigProperties.bussiPackage);

		// 实体类层 包
		localHashMap.put("entityPackage", this.tableVo.getEntityPackage());

		// 实体类名称
		localHashMap.put("entityName", this.tableVo.getEntityName());

		// 表名
		localHashMap.put("tableName", this.tableVo.getTableName());

		// 主键字段名称
		localHashMap.put("primaryKeyField", CodeConfigProperties.DB_TABLE_ID);

		// 主键Id策略
		localHashMap.put("tableId", this.tableVo.getPrimaryKeyPolicy());
		if (this.tableVo.getFieldRequiredNum() == null) {
			this.tableVo.setFieldRequiredNum(
					Integer.valueOf(StringUtils.isNotEmpty(CodeConfigProperties.PAGE_FIELD_REQUIRED_NUM)
							? Integer.parseInt(CodeConfigProperties.PAGE_FIELD_REQUIRED_NUM)
							: -1));
		}
		if (this.tableVo.getSearchFieldNum() == null) {
			this.tableVo.setSearchFieldNum(
					Integer.valueOf(StringUtils.isNotEmpty(CodeConfigProperties.PAGE_FIELD_SEARCH_NUM)
							? Integer.parseInt(CodeConfigProperties.PAGE_FIELD_SEARCH_NUM)
							: -1));
		}
		if (this.tableVo.getFieldRowNum() == null) {
			this.tableVo.setFieldRowNum(Integer.valueOf(Integer.parseInt(CodeConfigProperties.q)));
		}
		localHashMap.put("tableVo", this.tableVo);

		try {
			if (this.columns == null || this.columns.size() == 0) {
				this.columns = DbReadTableUtil.readTableColumn(this.tableVo.getTableName());
			}
			localHashMap.put("columns", this.columns);
			if (this.originalColumns == null || this.originalColumns.size() == 0) {
				this.originalColumns = DbReadTableUtil.readOriginalTableColumn(this.tableVo.getTableName());
			}
			localHashMap.put("originalColumns", this.originalColumns);
			for (ColumnVo localColumnVo : this.originalColumns) {
				if (localColumnVo.getFieldName().toLowerCase().equals(CodeConfigProperties.DB_TABLE_ID.toLowerCase())) {
					localHashMap.put("primaryKeyPolicy", localColumnVo.getFieldType());
				}
			}

		} catch (Exception localException) {

			throw localException;
		}
		long serialVersionUID = NonceUtils.randomLong() + NonceUtils.currentMills();
		localHashMap.put("serialVersionUID", Long.valueOf(serialVersionUID));
		log.info("code template data: " + localHashMap.toString());
		return (Map) localHashMap;
	}

	/**
	 * 生成代码文件
	 *
	 * @throws Exception
	 */
	public void generateCodeFile() throws Exception {
		log.info("---Code----Generation----[单表模型:" + this.tableVo.getTableName() + "]------- 生成中。。。");

		String projectPath = CodeConfigProperties.projectPath;
		Map<String, Object> templateData = initData();

		// 模板配置类
		CreateFileConfig createFileConfig = new CreateFileConfig(CodeConfigProperties.templatePath);

		try {
			// 生成文件核心代码
			generateFileCommon(createFileConfig, projectPath, templateData);
		} catch (Exception e) {

			e.printStackTrace();
		}
		log.info("----Code----Generation-----[单表模型：" + this.tableVo.getTableName() + "]------ 生成完成。。。");
	}

	public void generateCodeFile(String projectPath, String templatePath) throws Exception {
		if (projectPath != null && !"".equals(projectPath)) {
			CodeConfigProperties.setProject_path(projectPath);
		}
		if (templatePath != null && !"".equals(templatePath)) {
			CodeConfigProperties.setTemplatePath(templatePath);
		}
		generateCodeFile();
	}

	public static void main(String[] args) {
		System.out.println("------ Code------------- Generation -----[单表模型]------- 生成中。。。");
		TableVo localTableVo = new TableVo();
		localTableVo.setTableName("demo");
		localTableVo.setPrimaryKeyPolicy("uuid");
		localTableVo.setEntityPackage("test");
		localTableVo.setEntityName("Demo");
		localTableVo.setFtlDescription("测试demo");

		try {
			(new CodeGenerateOne(localTableVo)).generateCodeFile();
		} catch (Exception localException) {

			localException.printStackTrace();
		}
		System.out.println("------ Code------------- Generation -----[单表模型]------- 生成完成。。。");
	}
}
