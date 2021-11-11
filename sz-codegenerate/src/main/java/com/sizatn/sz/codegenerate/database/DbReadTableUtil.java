package com.sizatn.sz.codegenerate.database;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.sizatn.sz.codegenerate.config.CodeConfigProperties;
import com.sizatn.sz.codegenerate.database.util.CodeStringUtils;
import com.sizatn.sz.codegenerate.generate.pojo.ColumnVo;
import com.sizatn.sz.codegenerate.generate.util.TableConvert;

public class DbReadTableUtil {
	private static final Logger log = LoggerFactory.getLogger(DbReadTableUtil.class);
	private static Connection conn;
	private static Statement stmt;

	public static List<String> readAllTableNames() throws SQLException {
		String sql = null;
		List<String> tableNames = new ArrayList<String>(0);
		try {
			Class.forName(CodeConfigProperties.diverName);
			conn = DriverManager.getConnection(CodeConfigProperties.url, CodeConfigProperties.username,
					CodeConfigProperties.password);
			stmt = conn.createStatement(1005, 1007);
			if (CodeConfigProperties.databaseType.equals("mysql")) {
				sql = MessageFormat.format(
						"select distinct table_name from information_schema.columns where table_schema = {0}",
						new Object[] { TableConvert.getV(CodeConfigProperties.databaseName) });
			}
			if (CodeConfigProperties.databaseType.equals("dm")) {
				sql = "select distinct table_name from user_tables where tablespace_name = 'MAIN'";
			}
			if (CodeConfigProperties.databaseType.equals("oracle")) {
				sql = " select distinct colstable.table_name as  table_name from user_tab_cols colstable order by colstable.table_name";
			}
			if (CodeConfigProperties.databaseType.equals("postgresql")) {
				sql = "SELECT distinct c.relname AS  table_name FROM pg_class c";
			}
			if (CodeConfigProperties.databaseType.equals("sqlserver")) {
				sql = "select distinct c.name as  table_name from sys.objects c where c.type = 'U' ";
			}
			ResultSet rs = stmt.executeQuery(sql);
			while (rs.next()) {

				String tableName = rs.getString(1);
				tableNames.add(tableName);
			}
			return tableNames;
		} catch (Exception e) {

			e.printStackTrace();
		} finally {
			try {
				if (stmt != null) {
					stmt.close();
					stmt = null;
					System.gc();
				}
				if (conn != null) {
					conn.close();
					conn = null;
					System.gc();
				}
			} catch (SQLException e) {
				throw e;
			}
		}
		return tableNames;
	}

	public static List<ColumnVo> readTableColumn(String tableName) throws Exception {
		ResultSet localResultSet = null;
		String sql = null;
		List<ColumnVo> columntList = new ArrayList<>();

		try {
			Class.forName(CodeConfigProperties.diverName);
			conn = DriverManager.getConnection(CodeConfigProperties.url, CodeConfigProperties.username,
					CodeConfigProperties.password);
			stmt = conn.createStatement(1005, 1007);
			if (CodeConfigProperties.databaseType.equals("mysql")) {
				sql = MessageFormat.format(
						"select column_name,data_type,column_comment,numeric_precision,numeric_scale,character_maximum_length,is_nullable nullable from information_schema.columns where table_name = {0} and table_schema = {1}",
						new Object[] { TableConvert.getV(tableName.toUpperCase()),
								TableConvert.getV(CodeConfigProperties.databaseName) });
			}
			if (CodeConfigProperties.databaseType.equals("dm")) {
				sql = MessageFormat.format(
						"select atc.column_name, atc.data_type, acc.comments, atc.data_precision, atc.data_scale, atc.char_length, case atc.nullable when '''N''' then '''NO''' when '''Y''' then '''YES''' end nullable from all_tab_columns atc left join all_col_comments acc on acc.column_name = atc.column_name where atc.table_name = {0} and acc.table_name = {1} and atc.owner = {2} and acc.owner = {3}",
						new Object[] { TableConvert.getV(tableName.toUpperCase()), TableConvert.getV(tableName.toUpperCase()),
								TableConvert.getV(CodeConfigProperties.databaseName), TableConvert.getV(CodeConfigProperties.databaseName) });
			}
			if (CodeConfigProperties.databaseType.equals("oracle")) {
				sql = MessageFormat.format(
						" select colstable.column_name column_name, colstable.data_type data_type, commentstable.comments column_comment, colstable.Data_Precision column_precision, colstable.Data_Scale column_scale,colstable.Char_Length,colstable.nullable from user_tab_cols colstable  inner join user_col_comments commentstable  on colstable.column_name = commentstable.column_name  where colstable.table_name = commentstable.table_name  and colstable.table_name = {0}",
						new Object[] { TableConvert.getV(tableName.toUpperCase()) });
			}
			if (CodeConfigProperties.databaseType.equals("postgresql")) {
				sql = MessageFormat.format(
						"SELECT CodeStringUtils.attname AS  field,t.typname AS type,col_description(CodeStringUtils.attrelid,CodeStringUtils.attnum) as comment,null as column_precision,null as column_scale,null as Char_Length,CodeStringUtils.attnotnull  FROM pg_class c,pg_attribute  CodeStringUtils,pg_type t  WHERE c.relname = {0} and CodeStringUtils.attnum > 0  and CodeStringUtils.attrelid = c.oid and CodeStringUtils.atttypid = t.oid  ORDER BY CodeStringUtils.attnum ",
						new Object[] { TableConvert.getV(tableName.toLowerCase()) });
			}
			if (CodeConfigProperties.databaseType.equals("sqlserver")) {
				sql = MessageFormat.format(
						"select distinct cast(CodeStringUtils.name as varchar(50)) column_name,  cast(b.name as varchar(50)) data_type,  cast(e.value as varchar(200)) comment,  cast(ColumnProperty(CodeStringUtils.object_id,CodeStringUtils.Name,'''Precision''') as int) num_precision,  cast(ColumnProperty(CodeStringUtils.object_id,CodeStringUtils.Name,'''Scale''') as int) num_scale,  CodeStringUtils.max_length,  (case when CodeStringUtils.is_nullable=1 then '''y''' else '''n''' end) nullable,column_id   from sys.columns CodeStringUtils left join sys.types b on CodeStringUtils.user_type_id=b.user_type_id left join (select top 1 * from sys.objects where type = '''U''' and name ={0}  order by name) c on CodeStringUtils.object_id=c.object_id left join sys.extended_properties e on e.major_id=c.object_id and e.minor_id=CodeStringUtils.column_id and e.class=1 where c.name={0} order by CodeStringUtils.column_id",
						new Object[] { TableConvert.getV(tableName.toLowerCase()) });
			}
			localResultSet = stmt.executeQuery(sql);
			localResultSet.last();
			int fieldNum = localResultSet.getRow();
			int n = fieldNum;
			int i = n;
			if (i > 0) {

				ColumnVo columnVo = new ColumnVo();
				if (CodeConfigProperties.DB_FILED_CONVERT) {
					columnVo.setFieldName(e(localResultSet.getString(1).toLowerCase()));
				} else {
					columnVo.setFieldName(localResultSet.getString(1).toLowerCase());
				}
				columnVo.setFieldDbName(localResultSet.getString(1).toUpperCase());
				columnVo.setFieldType(e(localResultSet.getString(2).toLowerCase()));
				columnVo.setFieldDbType(e(localResultSet.getString(2).toLowerCase()));
				columnVo.setPrecision(localResultSet.getString(4));
				columnVo.setScale(localResultSet.getString(5));
				columnVo.setCharmaxLength(localResultSet.getString(6));
				columnVo.setNullable(TableConvert.getNullAble(localResultSet.getString(7)));
				a(columnVo);
				columnVo.setFiledComment(StringUtils.isBlank(localResultSet.getString(3)) ? columnVo.getFieldName()
						: localResultSet.getString(3));

				log.debug("columnt.getFieldName() -------------" + columnVo.getFieldName());

				String[] arrayOfString = new String[0];
				if (CodeConfigProperties.PAGE_FILTER_FIELDS != null) {
					arrayOfString = CodeConfigProperties.PAGE_FILTER_FIELDS.toLowerCase().split(",");
				}
				if (!CodeConfigProperties.DB_TABLE_ID.equals(columnVo.getFieldName())
						&& !CodeStringUtils.isIn(columnVo.getFieldDbName().toLowerCase(), arrayOfString)) {
					columntList.add(columnVo);
				}
				while (localResultSet.previous()) {

					ColumnVo localColumnVo2 = new ColumnVo();
					if (CodeConfigProperties.DB_FILED_CONVERT) {
						localColumnVo2.setFieldName(e(localResultSet.getString(1).toLowerCase()));
					} else {
						localColumnVo2.setFieldName(localResultSet.getString(1).toLowerCase());
					}
					localColumnVo2.setFieldDbName(localResultSet.getString(1).toUpperCase());
					log.debug("columnt.getFieldName() -------------" + localColumnVo2.getFieldName());
					if (!CodeConfigProperties.DB_TABLE_ID.equals(localColumnVo2.getFieldName())
							&& !CodeStringUtils.isIn(localColumnVo2.getFieldDbName().toLowerCase(), arrayOfString)) {
						localColumnVo2.setFieldType(e(localResultSet.getString(2).toLowerCase()));

						localColumnVo2.setFieldDbType(e(localResultSet.getString(2).toLowerCase()));
						log.debug("-----po.setFieldType------------" + localColumnVo2.getFieldType());
						localColumnVo2.setPrecision(localResultSet.getString(4));
						localColumnVo2.setScale(localResultSet.getString(5));
						localColumnVo2.setCharmaxLength(localResultSet.getString(6));
						localColumnVo2.setNullable(TableConvert.getNullAble(localResultSet.getString(7)));

						a(localColumnVo2);
						localColumnVo2.setFiledComment(
								StringUtils.isBlank(localResultSet.getString(3)) ? localColumnVo2.getFieldName()
										: localResultSet.getString(3));
						columntList.add(localColumnVo2);
					}

				}
			} else {

				throw new Exception("该表不存在或者表中没有字段");
			}
			log.debug("读取表成功");

			try {
				if (stmt != null) {

					stmt.close();
					stmt = null;
					System.gc();
				}
				if (conn != null) {
					conn.close();
					conn = null;
					System.gc();
				}

			} catch (SQLException localSQLException1) {

				throw localSQLException1;
			}
		} catch (ClassNotFoundException localClassNotFoundException) {

			throw localClassNotFoundException;
		} catch (SQLException localSQLException2) {

			throw localSQLException2;
		} finally {

			try {

				if (stmt != null) {

					stmt.close();
					stmt = null;
					System.gc();
				}
				if (conn != null) {
					conn.close();
					conn = null;
					System.gc();
				}

			} catch (SQLException e) {

				throw e;
			}
		}
		ArrayList<ColumnVo> localArrayList2 = new ArrayList<ColumnVo>();
		for (int j = columntList.size() - 1; j >= 0; j--) {

			ColumnVo localColumnVo1 = columntList.get(j);
			localArrayList2.add(localColumnVo1);
		}
		return localArrayList2;
	}

	public static List<ColumnVo> readOriginalTableColumn(String paramString) throws Exception {
		ResultSet localResultSet = null;
		String str = null;
		ArrayList<ColumnVo> localArrayList1 = new ArrayList<ColumnVo>();

		try {
			Class.forName(CodeConfigProperties.diverName);
			conn = DriverManager.getConnection(CodeConfigProperties.url, CodeConfigProperties.username,
					CodeConfigProperties.password);
			stmt = conn.createStatement(1005, 1007);
			if (CodeConfigProperties.databaseType.equals("mysql")) {
				str = MessageFormat.format(
						"select column_name,data_type,column_comment,numeric_precision,numeric_scale,character_maximum_length,is_nullable nullable from information_schema.columns where table_name = {0} and table_schema = {1}",
						new Object[] { TableConvert.getV(paramString.toUpperCase()),
								TableConvert.getV(CodeConfigProperties.databaseName) });
			}
			if (CodeConfigProperties.databaseType.equals("dm")) {
				str = MessageFormat.format(
						"select atc.column_name, atc.data_type, acc.comments, atc.data_precision, atc.data_scale, atc.char_length, case atc.nullable when '''N''' then '''NO''' when '''Y''' then '''YES''' end nullable from all_tab_columns atc left join all_col_comments acc on acc.column_name = atc.column_name where atc.table_name = {0} and acc.table_name = {1} and atc.owner = {2} and acc.owner = {3}",
						new Object[] { TableConvert.getV(paramString.toUpperCase()), TableConvert.getV(paramString.toUpperCase()),
								TableConvert.getV(CodeConfigProperties.databaseName), TableConvert.getV(CodeConfigProperties.databaseName) });
			}
			if (CodeConfigProperties.databaseType.equals("oracle")) {
				str = MessageFormat.format(
						" select colstable.column_name column_name, colstable.data_type data_type, commentstable.comments column_comment, colstable.Data_Precision column_precision, colstable.Data_Scale column_scale,colstable.Char_Length,colstable.nullable from user_tab_cols colstable  inner join user_col_comments commentstable  on colstable.column_name = commentstable.column_name  where colstable.table_name = commentstable.table_name  and colstable.table_name = {0}",
						new Object[] { TableConvert.getV(paramString.toUpperCase()) });
			}
			if (CodeConfigProperties.databaseType.equals("postgresql")) {
				str = MessageFormat.format(
						"SELECT base.attname AS  field,t.typname AS type,col_description(base.attrelid,base.attnum) as comment,null as column_precision,null as column_scale,null as Char_Length,base.attnotnull  FROM pg_class c,pg_attribute  base,pg_type t  WHERE c.relname = {0} and base.attnum > 0  and base.attrelid = c.oid and base.atttypid = t.oid  ORDER BY base.attnum ",
						new Object[] { TableConvert.getV(paramString.toLowerCase()) });
			}
			if (CodeConfigProperties.databaseType.equals("sqlserver")) {
				str = MessageFormat.format(
						"select distinct cast(base.name as varchar(50)) column_name,  cast(b.name as varchar(50)) data_type,  cast(e.value as varchar(200)) comment,  cast(ColumnProperty(base.object_id,base.Name,'''Precision''') as int) num_precision,  cast(ColumnProperty(base.object_id,base.Name,'''Scale''') as int) num_scale,  base.max_length,  (case when base.is_nullable=1 then '''y''' else '''n''' end) nullable,column_id   from sys.columns base left join sys.types b on base.user_type_id=b.user_type_id left join (select top 1 * from sys.objects where type = '''U''' and name ={0}  order by name) c on base.object_id=c.object_id left join sys.extended_properties e on e.major_id=c.object_id and e.minor_id=base.column_id and e.class=1 where c.name={0} order by base.column_id",
						new Object[] { TableConvert.getV(paramString.toLowerCase()) });
			}
			localResultSet = stmt.executeQuery(str);
			localResultSet.last();
			int i = localResultSet.getRow();
			int k = i;
			if (k > 0) {

				ColumnVo localColumnVo1 = new ColumnVo();
				if (CodeConfigProperties.DB_FILED_CONVERT) {
					localColumnVo1.setFieldName(e(localResultSet.getString(1).toLowerCase()));
				} else {
					localColumnVo1.setFieldName(localResultSet.getString(1).toLowerCase());
				}
				localColumnVo1.setFieldDbName(localResultSet.getString(1).toUpperCase());
				localColumnVo1.setPrecision(TableConvert.getNullString(localResultSet.getString(4)));
				localColumnVo1.setScale(TableConvert.getNullString(localResultSet.getString(5)));
				localColumnVo1.setCharmaxLength(TableConvert.getNullString(localResultSet.getString(6)));
				localColumnVo1.setNullable(TableConvert.getNullAble(localResultSet.getString(7)));

				localColumnVo1.setFieldType(a(localResultSet.getString(2).toLowerCase(), localColumnVo1.getPrecision(),
						localColumnVo1.getScale()));

				localColumnVo1.setFieldDbType(e(localResultSet.getString(2).toLowerCase()));

				a(localColumnVo1);
				localColumnVo1.setFiledComment(
						StringUtils.isBlank(localResultSet.getString(3)) ? localColumnVo1.getFieldName()
								: localResultSet.getString(3));

				log.debug("columnt.getFieldName() -------------" + localColumnVo1.getFieldName());
				localArrayList1.add(localColumnVo1);
				while (localResultSet.previous()) {
					ColumnVo localColumnVo2 = new ColumnVo();
					if (CodeConfigProperties.DB_FILED_CONVERT) {
						localColumnVo2.setFieldName(e(localResultSet.getString(1).toLowerCase()));
					} else {
						localColumnVo2.setFieldName(localResultSet.getString(1).toLowerCase());
					}
					localColumnVo2.setFieldDbName(localResultSet.getString(1).toUpperCase());
					localColumnVo2.setPrecision(TableConvert.getNullString(localResultSet.getString(4)));
					localColumnVo2.setScale(TableConvert.getNullString(localResultSet.getString(5)));
					localColumnVo2.setCharmaxLength(TableConvert.getNullString(localResultSet.getString(6)));
					localColumnVo2.setNullable(TableConvert.getNullAble(localResultSet.getString(7)));

					localColumnVo2.setFieldType(a(localResultSet.getString(2).toLowerCase(),
							localColumnVo2.getPrecision(), localColumnVo2.getScale()));

					localColumnVo2.setFieldDbType(e(localResultSet.getString(2).toLowerCase()));

					a(localColumnVo2);
					localColumnVo2.setFiledComment(
							StringUtils.isBlank(localResultSet.getString(3)) ? localColumnVo2.getFieldName()
									: localResultSet.getString(3));
					localArrayList1.add(localColumnVo2);
				}

			} else {

				throw new Exception("该表不存在或者表中没有字段");
			}
			log.debug("读取表成功");

			try {
				if (stmt != null) {

					stmt.close();
					stmt = null;
					System.gc();
				}
				if (conn != null) {
					conn.close();
					conn = null;
					System.gc();
				}

			} catch (SQLException e) {

				throw e;
			}
		} catch (ClassNotFoundException localClassNotFoundException) {

			throw localClassNotFoundException;
		} catch (SQLException localSQLException2) {

			throw localSQLException2;
		} finally {

			try {

				if (stmt != null) {

					stmt.close();
					stmt = null;
					System.gc();
				}
				if (conn != null) {
					conn.close();
					conn = null;
					System.gc();
				}

			} catch (SQLException e) {

				throw e;
			}
		}
		ArrayList<ColumnVo> localArrayList2 = new ArrayList<ColumnVo>();
		for (int j = localArrayList1.size() - 1; j >= 0; j--) {

			ColumnVo localColumnVo1 = localArrayList1.get(j);
			localArrayList2.add(localColumnVo1);
		}
		return localArrayList2;
	}

	public static boolean checkTableExist(String paramString) {
		String str = null;

		try {
			log.debug("数据库驱动: " + CodeConfigProperties.diverName);
			Class.forName(CodeConfigProperties.diverName);
			conn = DriverManager.getConnection(CodeConfigProperties.url, CodeConfigProperties.username,
					CodeConfigProperties.password);
			stmt = conn.createStatement(1005, 1007);
			if (CodeConfigProperties.databaseType.equals("mysql")) {
				str = "select column_name,data_type,column_comment,0,0 from information_schema.columns where table_name = '"
						+ paramString.toUpperCase() + "' and table_schema = '" + CodeConfigProperties.databaseName
						+ "'";
			}
			if (CodeConfigProperties.databaseType.equals("dm")) {
				str = "select atc.column_name, atc.data_type, acc.comments, atc.data_precision, atc.data_scale, atc.char_length, case atc.nullable when 'N' then 'NO' when 'Y' then 'YES' else '' end nullable "
						+ "from all_tab_columns atc left join all_col_comments acc on acc.column_name = atc.column_name where atc.table_name = '"
						+ paramString.toUpperCase() + "' and acc.table_name = '"+ paramString.toUpperCase() + "'and atc.owner = '" 
						+ CodeConfigProperties.databaseName + "' and acc.owner = '"+ CodeConfigProperties.databaseName +"'";
			}
			if (CodeConfigProperties.databaseType.equals("oracle")) {
				str = "select colstable.column_name column_name, colstable.data_type data_type, commentstable.comments column_comment from user_tab_cols colstable  inner join user_col_comments commentstable  on colstable.column_name = commentstable.column_name  where colstable.table_name = commentstable.table_name  and colstable.table_name = '"
						+ paramString.toUpperCase() + "'";
			}
			if (CodeConfigProperties.databaseType.equals("postgresql")) {
				str = MessageFormat.format(
						"SELECT base.attname AS  field,t.typname AS type,col_description(base.attrelid,base.attnum) as comment,null as column_precision,null as column_scale,null as Char_Length,base.attnotnull  FROM pg_class c,pg_attribute  base,pg_type t  WHERE c.relname = {0} and base.attnum > 0  and base.attrelid = c.oid and base.atttypid = t.oid  ORDER BY base.attnum ",
						new Object[] { TableConvert.getV(paramString.toLowerCase()) });
			}
			if (CodeConfigProperties.databaseType.equals("sqlserver")) {
				str = MessageFormat.format(
						"select distinct cast(base.name as varchar(50)) column_name,  cast(b.name as varchar(50)) data_type,  cast(e.value as varchar(200)) comment,  cast(ColumnProperty(base.object_id,base.Name,'''Precision''') as int) num_precision,  cast(ColumnProperty(base.object_id,base.Name,'''Scale''') as int) num_scale,  base.max_length,  (case when base.is_nullable=1 then '''y''' else '''n''' end) nullable,column_id   from sys.columns base left join sys.types b on base.user_type_id=b.user_type_id left join (select top 1 * from sys.objects where type = '''U''' and name ={0}  order by name) c on base.object_id=c.object_id left join sys.extended_properties e on e.major_id=c.object_id and e.minor_id=base.column_id and e.class=1 where c.name={0} order by base.column_id",
						new Object[] { TableConvert.getV(paramString.toLowerCase()) });
			}
			ResultSet localResultSet = stmt.executeQuery(str);
			localResultSet.last();
			int i = localResultSet.getRow();
			if (i > 0) {
				return true;
			}
		} catch (Exception localException) {

			localException.printStackTrace();
			return false;
		}
		return false;
	}

	private static String e(String paramString) {
		String[] arrayOfString = paramString.split("_");
		paramString = "";
		int i = 0;
		for (int j = arrayOfString.length; i < j; i++) {
			if (i > 0) {

				String str = arrayOfString[i].toLowerCase();

				str = str.substring(0, 1).toUpperCase() + str.substring(1, str.length());
				paramString = paramString + str;
			} else {

				paramString = paramString + arrayOfString[i].toLowerCase();
			}
		}
		return paramString;
	}

	public static String d(String paramString) {
		String[] arrayOfString = paramString.split("_");
		paramString = "";
		int i = 0;
		for (int j = arrayOfString.length; i < j; i++) {
			if (i > 0) {

				String str = arrayOfString[i].toLowerCase();

				str = str.substring(0, 1).toUpperCase() + str.substring(1, str.length());
				paramString = paramString + str;
			} else {

				paramString = paramString + arrayOfString[i].toLowerCase();
			}
		}
		paramString = paramString.substring(0, 1).toUpperCase() + paramString.substring(1);
		return paramString;
	}

	private static void a(ColumnVo paramColumnVo) {
		String str1 = paramColumnVo.getFieldType();
		String str2 = paramColumnVo.getScale();

		paramColumnVo.setClassType("inputxt");
		if ("N".equals(paramColumnVo.getNullable())) {
			paramColumnVo.setOptionType("*");
		}
		if ("datetime".equals(str1) || str1.contains("time")) {
			paramColumnVo.setClassType("easyui-datetimebox");
		} else if ("date".equals(str1)) {
			paramColumnVo.setClassType("easyui-datebox");
		} else if (str1.contains("int")) {
			paramColumnVo.setOptionType("n");
		} else if ("number".equals(str1)) {

			if (StringUtils.isNotBlank(str2) && Integer.parseInt(str2) > 0) {
				paramColumnVo.setOptionType("d");
			}
		} else if ("float".equals(str1) || "double".equals(str1) || "decimal".equals(str1)) {
			paramColumnVo.setOptionType("d");
		} else if ("numeric".equals(str1)) {
			paramColumnVo.setOptionType("d");
		}
	}

	private static String a(String paramString1, String paramString2, String paramString3) {
		if (paramString1.contains("char")) {
			paramString1 = "java.lang.String";
		} else if (paramString1.contains("bigint")) {
			paramString1 = "java.lang.Long";
		} else if (paramString1.contains("int")) {
			paramString1 = "java.lang.Integer";
		} else if (paramString1.contains("float")) {
			paramString1 = "java.lang.Float";
		} else if (paramString1.contains("double")) {
			paramString1 = "java.lang.Double";
		} else if (paramString1.contains("number")) {

			if (StringUtils.isNotBlank(paramString3) && Integer.parseInt(paramString3) > 0) {
				paramString1 = "java.math.BigDecimal";
			} else if (StringUtils.isNotBlank(paramString2) && Integer.parseInt(paramString2) > 10) {
				paramString1 = "java.lang.Long";
			} else {
				paramString1 = "java.lang.Integer";
			}

		} else if (paramString1.contains("decimal")) {
			paramString1 = "java.math.BigDecimal";
		} else if (paramString1.contains("date")) {
			paramString1 = "java.util.Date";
		} else if (paramString1.contains("time")) {
			paramString1 = "java.util.Date";
		} else if (paramString1.contains("blob")) {
			paramString1 = "byte[]";
		} else if (paramString1.contains("clob")) {
			paramString1 = "java.sql.Clob";
		} else if (paramString1.contains("numeric")) {
			paramString1 = "java.math.BigDecimal";
		} else if (paramString1.contains("text")) {
			paramString1 = "java.lang.String";
		} else {
			paramString1 = "java.lang.Object";
		}
		return paramString1;
	}

	public static void main(String[] args) throws SQLException {
		try {
			List<ColumnVo> localList = readTableColumn("demo");
			for (ColumnVo localColumnVo : localList) {
				System.out.println(localColumnVo.getFieldName());
			}
		} catch (Exception localException1) {
			localException1.printStackTrace();
		}
		new DbReadTableUtil();
		System.out.println(ArrayUtils.toString(readAllTableNames()));
	}
}