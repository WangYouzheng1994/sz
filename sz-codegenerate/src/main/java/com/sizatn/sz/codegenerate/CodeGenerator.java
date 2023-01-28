package com.sizatn.sz.codegenerate;

import com.sizatn.sz.codegenerate.generate.impl.CodeGenerateOne;
import com.sizatn.sz.codegenerate.generate.pojo.TableVo;

public class CodeGenerator {
	
	public static void main(String[] args) throws Exception {
		System.out.println("-------- Code------------- Generation -----[单表模型]------- 生成中。。。");
		TableVo vo = new TableVo();
		vo.setEntityName("DemoHolidayRequest");//类名称首字母大写
        vo.setEntityPackage("system.demo");//包名称/微服务名称
        vo.setTableName("test_user");//表名称
        vo.setPrimaryKeyPolicy("id");//表主键
        vo.setFtlDescription("");//实体注释
		(new CodeGenerateOne(vo)).generateCodeFile();
		System.out.println("-------- Code------------- Generation -----[单表模型]------- 生成完成。。。");
	}
}
