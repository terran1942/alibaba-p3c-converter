# alibaba-p3c-converter
阿里巴巴代码检查插件结果转换器。 可以将 [alibaba java coding guidelines](https://github.com/alibaba/p3c) 中
[idea plugin](https://github.com/alibaba/p3c/tree/master/idea-plugin) 的导出结果转换为xlsx格式

## 运行环境
+ Python 3.8

## 所需模块
+ bs4
+ openpyxl

## 使用说明
+ template.xlsx 配置导出后xlsx的标题行
+ nexus.xlsx 对应关系xlsx
  + code 配置项目代号与名称、团队的对应关系
  + alias 配置别名与项目代号的对应关系
  
1. 使用alibaba java coding guidelines的idea plugin将扫描结果导出为xml
2. 在xml_reporting下以工程名建立新目录，将导出的xml复制到其中
3. 运行main.py，导出结果放置于xlsx_reporting之下
  
