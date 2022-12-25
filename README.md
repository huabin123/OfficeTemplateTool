# todo list

* [ ]  word模版解析
* [ ]  base-log
* [ ]  ratelimiter
* [ ]  大数据量导入导出excel

# OfficeTemplateTool

**OfficeTemplateTool**是一个用来处理Office文件的工具，包含word和excel的模版处理。

## Word的解析

[使用传送门](/Users/huabin/workspace/playground/工具向/OfficeTemplateTool/word/Readme.md)

实现的功能：

* [ ]  通过配置表从外部读取数据，返回单条数据的称为单维指标，返回多条数据（列表嵌套列表）的称为多维指标
* [ ]  全局单维指标的替换，包含页眉、段落、表格
* [ ]  表格中多维指标的填充，包含：
  * [ ]  1、单行填充
  * [ ]  2、多行填充，多行填充时需要在表格末尾添加新行
* [ ]  段落和表格跨页的识别和特殊处理
* [ ]  中英文字体的全局设置
