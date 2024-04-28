# 日常小工具-业绩证明材料生成

# 0x1 功能描述

	脚本是一个将Excel数据转换为Word文档的Python程序。它使用了pandas、python-docx和json库来实现这个功能。主要功能是将一个包含工作业绩信息的Excel文件转换为一个Word文档，并在Word文档中添加一些格式化的内容，如标题、段落、表格等。同时，它还可以从配置文件中读取个人信息，并将这些信息添加到生成的Word文档中。

# 0x2 配置文件

	配置文件放置在脚本同目录，文件名需要为config.json

```python
{
    "name" : "技术负责人的名字",
    "tel" : "技术负责人的电话",
    "title" : "高级工程师",
    "work" : "xx主任",
    "depart" : "xxxxx部门",
    "company" : "xxxx公司",
    "myname" : "开证明的人的名字" 
}
```
# 0x3 Excel 格式

只需要完善标黄的内容，时间，和角色即可，对应的证明文件会自动生成。

![image](assets/image-20240428215820-e4r36nt.png)​

