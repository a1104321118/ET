# ET
```
ET框架------->excel-tool,用于快速处理Excle业务，基于反射原理

设计概念来自于 数据库的思想

//读数据
ExcelTool excelTool = new ExcelTool(filePath);
List<Person> peopleList = excelTool.readData(Person.class);

//写数据
excelTool.writeData(peopleList, outFilePath);

so easy

功能暂时还很有限，欢迎大家讨论。
blog:
https://my.oschina.net/u/3582320/blog/1501527
```
