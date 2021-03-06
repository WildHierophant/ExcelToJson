# ExcelToJson
Excel表格转换为Json的工具

## 使用说明：
1.工具读取表格的表格需要修改Sheet1表格页面的名称，不能以Sheet或sheet开头，也不能为空，否则无法导出  
2.表格填写格式为，第一行不读，可作为备注行，第二行为表头，不能为空，第三行为单元格类型，不能为空，且单元格的类型必须与第三行开始往后的数据类型一致  
3.工具会读取表格地址目录下，除子文件夹目录外的的所有文件xlsx与xls格式文件，并将读取到的所有表格导出为Json格式文件  

## 使用步骤：
1.填入需要转换的表格所在地址  
2.填入导出目录地址后  
3.点击导出Json按钮即可导出Json文件  

## 效果图
![image](https://github.com/WildHierophant/ExcelToJson/blob/master/2020528-185333.jpg)
![image](https://github.com/WildHierophant/ExcelToJson/blob/master/2020528-185505.jpg)
![image](https://github.com/WildHierophant/ExcelToJson/blob/master/2020528-185603.jpg)

Json转表工具.rar，压缩包中为编译好的文件，ExcelToJson文件夹为原始工程文件
