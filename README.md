# 处理Excel的小工具，使用Go语言实现

## 项目背景
今天碰到了一个场景，系统中有一个功能，需要上传一个Excel。一个Excel中只有一个表单。

但是业务方给我一个大的Excel，里面把所有的sheet全部放到一起了，但是我系统中，只能接受一个单一的表单。

于是我就做了一个小工具，你只需要给我一个Excel，就可以按照不同的表单名字进行拆分，也可以按照表头进行拆分


### Split模式  [excel-splitter-split.exe](https://github.com/langkemaoxin/utils-excel-go/releases/download/1.0/excel-splitter-merge.exe)
使用方式：把这个exe程序，放入到包含有excel文件的目录中，双击一下，就能把当前的这个excel，根据不同的表单，进行分割输出。

### Merge模式 [excel-splitter-merge.exe](https://github.com/langkemaoxin/utils-excel-go/releases/download/1.0/excel-splitter-split.exe)
使用方式：把这个exe程序，放入到包含有excel文件的目录中，双击一下，就能把当前的这个excel，根据不同的表单，进行分割输出。在此基础上，
如果发现表头的列是一致的，则会自动的吧表头一致的输出到同一个文件中。



