## 功能介绍

采用古老的js/vbs的ASP代码，用于临时使用，本人较为喜欢JS所以index.asp为 jscript，但是有个上传组件使用的是vbs所以后端的cb目录用vbs

1. 用于临时转存文件和文本
2. 如下图
   1. 上传为转存文件，可以支持多个文件
   2. 提交为转存文本

![image-20240326214216020](screenshot/image-20240326214216020.png)

## 使用方法

1. 需要windows的IIS服务和ASP功能,由于需要上传,所以组件需要如下

   1.  Scripting.FileSystemObject (FSO)
   2.  Scripting.Dictionary 
   3. ADODB.Recordset
   4. Adodb.Stream

2. 支持相对目录,文件夹之间互不影响,可以建立文件夹,比如 tmp

   ```
   /tmp/
   ```

3. 配置IIS的访问入口默认为 index.asp
4. 仓库的 src 目录中的代码复制过去即可使用
5. 打开浏览器,通过 http://网址/目录/ 来实现访问
6. 如果要修改地址和上传后缀控制，请修改 index.asp中的

  ```
// 临时文件路径 
var filePath = './temp/'
// 上传后缀控制
var extList = "|txt|ini|md|mp3|m4a|bmp|jpg|jpeg|png|zip|7z|rar|pdf|doc|docx|xls|xlsx|ppt|pptx|epub|apk"
  ```