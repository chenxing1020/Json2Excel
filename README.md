# JSON2EXCEL

&emsp;&emsp;事情的起因，今天一个商学的同学让我帮忙搞个数据，她的addidas neo的德国合作商想让她帮忙整理一下从微博API上爬下来的数据。我一看是JSON格式，那用nodejs处理真的是再合适不过了，而且导出也不是很难，借用了npm的`exceljs`的包。  
&emsp;&emsp;把数据导出来看，大概就是一些运动厂商的官博发的一些转发微博，爬取了微博的评论，然后想了解一下用户的信息（数据我只放了一个仓库里）。  

```javascript
var Excel = require('exceljs'),
    fs=require('fs');

var workbook = new Excel.stream.xlsx.WorkbookWriter({
  filename: './streamed-workbook.xlsx'
});
var worksheet = workbook.addWorksheet('Sheet');

worksheet.columns = [
  { header: 'created_at',key:'created_at'},
  { header: 'user_text',key:'user_text'},
  ...
];

for (var ii=1;ii<181;ii++){
  var file=ii.toString()+'.json';
  var result=JSON.parse(fs.readFileSync(file));
  var comments=result.comments;
  var s_t=result.status;
  var us_t=result.status.user;
  for (var i in comments){
    let u_t=comments[i].user;

    let comm={
      created_at:comments[i].created_at,
      user_text:comments[i].text,
      ...
    }
    worksheet.addRow(comm).commit();
  }
}
workbook.commit();
```

&emsp;&emsp;然后中途还发生了一件我觉得还蛮意思的事，因为她发来的数据文件大概有180多个，而且名字都是特别奇葩的，所以用excel批量改了一下文件名，流程大概如下：

    从控制台进入目标文件夹，然后输入:

```shell
dir/b>1.xls
```

    这样在目标文件夹下会生成一个名为1.xls的文件夹，然后在excel中编辑想要更改的名称如下图：

![image](./1.png)

    改完之后，将excel中的内容复制到txt文本中，将中间的tab全部替换成空格，修改文件名为bat，实现批处理即可。