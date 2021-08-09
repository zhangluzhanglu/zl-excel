# zl-excel

  作用：在进行node开发时，快速生成excel表格。
   * 可将表格导出到服务器本地
   * 也可将表格直接响应到前端

## 1. 起因

   在开发中，经常涉及到excel表格导出的需求，在每个新项目里面写时，都要把其他项目代码复制一遍到新项目...感觉很繁琐，而且也不好维护，
   所以这里我就干脆写个通用的npm模块：zl-excel 实现此功能，这样以后在新的项目中只需要引入此模块即可使用，简单方便。

## 2. 安装模块

* ***使用`require`方式在nodejs中引入使用***
   ```js
 
       1. 安装： npm i  zl-excel -S

       2. 引入： const ExcelMod = require("zl-excel")
   ```

## 3. 使用示例--在服务器本地创建excel表格

```js
    // 表格数据信息
    const excelInfo = [
      {
        sheetName: '用户表', // 第一个sheet的名字
        columns: [
          { header: '姓名', key: 'name', width: 15 },
          { header: '性别', key: 'sex', width: 10 },
          { header: '年龄', key: 'age', width: 20, default: 0 },
          { header: '爱好', key: 'hobby', width: 15, default: 0 },
          { header: '成功率', key: 'sucPre', width: 15, default: 0, style: { numFmt: '0.00%' } },
        ],
        rows: [
          { name: '张三', sex: '男', age: 18, hobby: '小说、音乐', sucPre: '0.6789' },
          { name: '李四', sex: '女', age: 19, hobby: '小说、音乐、学习', sucPre: '0.8888' },
        ],
      },
      {
        sheetName: '书籍表', // 第二个sheet的名字
        columns: [
          { header: '书名', key: 'bookName', width: 20, default: 0 },
          { header: '出版社', key: 'publish', width: 15, default: 0 },
          { header: '内容摘要', key: 'content', width: 15, default: 0 },
        ],
        rows: [
          { bookName: '《经济学》', publish: '北大出版社', content: '适合经济学学入门的书籍' 
          },
          { bookName: '《数学概览》', publish: '清华出版社', content: '适合现代数学入门的书籍' 
          },
        ],
      },
    ];

    // 自动导出表格到本地
    await ExcelMod.getWorkbook({
      excelPath: './test表格66.xlsx',
      excelData: excelInfo,
    });

```
```js
   //------如果需要对表格进行复杂的操作，这里可传入参数returnCall: true,从而将表格返回到调用处,然后进行你想要的修改，最后在commit创建表格 -----

    const workbook = await ExcelMod.getWorkbook({ // 手动执行commit,导出表格
      excelPath: './test表格66.xlsx',
      excelData: excelInfo,
      returnCall: true,
    });
    // 。。。对表格进行复杂的处理。。。

    // 提交更改，创建表格
    workbook.commit();
```

## 4. 使用示例--在egg.js中使用

```js
// =========下面代码在egg.js的路由里面执行，需要获取到ctx===========
    const { ctx } = this;
    // 框架信息
    const eggFrameInfo = { name: 'egg.js', filename: '测试表格.xlsx', other: ctx };
    // 表格数据信息
    const excelInfo = [
      {
        sheetName: '用户表', // 第一个sheet的名字
        columns: [
          { header: '姓名', key: 'name', width: 15 },
          { header: '性别', key: 'sex', width: 10 },
          { header: '年龄', key: 'age', width: 20, default: 0 },
          { header: '爱好', key: 'hobby', width: 15, default: 0 },
          { header: '成功率', key: 'sucPre', width: 15, default: 0, style: { numFmt: '0.00%' } },
        ],
        rows: [
          { name: '张三', sex: '男', age: 18, hobby: '小说、音乐', sucPre: '0.6789' },
          { name: '李四', sex: '女', age: 19, hobby: '小说、音乐、学习', sucPre: '0.8888' },
        ],
      },
      {
        sheetName: '书籍表', // 第二个sheet的名字
        columns: [
          { header: '书名', key: 'bookName', width: 20, default: 0 },
          { header: '出版社', key: 'publish', width: 15, default: 0 },
          { header: '内容摘要', key: 'content', width: 15, default: 0 },
        ],
        rows: [
          { bookName: '《经济学》', publish: '北大出版社', content: '适合经济学学入门的书籍' },
          { bookName: '《数学概览》', publish: '清华出版社', content: '适合现代数学入门的书籍' },
        ],
      },
    ];
    // 自动导出表格到前端
    await ExcelMod.getWorkbook({
      excelData: excelInfo,
      frameInfo: eggFrameInfo,
    });

```
```js
   //------如果需要对表格进行复杂的操作，这里不传入框架信息即可，它会将表格返回到调用处,然后进行你想要的修改，最后在发送到前端 -----
   
   //将表格返回到调用处    
    const workbook = await ExcelMod.getWorkbook({ 
      excelData: excelInfo,
    });

    // 。。。对表格进行复杂的处理。。。

    // 设置表格信息请求头，然后进行导出到前端
    const fileName = encodeURIComponent('测试表格.xlsx');
    ctx.set('content-disposition', `attachment; filename* = UTF-8''${fileName}`);
    ctx.status = 200;
    await workbook.xlsx.write(ctx.res);
    ctx.res.end();
```

## 4. 提示

   此模块底层调用的exceljs模块，文档：https://github.com/exceljs/exceljs/blob/master/README_zh.md

## 5. 后续
   后续空了会依次添加主流node框架的直接支持，客户端js等导出，代码里面也预留了相关接口。
