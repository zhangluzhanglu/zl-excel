# zl-excel
   在进行node开发时，快速生成excel表格

## 1. 起因

   在开发中，经常涉及到excel表格导出的需求，在每个项目里面写时，都要把其他项目代码复制一遍到新项目。
然后在调整相关代码...感觉很繁琐，而且也不好维护，所以这里我就干脆写个通用的npm模块：zl-excel 实现此功能
这样以后在新的项目中只需要引入此模块即可使用，简单方便。


## 2. 安装模块

* ***使用`require`方式在nodejs中引入使用***
   ```js
 
       1. 安装： npm i  zl-excel -S

       2. 引入： const Excel = require("zl-excel")
   ```

## 3. 使用示例（目前主要是在egg.js中使用）

```js
// =========下面代码在egg.js的路由里面执行，需要获取到ctx===========
    const { ctx } = this;
    // 框架信息
    const frameInfo = { name: 'egg.js', filename: '测试表格.xlsx', other: ctx };
    // 表格数据信息
    const tableInfo = [
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
    // 导出表格
    await ExcelMod.getWorkbook(frameInfo, tableInfo);
```

## 4. 提示

   此模块底层调用的exceljs模块，文档：https://github.com/exceljs/exceljs/blob/master/README_zh.md

## 5. 后续
   后续空了会依次添加主流node框架的直接支持，代码里面也预留了相关接口。
