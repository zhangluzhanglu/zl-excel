const ExcelMod = require("ExcelMod");

// egg.js中的controller代码

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
