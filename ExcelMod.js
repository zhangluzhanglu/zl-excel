'use strict';
// 模块底层调用的exceljs模块，文档：https://github.com/exceljs/exceljs/blob/master/README_zh.md
const Excel = require('exceljs');
class ExcelMod {
    /**
          * @function 生成表格对象的功能函数
          * @description 接收一个数组或对象，生成含有一个或者多个sheet的Excel表格
          * @param frameInfo {Object} 表示此表格将在何种框架中使用。如：{ name:"egg.js",filename:"测试表格",other:ctx}
          * @param frameInfo.name {Object} 表示此模块在哪个框架信息中使用
          * @param frameInfo.name {String} 导出的表格文件名
          * @param frameInfo.other {Object} 此框架需要的其他信息对象集合
          * @param params {Object|Array} 对象或者对象数组，传入对象表示只创建一个sheet,对象数组表示创建多个sheet
          * @return workbook表格对象
          * @author 张路 2021/08/09
          * @example
          * 参数一：框架信息【可选】
          * { name:"egg.js",filename:"测试表格",other:ctx}
          *
          * 参数二：表格数据信息【必填】
          * [
          *    {
          *       sheetName: '用户表', //第一个sheet的名字
          *       columns: [
          *                 { header: '姓名', key: 'name', width: 15 },
          *                 { header: '性别', key: 'sex', width: 10 },
          *                 { header: '年龄', key: 'age', width: 20, default: 0 },
          *                 { header: '爱好', key: 'hobby', width: 15, default: 0 },
          *            ],
          *        rows: [
          *                { name: "张三", sex: "男", age: 18, hobby: "小说、音乐"},
          *                { name: "李四", sex: "女", age: 19, hobby: "小说、音乐、学习"}
          *             ]
          *     }
          *  ]
        */
    static async getWorkbook(frameInfo, params) {
        const workbook = new Excel.Workbook();
        workbook.creator = 'zl-table';
        workbook.created = new Date(Date.now());
        workbook.modified = new Date(Date.now());
        // 传入多个sheet对象构成的数组
        let options;
        if (Object.prototype.toString.call(params) == '[object Object]') {
            options = [params]; // 表示传入的对象
        } else if (Array.isArray(params)) {
            options = params; // 表示传入的数组
        } else {
            throw new Error('excel配置参数格式错误');
        }
        // 配置默认的表格相关样式
        const defaultStyle = {
            alignment: {
                vertical: 'center',
                horizontal: 'center',
            },
        };
        // 循环创建表格里面的各个sheet
        for (let i = 0; i < options.length; i++) {
            const obj = options[i];
            const { sheetName = '', columns = [], rows = [] } = obj;
            const sheet = workbook.addWorksheet(sheetName);
            sheet.columns = columns.map(item => {
                const style = item.style ? Object.assign(item.style, defaultStyle) : defaultStyle;
                return {
                    ...item,
                    style,
                };
            });
            // 构建默认表头
            const defualtRowval = {};
            columns.forEach(col => {
                // 如果构建表头时传入了默认值col.default，那么就使用默认的值构建默认的表头对象
                // 这可以解决这个表头字段，在数据行中没有对应值时，可以取默认值进行渲染
                if (col.default != undefined && col.key != undefined) {
                    defualtRowval[col.key] = col.default;
                }
            });
            // 添加具体的表体行数据
            rows.forEach((row, index) => {
                for (const key in row) {
                    if (row[key] == undefined) {
                        row[key] = defualtRowval[key];
                    }
                }
                const rowData = {
                    ...row,
                    index: index + 1, // 每行的索引
                };
                sheet.addRow(rowData);
            });
        }
        // console.log('==========frameInfo', frameInfo);
        // 如果没有传入任何的框架信息，那么就直接直接返回Excel对象
        if (!frameInfo) {
            return workbook;
        }
        // egg.js框架
        else if (frameInfo.name == 'egg.js') {
            const { other, filename } = frameInfo;
            if (!other || !filename) {
                throw frameInfo.name + '必须传入ctx和filename';
            }
            const ctx = other;
            // 表格信息请求头
            ctx.set('content-disposition', `attachment; filename* = UTF-8''${encodeURIComponent(filename)}`);
            ctx.status = 200;
            await workbook.xlsx.write(ctx.res);
            ctx.res.end();
        }
        // 其他
        else {
            console.log('暂不支持出egg.js之外的其他框架自动构建导出');
        }
    }
}

module.exports = ExcelMod;
