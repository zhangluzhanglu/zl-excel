'use strict';
// 模块底层调用的exceljs模块，文档：https://github.com/exceljs/exceljs/blob/master/README_zh.md
const Excel = require('exceljs');
class ExcelMod {
  /**
      * @function 生成表格对象的功能函数
      * @description 接收表格相关信息，生成含有一个或者多个sheet的Excel表格
      *
      * @param paramsObj {Object} 完整的表格参数信息
      * @param paramsObj.excelData {Object|Array} 生成的excel表格数据，为对象或者对象数组，传入对象表示只创建一个sheet,对象数组表示创建多个sheet
      * @param paramsObj.cellDefaultStyle {Object} 每个单元格的默认样式（如果构建表头数据时，未传入此对象，就会以此默认样式为准备），默认值:{alignment: { wrapText: true, vertical: 'middle', horizontal: 'center' }}
      * @param paramsObj.excelPath {String} 生成的excel表格路径
      * @param paramsObj.frameInfo {Object} 表示此表格将在何种框架中使用。如：{ name:"egg.js",filename:"测试表格",other:ctx}
      * @param paramsObj.frameInfo.filename {String} 表示生成的ecel名字
      * @param paramsObj.frameInfo.name {String} 表示此模块在哪个框架信息中使用
      * @param paramsObj.frameInfo.other {Object} 此框架需要的其他信息对象集合
      * @param paramsObj.returnCall {Boolean} 将表格对象返回到调用处,默认为false即不返回
      * @return workbook表格对象
      * @author zl-fire 2021/08/09
      * @example
      * {
      *    // 表格数据信息【必填】
      *    excelData:[
      *     {
      *       sheetName: '用户表', //第一个sheet的名字
      *       columns: [
      *        { header: '姓名', key: 'name', width: 15 },
      *        { header: '性别', key: 'sex', width: 10 },
      *        { header: '年龄', key: 'age', width: 20, default: 0 },
      *        { header: '爱好', key: 'hobby', width: 15, default: 0 },
      *       ],
      *       rows: [
      *        { name: "张三", sex: "男", age: 18, hobby: "小说、音乐"},
      *        { name: "李四", sex: "女", age: 19, hobby: "小说、音乐、学习"}
      *       ]
      *     }
      *   ],
      *   //框架信息【可选】
      *   frameInfo:{ name:"egg.js",filename:"测试表格",other:ctx},
      * }
    */
  static async getWorkbook(paramsObj) {
    const { excelPath, excelData, frameInfo, returnCall = false, cellDefaultStyle } = paramsObj;
    let workbook;
    // 传了表格路径，则表示 表格将生成在本地硬盘
    if (excelPath) {
      workbook = new Excel.stream.xlsx.WorkbookWriter({
        filename: excelPath,
        useStyles: true,  //设置允许写入样式信息，默认为了节省性能时不允许的
      });
    }
    // 其他表示通过路由返回到前端
    else {
      workbook = new Excel.Workbook();
    }
    // 具体的表格信息
    workbook.creator = 'zl-excel';
    workbook.created = new Date(Date.now());
    workbook.modified = new Date(Date.now());
    // 传入多个sheet对象构成的数组
    let options;
    if (Object.prototype.toString.call(excelData) == '[object Object]') {
      options = [excelData]; // 表示传入的对象
    } else if (Array.isArray(excelData)) {
      options = excelData; // 表示传入的数组
    } else {
      throw new Error('excel配置参数格式错误');
    }
    // 配置默认的表格相关样式
    let defaultStyle;
    if (!cellDefaultStyle) {
      defaultStyle = {
        alignment: { wrapText: true, vertical: 'middle', horizontal: 'center' },
        // font: { bold: true },
      };
    }
    else defaultStyle = cellDefaultStyle;
    // 循环创建表格里面的各个sheet
    for (let i = 0; i < options.length; i++) {
      const obj = options[i];
      const { sheetName = '', columns = [], rows = [] } = obj;
      const sheet = workbook.addWorksheet(sheetName);
      // 设置列的样式
      columns.forEach(item => {
        item.style = item.style || defaultStyle;
      });
      sheet.columns = columns;

      /*
        处理表头，得到真正的表头
        逻辑：
        1. 表头数组有几个数组元素就表示表头占据几行，每个数组元素表示一行
        2. colspan: 3 表示此单元格跨几列
        3. rowspan: 2 表示此单元格跨几行
  
        */



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
          // 当值为null,undefined,"",o 等值时,且用户又设置了默认值时，就取默认值
          if (!row[key] && defualtRowval[key] != undefined) {
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
    // ==================表格生成完毕，开始处理==================
    // 如果传了表格路径，则表示 表格将生成在本地硬盘
    if (excelPath) {
      // 直接生成表格，不返回到调用处
      if (!returnCall) {
        workbook.commit();
        return;
      }
      // 将表格信息返回到调用处，供二次加工处理，然后手动调用commit();
      if (returnCall) {
        return workbook;
      }
    }
    // 如果传了框架信息，那么必然是直接将表格通过路由响应到前端
    else if (frameInfo) {
      // egg.js框架
      if (frameInfo.name == 'egg.js') {
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
    // 如果没传路径，且没传框架信息 ，那么就返回Excel对象到调用处
    else {
      return workbook;
    }
  }
}

module.exports = ExcelMod;
