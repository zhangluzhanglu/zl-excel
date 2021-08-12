let columns = [
    [
        { header: '基本信息', key: '', width: 15, colspan: 3 }, // 跨3列
        { header: '成功率', key: 'sucPre', width: 15, rowspan: 2 }, // 跨2行
        { header: '测试字段', key: 'test', width: 15, rowspan: 2 }, // 跨2行
    ],
    [
        { header: '姓名', key: 'name', width: 15 },
        { header: '性别', key: 'sex', width: 10, default: '默认性别' },
        { header: '年龄', key: 'age', width: 20, default: 0 },
    ],
];

// 遍历表头数组的每一个数组元素
for (let i = 0; i < columns.length; i++) {
    let arr = columns[i];

    for (let j = 0; j < arr.length; j++) {
        let obj = arr[j]; //每个对象
        let { colspan, rowspan } = obj;
        //----计算跨列---------
        if (colspan) {
            let n = 1;
            while (colspan > n) {
                arr.splice(j + n, 0, { ...obj, colspan: 1, rowspan: 1 });

                n++;
            }
            //     obj.colspan=1;
        }

        // //------计算跨行-------

        // 如果只存在跨行
        if (rowspan && !colspan) {
            let m = 1;
            while (rowspan > m) {
                columns[i + m][j] = { ...obj, colspan: 1, rowspan: 1 };
                m++;
            }
            //    obj.rowspan=1;
        }
        //如果同时存在跨行+跨列
        if (rowspan && colspan) {
            for (let z = 0; z < colspan; z++) {
                let n = 1;
                while (colspan > n) {
                    arr.splice(j + n, 0, { ...obj, colspan: 1, rowspan: 1 });

                    n++;
                }
                //     obj.rowspan=1;
            }

        }
    }
}

//  明天先看下，为啥样式不生效，本地调试下，然后在处理node+ts问题
//  剩下的逻辑：在表格生成完后，在遍历一次处理后的表头数组，实现excel的跨列，跨行合并
//  不过这里需要写个函数，传入一个数字，从0开始，能够自动计算出A,B,C等对应关系，方便合并操作
//  然后在处理行数据的合并操作