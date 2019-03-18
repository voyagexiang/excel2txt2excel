const xlsx = require('node-xlsx');
const fs = require('fs')

// 读取文件，处理文件
function readExcel(path) {
    let sheets = xlsx.parse('./public/' + path); //获取到所有sheets
    let excel = []
    let title = ["界址点号", "纵坐标（x）", "横坐标（y）", "反算边长（米）", "地块圈号", "备注"]
    excel.push(title)
	
	// txt创建
    let txtOut = fs.createWriteStream('./output/' + path.split('.')[0] + '.txt', {
        encoding: 'utf8'
    });
    txtOut.write(title.join('   '))


    let excelSet = new Set()
    sheets.forEach(function(sheet) {
        for (let rowId in sheet['data']) {
            let row = sheet['data'][rowId];
            if (sheet['data'][rowId].length == 4 && !(row[3] == 'y(m)')) {               
                if (row[1] == 'J1' || !excelSet.has(row.slice(2,4).join('$'))) {                	
                    excelSet.add(row.slice(2,4).join('$'))
                    row.shift()
                    row[4] = 1
                    excel.push(row)

                    // 写入txt文件中
                    txtOut.write('\r\n')
                    txtOut.write(row.join('      '))

                }

            }
        }
    });
    // txt输入完成
    txtOut.end();

    // 写入excel文件中
    let buffer = xlsx.build([{ name: sheets[0]['name'], data: excel }]);
    fs.writeFileSync('./output/' + path.split('.')[0] + '.xls', buffer, 'binary')
}


// 生成excel文件
function readBuildExcel(obj) {
    for (let file in obj) {
        let fileName = obj[file]['filename']
        readExcel(fileName)
    }
}

// 获取文件夹中的文件
function readFileList(path, filesList) {
    var files = fs.readdirSync(path);
    files.forEach(function(itm, index) {
        var stat = fs.statSync(path + itm);
        if (stat.isDirectory()) {
            //递归读取文件
            readFileList(path + itm + "/", filesList)
        } else {

            var obj = {}; //定义一个对象存放文件的路径和名字
            obj.path = path; //路径
            obj.filename = itm //名字
            filesList.push(obj);
        }

    })
}
var getFiles = {
    //获取文件夹下的所有文件
    getFileList: function(path) {
        var filesList = [];
        readFileList(path, filesList);
        return filesList;
    }
};

//获取文件夹下的所有文件
const filelist = getFiles.getFileList("./public/")
console.log(filelist)
readBuildExcel(filelist)