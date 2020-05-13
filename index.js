var nodeExcel = require('excel-export');
var fs = require('fs')
var path = require('path')
var conf = {};
let destPath = '/result.xlsx'
var writeStream = fs.createWriteStream(path.join(__dirname, destPath))
conf.stylesXmlFile = 'styles.xml';
// 头部
conf.cols = [
  {
    caption: 'id',
    type: 'string',
    beforeCellWrite: function(row, cellData) {
      return cellData.toUpperCase();
    },
    width: 28.7109375
  },
  {
    caption: '创建日期',
    type: 'date',
    beforeCellWrite: (function() {
      var originDate = new Date(Date.UTC(1899, 11, 30));
      return function(row, cellData, eOpt) {
        if (eOpt.rowNum % 2) {
          eOpt.styleIndex = 1;
        } else {
          eOpt.styleIndex = 2;
        }
        if (cellData === null) {
          eOpt.cellType = 'string';
          return '0000-00-00';
        } else return (cellData - originDate) / (24 * 60 * 60 * 1000);
      };
    })()
  },
  {
    caption: 'bool',
    type: 'bool'
  },
  {
    caption: 'number',
    type: 'number'
  }
];
// 数据
conf.rows = [
  ['pi', new Date(Date.UTC(2013, 4, 1)), true, 3.14],
  ['e', new Date(2012, 4, 1), false, 2.7182],
  ["M&M<>'", new Date(Date.UTC(2013, 6, 9)), false, 1.61803],
  ['null date', null, true, 1.414]
];
var result = nodeExcel.execute(conf);
writeStream.write(result, 'binary');
writeStream.end()