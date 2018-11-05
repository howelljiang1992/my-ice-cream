const express = require('express');
const request = require('request');
const iconv = require('iconv-lite');
const XLSX = require('xlsx');
let app = express();

app.get('/', async function (req, res) {

  const data = await handleData();

  let sheets = {};
  let needCompanyList = extratExcel();
  console.log(needCompanyList.length);

  Object.keys(data).forEach((sheetName)=> {
    let sheetList = data[sheetName].list.map((item)=>{
       return [
        item.industry,
        item.industryrate,
        item.Pricelimit,
        item.stockNumber,
        item.lootingchips,
        item.Scramble,
        item.rscramble,
        item.Strongstock
       ]
    });
    
    let filterList = [];
    // 过滤sheetList里需要的公司
    needCompanyList.forEach((companyCode)=>{
      sheetList.forEach((item)=>{
        let code = item[0].match(/[^\(\)]+(?=\))/g);
        if (code[0] === companyCode){
          filterList.push(item);
        }
      });
    });

    let sheetTile = [['股票名称/代码','总得分','等级','股东责任','员工责任','供应商、客户和消费者责任','环境责任','社会责任']];
    filterList = sheetTile.concat(filterList);
    sheets[sheetName] = XLSX.utils.aoa_to_sheet(filterList);
  });

  let fileBuffer = XLSX.write({
    Sheets: sheets,
    SheetNames: Object.keys(sheets)
  }, { type: 'buffer' });

  res.set({
    'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'Content-Disposition': 'attachment;filename=' + encodeURIComponent('爬取数据.xlsx'),
    'Pragma': 'no-cache',
    'Expires': 0
  }).send(fileBuffer);
});

// 数据请求封装
function getData(url) {
  return new Promise((resolve, reject) => {
    request({ url, encoding: null }, function (error, response, body) {
      if (error) reject(error);

      let buf = new Buffer(body.slice(13, -1));
      let result = iconv.decode(buf, 'gb2312'); // 汉字不乱码
      let testJson = eval("(" + result + ")");
      resolve(testJson);
    });
  });
}

// 处理数据, 按页、按年
async function handleData() {
  let dataList = {};
  let year = 2016;
  for (year; year <= 2016; year++) {
    let page = 1;
    let oneYearList = [];
    for (page; page <= 10000; page++) {
      let url = `http://stockdata.stock.hexun.com/zrbg/data/zrbList.aspx?date=${year}-12-31&count=20&pname=20&titType=null&page=${page}&callback=hxbase_json11539258710099`;
      let result = await getData(url);

      result.list.forEach((item)=> oneYearList.push(item));
      if (!result.list.length) break;
    }
    dataList[year] = {
        list:oneYearList
    };
  }

  return dataList;
}

// 提取excel
function extratExcel() {
  let path = '2012.xlsx';
  let needCodeList = readXlsx(path);
  return needCodeList;
}

// 读取excel表格
function readXlsx(path) {
  let companyCodeList = [];
  const workbook = XLSX.readFile(path);
  const sheetNames = workbook.SheetNames;
  const sheet1 = workbook.Sheets[sheetNames[0]];

  Object.keys(sheet1).forEach((item) => {
    sheet1[item].v && companyCodeList.push(sheet1[item].v)
  });
  return companyCodeList;
}

app.set('host', process.env.IP || 'localhost');
app.set('port', process.env.PORT || 8100);

let server = app.listen(app.get('port'), app.get('host'), function () {
  console.log('Express server listening on port', app.get('host'), server.address().port);
});




















// const dataFormat = {
//   sum: 3557,
//   list: [
//     { Number: '3541', StockNameLink: 'stock_bg.aspx?code=000693&date=2017-12-31', industry: '*ST华泽(000693)', stockNumber: '-10.32', industryrate: '-6.32', Pricelimit: 'E', lootingchips: '4.00', Scramble: '0.00', rscramble: '0.00', Strongstock: '0.00', Hstock: ' <a href ="" target="_blank"><img alt="" src="img/table_btn1.gif"></img ></a>', Wstock: '<a href ="http://stockdata.stock.hexun.com/000693.shtml" target="_blank"><img alt="" src="img/icon_02.gif"></img ></a>', Tstock: '<img alt="" onclick="addIStock(\'000693\',\'1\');"  code="" codetype="" " src="img/icon_03.gif"></img >' },
//     { Number: '3542', StockNameLink: 'stock_bg.aspx?code=600716&date=2017-12-31', industry: '凤凰股份(600716)', stockNumber: '4.65', industryrate: '-6.35', Pricelimit: 'E', lootingchips: '4.00', Scramble: '0.00', rscramble: '0.00', Strongstock: '-15.00', Hstock: ' <a href ="" target="_blank"><img alt="" src="img/table_btn1.gif"></img ></a>', Wstock: '<a href ="http://stockdata.stock.hexun.com/600716.shtml" target="_blank"><img alt="" src="img/icon_02.gif"></img ></a>', Tstock: '<img alt="" onclick="addIStock(\'600716\',\'1\');"  code="" codetype="" " src="img/icon_03.gif"></img >' },
//     { Number: '3543', StockNameLink: 'stock_bg.aspx?code=002072&date=2017-12-31', industry: '凯瑞德(002072)', stockNumber: '-1.49', industryrate: '-6.44', Pricelimit: 'E', lootingchips: '1.42', Scramble: '0.00', rscramble: '0.00', Strongstock: '-6.37', Hstock: ' <a href ="" target="_blank"><img alt="" src="img/table_btn1.gif"></img ></a>', Wstock: '<a href ="http://stockdata.stock.hexun.com/002072.shtml" target="_blank"><img alt="" src="img/icon_02.gif"></img ></a>', Tstock: '<img alt="" onclick="addIStock(\'002072\',\'1\');"  code="" codetype="" " src="img/icon_03.gif"></img >' },
//     { Number: '3544', StockNameLink: 'stock_bg.aspx?code=000760&date=2017-12-31', industry: '斯太尔(000760)', stockNumber: '0.31', industryrate: '-6.69', Pricelimit: 'E', lootingchips: '3.00', Scramble: '0.00', rscramble: '0.00', Strongstock: '-10.00', Hstock: ' <a href ="" target="_blank"><img alt="" src="img/table_btn1.gif"></img ></a>', Wstock: '<a href ="http://stockdata.stock.hexun.com/000760.shtml" target="_blank"><img alt="" src="img/icon_02.gif"></img ></a>', Tstock: '<img alt="" onclick="addIStock(\'000760\',\'1\');"  code="" codetype="" " src="img/icon_03.gif"></img >' },
//     { Number: '3545', StockNameLink: 'stock_bg.aspx?code=300167&date=2017-12-31', industry: '迪威迅(300167)', stockNumber: '6.16', industryrate: '-7.12', Pricelimit: 'E', lootingchips: '1.72', Scramble: '0.00', rscramble: '0.00', Strongstock: '-15.00', Hstock: ' <a href ="" target="_blank"><img alt="" src="img/table_btn1.gif"></img ></a>', Wstock: '<a href ="http://stockdata.stock.hexun.com/300167.shtml" target="_blank"><img alt="" src="img/icon_02.gif"></img ></a>', Tstock: '<img alt="" onclick="addIStock(\'300167\',\'1\');"  code="" codetype="" " src="img/icon_03.gif"></img >' },
//     { Number: '3546', StockNameLink: 'stock_bg.aspx?code=000995&date=2017-12-31', industry: '*ST皇台(000995)', stockNumber: '-10.26', industryrate: '-7.91', Pricelimit: 'E', lootingchips: '3.00', Scramble: '0.00', rscramble: '0.00', Strongstock: '-0.65', Hstock: ' <a href ="" target="_blank"><img alt="" src="img/table_btn1.gif"></img ></a>', Wstock: '<a href ="http://stockdata.stock.hexun.com/000995.shtml" target="_blank"><img alt="" src="img/icon_02.gif"></img ></a>', Tstock: '<img alt="" onclick="addIStock(\'000995\',\'1\');"  code="" codetype="" " src="img/icon_03.gif"></img >' },
//     { Number: '3547', StockNameLink: 'stock_bg.aspx?code=600759&date=2017-12-31', industry: '洲际油气(600759)', stockNumber: '5.01', industryrate: '-8.45', Pricelimit: 'E', lootingchips: '1.54', Scramble: '0.00', rscramble: '0.00', Strongstock: '-15.00', Hstock: ' <a href ="" target="_blank"><img alt="" src="img/table_btn1.gif"></img ></a>', Wstock: '<a href ="http://stockdata.stock.hexun.com/600759.shtml" target="_blank"><img alt="" src="img/icon_02.gif"></img ></a>', Tstock: '<img alt="" onclick="addIStock(\'600759\',\'1\');"  code="" codetype="" " src="img/icon_03.gif"></img >' },
//     { Number: '3548', StockNameLink: 'stock_bg.aspx?code=600870&date=2017-12-31', industry: '*ST厦华(600870)', stockNumber: '-9.64', industryrate: '-8.74', Pricelimit: 'E', lootingchips: '0.96', Scramble: '0.00', rscramble: '0.00', Strongstock: '-0.06', Hstock: ' <a href ="" target="_blank"><img alt="" src="img/table_btn1.gif"></img ></a>', Wstock: '<a href ="http://stockdata.stock.hexun.com/600870.shtml" target="_blank"><img alt="" src="img/icon_02.gif"></img ></a>', Tstock: '<img alt="" onclick="addIStock(\'600870\',\'1\');"  code="" codetype="" " src="img/icon_03.gif"></img >' },
//     { Number: '3549', StockNameLink: 'stock_bg.aspx?code=600807&date=2017-12-31', industry: '*ST天业(600807)', stockNumber: '1.94', industryrate: '-9.06', Pricelimit: 'E', lootingchips: '4.00', Scramble: '0.00', rscramble: '0.00', Strongstock: '-15.00', Hstock: ' <a href ="" target="_blank"><img alt="" src="img/table_btn1.gif"></img ></a>', Wstock: '<a href ="http://stockdata.stock.hexun.com/600807.shtml" target="_blank"><img alt="" src="img/icon_02.gif"></img ></a>', Tstock: '<img alt="" onclick="addIStock(\'600807\',\'1\');"  code="" codetype="" " src="img/icon_03.gif"></img >' },
//     { Number: '3550', StockNameLink: 'stock_bg.aspx?code=600601&date=2017-12-31', industry: '方正科技(600601)', stockNumber: '-8.66', industryrate: '-9.24', Pricelimit: 'E', lootingchips: '0.92', Scramble: '0.00', rscramble: '0.00', Strongstock: '-1.50', Hstock: ' <a href ="" target="_blank"><img alt="" src="img/table_btn1.gif"></img ></a>', Wstock: '<a href ="http://stockdata.stock.hexun.com/600601.shtml" target="_blank"><img alt="" src="img/icon_02.gif"></img ></a>', Tstock: '<img alt="" onclick="addIStock(\'600601\',\'1\');"  code="" codetype="" " src="img/icon_03.gif"></img >' },
//     { Number: '3551', StockNameLink: 'stock_bg.aspx?code=002198&date=2017-12-31', industry: '嘉应制药(002198)', stockNumber: '-9.12', industryrate: '-9.98', Pricelimit: 'E', lootingchips: '0.74', Scramble: '0.00', rscramble: '0.00', Strongstock: '-1.60', Hstock: ' <a href ="" target="_blank"><img alt="" src="img/table_btn1.gif"></img ></a>', Wstock: '<a href ="http://stockdata.stock.hexun.com/002198.shtml" target="_blank"><img alt="" src="img/icon_02.gif"></img ></a>', Tstock: '<img alt="" onclick="addIStock(\'002198\',\'1\');"  code="" codetype="" " src="img/icon_03.gif"></img >' },
//     { Number: '3552', StockNameLink: 'stock_bg.aspx?code=002306&date=2017-12-31', industry: '*ST云网(002306)', stockNumber: '-8.58', industryrate: '-10.24', Pricelimit: 'E', lootingchips: '0.16', Scramble: '0.00', rscramble: '0.00', Strongstock: '-1.82', Hstock: ' <a href ="" target="_blank"><img alt="" src="img/table_btn1.gif"></img ></a>', Wstock: '<a href ="http://stockdata.stock.hexun.com/002306.shtml" target="_blank"><img alt="" src="img/icon_02.gif"></img ></a>', Tstock: '<img alt="" onclick="addIStock(\'002306\',\'1\');"  code="" codetype="" " src="img/icon_03.gif"></img >' },
//     { Number: '3553', StockNameLink: 'stock_bg.aspx?code=600610&date=2017-12-31', industry: '*ST毅达(600610)', stockNumber: '-11.26', industryrate: '-10.77', Pricelimit: 'E', lootingchips: '0.51', Scramble: '0.00', rscramble: '0.00', Strongstock: '-0.02', Hstock: ' <a href ="" target="_blank"><img alt="" src="img/table_btn1.gif"></img ></a>', Wstock: '<a href ="http://stockdata.stock.hexun.com/600610.shtml" target="_blank"><img alt="" src="img/icon_02.gif"></img ></a>', Tstock: '<img alt="" onclick="addIStock(\'600610\',\'1\');"  code="" codetype="" " src="img/icon_03.gif"></img >' },
//     { Number: '3554', StockNameLink: 'stock_bg.aspx?code=300106&date=2017-12-31', industry: '西部牧业(300106)', stockNumber: '-11.16', industryrate: '-11.15', Pricelimit: 'E', lootingchips: '0.65', Scramble: '0.00', rscramble: '0.00', Strongstock: '-0.64', Hstock: ' <a href ="" target="_blank"><img alt="" src="img/table_btn1.gif"></img ></a>', Wstock: '<a href ="http://stockdata.stock.hexun.com/300106.shtml" target="_blank"><img alt="" src="img/icon_02.gif"></img ></a>', Tstock: '<img alt="" onclick="addIStock(\'300106\',\'1\');"  code="" codetype="" " src="img/icon_03.gif"></img >' },
//     { Number: '3555', StockNameLink: 'stock_bg.aspx?code=000691&date=2017-12-31', industry: '亚太实业(000691)', stockNumber: '1.06', industryrate: '-11.83', Pricelimit: 'E', lootingchips: '0.32', Scramble: '0.00', rscramble: '0.00', Strongstock: '-13.21', Hstock: ' <a href ="" target="_blank"><img alt="" src="img/table_btn1.gif"></img ></a>', Wstock: '<a href ="http://stockdata.stock.hexun.com/000691.shtml" target="_blank"><img alt="" src="img/icon_02.gif"></img ></a>', Tstock: '<img alt="" onclick="addIStock(\'000691\',\'1\');"  code="" codetype="" " src="img/icon_03.gif"></img >' },
//     { Number: '3556', StockNameLink: 'stock_bg.aspx?code=300104&date=2017-12-31', industry: '乐视网(300104)', stockNumber: '-10.80', industryrate: '-13.02', Pricelimit: 'E', lootingchips: '0.07', Scramble: '0.00', rscramble: '0.00', Strongstock: '-2.29', Hstock: ' <a href ="" target="_blank"><img alt="" src="img/table_btn1.gif"></img ></a>', Wstock: '<a href ="http://stockdata.stock.hexun.com/300104.shtml" target="_blank"><img alt="" src="img/icon_02.gif"></img ></a>', Tstock: '<img alt="" onclick="addIStock(\'300104\',\'1\');"  code="" codetype="" " src="img/icon_03.gif"></img >' },
//     { Number: '3557', StockNameLink: 'stock_bg.aspx?code=600724&date=2017-12-31', industry: '宁波富达(600724)', stockNumber: '-1.59', industryrate: '-13.20', Pricelimit: 'E', lootingchips: '1.31', Scramble: '0.00', rscramble: '0.00', Strongstock: '-12.92', Hstock: ' <a href ="" target="_blank"><img alt="" src="img/table_btn1.gif"></img ></a>', Wstock: '<a href ="http://stockdata.stock.hexun.com/600724.shtml" target="_blank"><img alt="" src="img/icon_02.gif"></img ></a>', Tstock: '<img alt="" onclick="addIStock(\'600724\',\'1\');"  code="" codetype="" " src="img/icon_03.gif"></img >' }
//   ]
// }