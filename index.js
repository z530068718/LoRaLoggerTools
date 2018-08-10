const config = require('./config');
const xlsx = require('node-xlsx').default;
const fs = require('fs');
const mongoose = require('mongoose');
const mongoSavedSchema = require('./lib/schema/MongoSavedMsg');
const moment = require('moment');

// 获取参数
const arguments = process.argv.splice(2);
const startTime = arguments[0];
const endTime = arguments[1];

function initMongodb(mongoConfig) {
    const url = `mongodb://${mongoConfig.host}:${mongoConfig.port}/${mongoConfig.db}`;
    mongoose.connect(url, { useNewUrlParser: true });

    const db = mongoose.connection;

    /* compile schema */
    db.on('error', function (err) {
        console.error(err);
    });
}
initMongodb(config.mongoose);

const fileDir = config.excel.fileDir;

// 读取文件
let originWorkSheets = null;
try {
    originWorkSheets = xlsx.parse(fs.readFileSync(fileDir));
} catch (error) {
    console.error("Read the excle File error！\n" + fileDir + error);
    return;
};

//从mongo中读取所需的数据
const collectionName = config.mongo.collectonName;
const msgModel = mongoose.model(collectionName, mongoSavedSchema);
const whereOpts = {
    gatewayId: config.mongo.gatewayId,
    msgType: config.mongo.MONGO_SAVEDMSG_TYPE.uplink_msg,
    createdTime: { $gte: moment(startTime).unix(), $lte: moment(endTime).unix() },
};

msgModel.find(whereOpts).then((err, mongoResult) => {
    if (err) {
        console.error(err);
        return;
    }

    let addSheetName = config.excel.sheetName;
    let addData = [];

    function converter(mongoDataArr) {
        let outputArr = [];
        let excelObj = ['DevAddr', 'GatewayId', 'msgType', 'createdTime', 'datr', 'freq', 'rssi', 'lsnr', 'Fcnt'];
        excelObj = [];
        for (let i = 0; i < mongoDataArr.length; i++) {
            excelObj.push(mongoDataArr[i].DevAddr);
            excelObj.push(mongoDataArr[i].GatewayId);
            excelObj.push(mongoDataArr[i].msgType);
            excelObj.push(mongoDataArr[i].createdTime);
            excelObj.push(mongoDataArr[i].data.rxpk.datr);
            excelObj.push(mongoDataArr[i].data.rxpk.freq);
            excelObj.push(mongoDataArr[i].data.rxpk.rssi);
            excelObj.push(mongoDataArr[i].data.rxpk.lsnr);
            excelObj.push(mongoDataArr[i].data.rxpk.data.MACPayload.FHDR.FCnt);

            //去除完全相同的数据包统计
            if (i >= 1 && mongoDataArr[i].createdTime === mongoDataArr[i - 1].createdTime) {
                continue;
            } else {
                outputArr.push(excelObj);
                excelObj = [];
            }
        }
        return outputArr;
    }

    addData = converter(mongoResult);
    // 获取指定sheet的数据，并向目标表添加数据

    let updateDatas = [].concat(originWorkSheets);
    for (var i = 0; i < originWorkSheets.length; i++) {
        if (originWorkSheets[i]) {
            let originSheetName = originWorkSheets[i].name;
            if (addSheetName === originSheetName) {
                let originSheetData = originWorkSheets[i].data;
                updateDatas[i].data = originSheetData.concat([[]], addData);
            }
        }
    }

    const buffer = xlsx.build(updateDatas);

    try {
        fs.writeFileSync('test.xlsx', buffer, {});
        console.log('Success!');
    } catch (error) {
        if (error.code == 'EBUSY') {
            console.log('Please close the excle file');
        } else {
            console.error(error);
        }
    }
});
