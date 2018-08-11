#!/usr/bin/env node

const config = require('./config');
const xlsx = require('node-xlsx').default;
const fs = require('fs');
const mongoose = require('mongoose');
const mongoSavedSchema = require('./lib/schema/MongoSavedMsg');
const moment = require('moment');
var program = require('commander');

// 获取输入参数
program
    .version('0.1.0', '-v, --version')
    .option('-e, --end <n>', 'Add end time of data collection')
    .option('-s, --start <n>', 'Add start time of data collection')
    .option('-g, --gatewayId <n>', 'Add one gatewayId')
    .option('-d, --devaddr <n>', 'Add one devaddr')
    .parse(process.argv);

let mongoWhereOpt = {
    msgType: config.mongo.MONGO_SAVEDMSG_TYPE.uplink_msg,
};
let createdTime = {};
if (program.end) {
    createdTime.$gte = moment(program.end).unix();
    mongoWhereOpt.createdTime = createdTime;
}
if (program.start) {
    createdTime.$lte = moment(program.start).unix();
    mongoWhereOpt.createdTime = createdTime;
}
if (program.gatewayId) {
    mongoWhereOpt['data.gatewayId'] = program.gatewayId;
} else {
    mongoWhereOpt['data.gatewayId'] = config.mongo.gatewayId;
}
if (program.devaddr) {
    mongoWhereOpt.DevAddr = program.devaddr;
}

console.log(mongoWhereOpt);

function initMongodb(mongoConfig) {
    const url = `mongodb://${mongoConfig.host}:${mongoConfig.port}/${mongoConfig.db}`;
    mongoose.connect(url, { useNewUrlParser: true });

    const db = mongoose.connection;

    /* compile schema */
    db.on('error', function (err) {
        console.error(err);
    });
}
initMongodb(config.database.mongodb);

const fileDir = config.excel.fileDir;

// 读取文件
let originWorkSheets = [];
try {
    originWorkSheets = xlsx.parse(fs.readFileSync(fileDir));
} catch (error) {
    console.error('Read the excle File error！\n' + fileDir + error);
};

//从mongo中读取所需的数据
const collectionName = config.mongo.collectionName;
const msgModel = mongoose.model(collectionName, mongoSavedSchema, collectionName);

msgModel.find(mongoWhereOpt).then(mongoResult => {
    console.log('mongoResult', mongoResult);

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
    let updateData = [];

    if (originWorkSheets.length === 0) {
        let sheetObj = {};
        sheetObj.name = addSheetName;
        sheetObj.data = addData;
        updateData.push(sheetObj);
    } else {
        updateData = [].concat(originWorkSheets);
        let flag = true;
        for (var i = 0; i < originWorkSheets.length; i++) {
            if (originWorkSheets[i]) {
                let originSheetName = originWorkSheets[i].name;
                if (addSheetName === originSheetName) {
                    flag = false;
                    let originSheetData = originWorkSheets[i].data;
                    updateData[i].data = originSheetData.concat([[]], addData);
                }
            }
        }
        if (flag) {
            let sheetObj = {};
            sheetObj.name = addSheetName;
            sheetObj.data = addData;
            updateData.push(sheetObj);
        }
    }
    const buffer = xlsx.build(updateData);

    try {
        fs.writeFileSync(fileDir, buffer, {});
        console.log('Success!');
    } catch (error) {
        if (error.code == 'EBUSY') {
            console.log('Please close the excle file');
        } else {
            console.error(error);
        }
    }

    mongoose.disconnect();
}).catch(err => {
    console.error(err);
});
