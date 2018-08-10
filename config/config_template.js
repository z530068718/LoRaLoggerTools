'use strict'

module.exports = {
    database: {
        mongodb: {
            host: 'localhost',
            port: 27017,
            db: 'mqttLogger',
            cluster: false,
        }
    },
    excel: {
        sheetName: 'test',
        fileDir: `${__dirname}/` + '../test.xlsx',
        saveDir: `${__dirname}/` + '../test.xlsx',
    },
    mongo: {
        collectonName: '',
        MONGO_SAVEDMSG_TYPE = {
            uplink_joinReq: 'UPLINK_JOINREQ',
            uplink_msg: 'UPLINK_MSG',
            uplink_gatewayStat: 'GATEWAYSTAT',
            downlink_joinAns: 'DONWLINK_JOINANS',
            downlink_msg: 'DOWNLINK_MSG',
        },
        DevAddr :'',
        gatewayId :'32f9bdfffee3e603',
    }
}