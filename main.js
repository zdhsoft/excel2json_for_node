const path = require('path');
const fs = require('fs');
const glob = require('glob');
const {utils, common_ret} = require('xmcommon');
const xlsx = require('node-xlsx');


/** 表信息 */
class XSheetInfo {
    /** 输出的文件名 */
    fileName  = '';
    /** 输出的类名 */
    className = '';
    /** 配置表的表格名称 */
    sheetName = '';
    /** 参数配置数据 */
    data = []
}
/** 表头数据类型 */
class XHeadType {
    /** 是否是数组 */
    isArray = false;
    /** 类型定义 */
    type = '';
}

/** 表头信息 */
class XSheetHead {
    /** 通用名称 */
    normalName = '';
    /** 是否导出 */
    export     = false;
    /** 类型 */
    type      = new XHeadType();
    /** 输出名称 */
    name      = '';
    /** 输出范围 表示是客户端，还是服务端输出 */
    scope     = [];
}
/** 表头结果 */
class XSheetHeadResult {
    /** 表头结果 */
    result = false;
    /** 错误信息 */
    errMsg = '';


}

/**
 * 设置错误信息
 * @param {{result: boolean, errMsg: string}} ret 返回的信息
 * @param {{string}} errMsg 错误信息
 * @return {{result: boolean, errMsg: string}}
 */
function setRetError (ret, errMsg){
    ret.result = false;
    ret.errMsg = errMsg;
    return ret;
}

/**
 * 类型定义常量
 * @type {{Array: string, Integer: string, String: string, Boolean: string, Any: string}}
 */
const TypeDef = {
    /** 数组 */
    Array  : 'array',
    /** 字符串 */
    String : 'string',
    /** 整数 */
    Integer: 'int',
    /** 布尔 */
    Boolean: 'bool',
    /** 任意类型 */
    Any    : 'any'
};
/**
 * 类型列表
 * @type {string[]}
 */
const TypeList = Object.values(TypeDef);

/**
 * 打印日志
 * @param {...} argv 要打印的参数列表
 */
function log(...argv) {
    console.log(...argv);
}

let f = Number.MAX_SAFE_INTEGER;

/**
 * 检查名称的正则表达式
 * @type {RegExp}
 */
const checkNameReg = /^[a-zA-Z_\$][a-zA-Z\d_\$]{0,200}$/;
const checkIntegerReg = /^[+-]{0,1}\d{1,17}$/;

const OutFlag = {
    Server: 's',
    Client: 'c'
}
const OutFlagList = Object.values(OutFlag);

// 输出的跟目录
const destRoot = '.';
// 服务器端输出目录
const svrRoot  = 'svr';
// 前端输出目录
const clientRoot = 'client';

/**
 * 确定指定文件的目录存在
 * @param {string} paramFullFileName
 */
function ensureDirectory(paramFullFileName) {
    path.parse()
}

/**
 * 生成表头
 * @param {string[]} paramNormalNameList 普通名称列表
 * @param {string[]} paramTypeList 类型列表
 * @param {string[]} paramNameList 名称列表
 * @param {string[]} paramScopeList 适用范围列表
 * @return {{result: boolean, errMsg: string, head:[{normalName:string, export: boolean, type: {isArray: boolean, type: string}, name: string, scope: string[]}]}} 处理结果
 */
function makeTableHead(paramNormalNameList, paramTypeList, paramNameList, paramScopeList) {
    let ret = {
        result: false,
        errMsg: '',
        head: []
    };
    do {

        if (!Array.isArray(paramNormalNameList)) {
            setRetError(ret, `传入的普通名称列表不是数组: ${paramNormalNameList}`);
            break;
        }
        if (!Array.isArray(paramTypeList)) {
            setRetError(ret, `传入的类型列表不是数组: ${paramTypeList}`);
            break;
        }
        if (!Array.isArray(paramNameList)) {
            setRetError(ret, `传入的名称列表不是数组: ${paramNameList}`);
            break;
        }
        if (!Array.isArray(paramScopeList)) {
            setRetError(ret, `传入的适用范围列表不是数组: ${paramScopeList}`);
            break;
        }

        let len = paramNormalNameList.length;
        if (paramTypeList.length !== len || paramNameList.length !== len || paramScopeList.length !== len) {
            setRetError(ret, `传入的参数数组长度不一样：普通名称列表:${len}, 类型列表:${paramNameList.length}, 名称列表:${paramTypeList.length}, 适用范围列表:${paramScopeList.length}`);
            break;
        }

        if (len === 0) {
            setRetError(ret, `传入的参数，列表长度为0`);
            break;
        }
        /**
         * 导出标志列表
         * @type {[{normalName:string, export: boolean, type: {isArray: boolean, type: string}, name: string, scope: string[]}]}}
         */
        let head = [];
        // 初始化头的值
        for(let i = 0; i < len; i++) {
            head.push({
                normalName: paramNormalNameList[i],
                export: true,
                type: {isArray: false, type: TypeDef.Any},
                name: '',
                scope:[],
                index: i,
            })
        }
        // 检查名称
        let checkFail = false;
        let nameSet = new Set();
        for(let i = 0; i < len; i++) {
            const name = paramNameList[i];
            log(`name[${i}]=${name}`);
            if (utils.isNull(name)) {
                // 如果为空的，将不会被导出
                head[i].export = false;
                continue;
            }
            if (!checkNameReg.test(name)) {
                setRetError(ret, `名称列表[${i}]=${name} : 不是有效的名称`);
                checkFail = true;
                break;
            }
            if (nameSet.has(name)) {
                setRetError(ret, `名称列表[${i}]=${name} : 存在重复的名称`);
                checkFail = true;
                break;
            }
            nameSet.add(name);
            head[i].name = name;
        }
        if (checkFail) {
            break;
        }

        // 检查范围
        for(let i = 0; i < len; i++) {
            if (!head[i].export) {
                continue;
            }
            let scope = paramScopeList[i];
            if (utils.isNull(scope)) {
                head[i].export = false;
                continue;
            }

            let s = scope.trim().split('');
            if (s.length === 0) {
                head[i].export = false;
                continue;
            }
            if (s.length > 2) {
                setRetError(ret, `适用范围[${i}]=${scope}，长度超过2！`);
                checkFail = true;
                break;
            }
            if (s.length === 2 && s[0] === s[1]) {
                setRetError(ret, `适用范围[${i}]=${scope}, 两个标志相同！`);
                checkFail = true;
                break;
            }

            if (!OutFlagList.includes(s[0])) {
                setRetError(ret, `适用范围[${i}]=${scope}中的${s[0]}不是指定的标志，只能是${OutFlagList.join()}`);
                checkFail = true;
                break;
            }

            if (s.length === 2) {
                if (!OutFlagList.includes(s[1])) {
                    setRetError(ret, `适用范围[${i}]=${scope}中的${s[1]}不是指定的标志，只能是${OutFlagList.join()}`);
                    checkFail = true;
                    break;
                }
            }
            head[i].scope = s;
        }
        if (checkFail) {
            break;
        }
        // 检查类型
        // 支待的类型  string, boolean, int, number, any, array:string, array:boolean, array:int, array: number, array:any
        for (let i = 0; i < len; i++) {
            let t = {
                isArray: false,
                type   : TypeDef.Any
            }
            if(!head[i].export) {
                // 如果不用导出的，则不用检查类型
                continue;
            }
            let oriType = paramTypeList[i];
            let type = oriType.split(':');

            if (!Array.isArray(type) || (!(type.length === 1 || type.length === 2))) {
                setRetError(ret, `类型[${i}]=${oriType}是无效的！`);
                checkFail = true;
                break;
            }
            let typeLen    = type.length;
            // 主类型
            let mainType   = type[0].trim();
            // 次类型
            let secondType = typeLen > 1? type[1].trim() : null;
            // 仅有主类型是数组的时候，才会有次类型
            if (mainType !== TypeDef.Array && typeLen > 1) {
                setRetError(ret, `类型[${i}]=${oriType}不是数组类型定义，却有:分隔！`);
                checkFail = true;
                break;
            }
            // 如果主类型是数组
            if (mainType === TypeDef.Array) {
                if (secondType === TypeDef.Array) {
                    setRetError(ret, `类型[${i}]=${oriType}的主类型与第二类型，都是${TypeDef.Array}`);
                    break;
                }
                t.isArray = true;
                t.type    = secondType;
            } else {
                t.isArray = false;
                t.type = mainType;
            }

            if (!TypeList.includes(t.type)) {
                setRetError(ret, `类型[${i}]=${oriType}的类型是未定义的类型，只支持以下类型:${TypeList.join()}`);
                break;
            }
            head[i].type = t;
        }
        if (checkFail) {
            break;
        }
        ret.result = true;
        ret.head = head;
    } while(false);
    return ret;
}

/**
 * 生成数组
 * @param {boolean} paramIsArray 是否是数组
 * @param {string} paramType 数据类型
 * @param {string} paramData 配置原始数据
 * @return {{result: boolean, errMsg: string, value: any}}
 */
function GetValue(paramIsArray, paramType, paramData) {

}

/**
 * 根据类型生成返回数组
 * @param {string} paramType 数据类型
 * @param {string} paramData 配置原始数据
 * @return {{result: boolean, errMsg: string, value: any}}
 */
function GetValueByType(paramType, paramData) {
    let ret = {
        result: false,
        errMsg: '',
        data: null
    };
    switch (paramType) {
        case TypeDef.Boolean: {
                let t = 'false';
                if (utils.isNotNull(paramData)) {
                    t = paramData.trim().toLowerCase();
                }
                ret.data = t === 'true' ? true : false;
            }
            break;
        case TypeDef.Integer: {
            // checkIntegerReg.test(paramData)
        }

    }
}



/**
 * 判断是不是配置的ExcelShell
 * @param {{name: string, data:[]}} paramExcelSheet
 * @return {{result:boolean, errMsg: string, head:[], info:{fileName: string, className: string, sheetName: string, data: []}}} 判断结果
 */
function isConfig(paramExcelSheet) {
    let ret = {
        result: false,
        errMsg: '',
        head:[],
        info: {
            fileName: '',
            className: '',
            sheetName: '',
            data:[]
        },
    }

    do {
        const data = paramExcelSheet.data;
        if (data.length < 6) {
            setRetError(ret, '不是有效的配置格式，行数小于6');
            break;
        }

        /** @type {string[]} 文件名信息*/
        const fileName = data[0];
        /** @type {string[]} 类名信息 */
        const className = data[1];

        if (!Array.isArray(fileName)) {
            setRetError(ret, '第1行不是有效的数组:' + JSON.stringify(fileName));
            break;
        }

        if (fileName.length < 2) {
            setRetError(ret, `第1行有效的元素个数=${fileName.length} < 2`);
            break;
        }

        if (fileName[0].replace(/\s*/g,'') !== '文件名：') {
            setRetError(ret, `第1行有效的第一个元素不是'文件名：'`);
            break;
        }

        if (fileName[1].trim() === '') {
            setRetError(ret, `第1行有效的第二个元素不是有效文件名！`);
            break;
        }

        if (!Array.isArray(className)) {
            setRetError(ret, '第2行不是有效的数组:' + JSON.stringify(className));
            break;
        }

        if (className.length < 2) {
            setRetError(ret, `第2行有效的元素个数=${className.length} < 2`);
            break;
        }

        if (className[0].replace(/\s*/g,"") !== '类名：') {
            setRetError(ret, `第2行有效的第一个元素不是'类名：'`);
            break;
        }

        const tempClassName = className[1].trim();
        if (!checkNameReg.test(tempClassName)) {
            setRetError(ret, `第2行有效的第二个元素不是有效类名！${tempClassName}`);
            break;
        };

        ret.info.fileName  = fileName[1].trim();
        ret.info.className = tempClassName;

        let checkResult = makeTableHead(paramExcelSheet.data[2], paramExcelSheet.data[3], paramExcelSheet.data[4], paramExcelSheet.data[5]);
        if (!checkResult.result) {
            ret.result = false;
            ret.errMsg = checkResult.errMsg;
            break;
        }
        ret.result = true;
        ret.info.data = paramExcelSheet.data.splice(6);
        ret.head = checkResult.head;

    } while (false);
    return ret;
}

/**
 * 打印运行参数
 */
function printargs() {
    log('excel2json for nodejs 1.0');
    log('args: excelpath destpath');
}

/**
 * 转换成绝对路径path
 * @param  {string} paramPath 当前路径
 * @return {string} 绝对路径
 */
function toAbsolutePath(paramPath) {
    let p = paramPath;
    do {
        if (path.isAbsolute(p)) {
            break;
        }
        p = path.resolve(p);
    } while(false);
    return p;
}

/**
 * 入口函数
 * @param {string[]} argv 启动参数
 */
async function main(argv) {
    let ret = new common_ret();
    do {
        if (!utils.isArray(argv) || argv.length !== 2) {
            printargs();
            break;
        }
        const excelPath = toAbsolutePath(argv[0]);
        const destPath = toAbsolutePath(argv[1]);

        log('excelPath:' + excelPath);
        log('destPath:' + destPath);

        const sheetList = [];

        ret = await xlsxList(excelPath);
        if (ret.isNotOK) {
            break;
        }
        /** @type {string[]} 文件列表 */
        const files = ret.data;
        files.forEach(excelFile=>{
            const fullPath = path.join(excelPath, excelFile);
            const r = xlsx.parse(fullPath);
            log(`开始解析文件: ${fullPath}`);
            r.forEach((sheet)=>{
                log(`    表：${sheet.name}`);
                let checkResult = isConfig(sheet);
                if (checkResult.result) {
                    log(JSON.stringify(checkResult.info, null, 2));
                    log(JSON.stringify(checkResult.head, null, 2));
                    const sheetInfo = {info:checkResult.info, head: checkResult.head};
                    sheetList.push();

                } else {
                    log(JSON.stringify(checkResult));
                }

            });

        });

    } while (false);
    if (ret.isNotOK) {
        log('出错了:' + ret.getErrorInfo());
    }
    process.exit(0);
}

/**
 * 取excel文件列表
 * @param {string} paramExcelPath excel所在的文件目录
 * @return {common_ret}
 */
async function xlsxList(paramExcelPath) {
    let ret = new common_ret();
    do {
        if (!fs.existsSync(paramExcelPath)) {
            ret.setError(-1, '目录不存在！'+paramExcelPath);
            break;
        }
        const [err, files] = await utils.WaitFunctionEx(glob, '+(*.xlsx|*.xls)', {cwd: paramExcelPath});
        if (utils.isNotNull(err)) {
            ret.setError(-2, JSON.stringify(err));
            break;
        }
        ret.setOK(files);
    } while (false);
    return ret;
}


main(process.argv.splice(2));



