$(function() {
  let enumName = [
    '持续改进',
    '发展他人',
    '建立关系',
    '尽职尽责',
    '团队合作',
    '以客户为中心',
    '主动学习'
  ];

  $('#excelFile').change(function(parentEvent) {
    let files = parentEvent.target.files;

    let fileReader = new FileReader();

    let getExcelList = [];
    fileReader.onload = function(childEvent) {
      let excelBinaryData;
      // 读取上传的excel文件
      try {
        let excelData = childEvent.target.result;
        excelBinaryData = XLSX.read(excelData, {
          type: 'binary'
        });
      } catch (parentEvent) {
        console.log('该文件类型不能识别');
        return;
      }

      // 获取excel所有元素
      for (let sheet in excelBinaryData.Sheets) {
        if (excelBinaryData.Sheets.hasOwnProperty(sheet)) {
          let excelSheet = XLSX.utils.sheet_to_json(excelBinaryData.Sheets[sheet]);
          getExcelList[getExcelList.length] = excelSheet;
        }
      }

      let newExcelList = [];

      // 找出可以合并的数据
      for (let i = 0; i < getExcelList[0].length; i = i + 7) {
        let columnTemp = getBaseColumn(getExcelList[0][i]);
        for (let j = i + 1; j < i + 8 && j < getExcelList[0].length; j++) {
          if (getExcelList[0][i]['对象 员工工号'] == getExcelList[0][j]['对象 员工工号']) {
            columnTemp = getNewColumn(columnTemp, getExcelList[0][j]);
          }
        }
        newExcelList[newExcelList.length] = columnTemp;
      }

      download(newExcelList, files);
    };

    // 以二进制方式打开文件
    fileReader.readAsBinaryString(files[0]);
  });

  $('#excelFile2').change(function(parentEvent) {
    let files = parentEvent.target.files;

    let fileReader = new FileReader();

    let getExcelList = [];
    fileReader.onload = function(childEvent) {
      let excelBinaryData;
      // 读取上传的excel文件
      try {
        let excelData = childEvent.target.result;
        excelBinaryData = XLSX.read(excelData, {
          type: 'binary'
        });
      } catch (parentEvent) {
        console.log('该文件类型不能识别');
        return;
      }

      // 获取excel所有元素
      for (let sheet in excelBinaryData.Sheets) {
        if (excelBinaryData.Sheets.hasOwnProperty(sheet)) {
          let excelSheet = XLSX.utils.sheet_to_json(excelBinaryData.Sheets[sheet]);
          getExcelList[getExcelList.length] = excelSheet;
        }
      }

      let newExcelList = [];

      // 找出可以分解的数据
      for (let i = 0; i < getExcelList[0].length; i++) {
        for (let j = 0; j < enumName.length; j++) {
          newExcelList[newExcelList.length] = getChildColumn(getExcelList[0][i], enumName[j]);
        }
      }

      download(newExcelList, files);
    };

    // 以二进制方式打开文件
    fileReader.readAsBinaryString(files[0]);
  });

  function getChildColumn(column, chooseColumnName) {
    let columnTemp = {};
    columnTemp['对象 员工工号'] = column['对象 员工工号'];
    columnTemp['对象 姓名'] = column['对象 姓名'];
    columnTemp['资格姓名'] = chooseColumnName;
    columnTemp['资格正式评分'] = column[chooseColumnName];

    return columnTemp;
  }

  function getBaseColumn(columnOne) {
    let columnTemp = {};
    columnTemp['对象 员工工号'] = columnOne['对象 员工工号'];
    columnTemp['对象 姓名'] = columnOne['对象 姓名'];
    columnTemp = getNewColumn(columnTemp, columnOne);
    return columnTemp;
  }

  function getNewColumn(columnTemp, columnTwo) {
    let key = columnTwo['资格姓名'];
    let value = columnTwo['资格正式评分'];

    columnTemp[key] = value;
    return columnTemp;
  }

  function download(newExcelList, files) {
    const newSheet = {
      SheetNames: ['Sheet0'],
      Sheets: {},
      Props: {}
    };
    const sheetDownloadType = { bookType: 'xlsx', bookSST: false, type: 'binary' };

    newSheet.Sheets['Sheet0'] = XLSX.utils.json_to_sheet(newExcelList);
    saveAs(
      new Blob(
        [
          stringToArrayBuffer(XLSX.write(newSheet, sheetDownloadType))
        ], {
          type: "application/octet-stream"
        }
      ),
      files[0].name
    );
  }

  function stringToArrayBuffer(data) {
    if (typeof ArrayBuffer !== 'undefined') {
      let arrayBuffer = new ArrayBuffer(data.length);
      let unitArray = new Uint8Array(arrayBuffer);
      for (let unitI = 0; unitI != data.length; unitI++) {
        unitArray[unitI] = data.charCodeAt(unitI) & 0xFF;
      }
      return arrayBuffer;
    } else {
      let arrayBuffer = new Array(data.length);
      for (let bufferI = 0; bufferI != data.length; bufferI++) {
        arrayBuffer[bufferI] = data.charCodeAt(bufferI) & 0xFF;
      }
      return arrayBuffer;
    }
  }

  function saveAs(content, fileName) {
    let clickDiv = document.createElement("a");
    clickDiv.download = fileName || "下载";
    clickDiv.href = URL.createObjectURL(content);
    clickDiv.click();
    setTimeout(function () {
      URL.revokeObjectURL(content);
    }, 100);
  }
})
