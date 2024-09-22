/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
let binaryStrings;
let logContent;
let groupedData;


document.getElementById('fileInput').addEventListener('change', function(event) {
  var file = event.target.files[0]; // 获取用户选择的文件
  if (!file) {
      return;
  }

  var reader = new FileReader(); // 创建FileReader对象
  reader.readAsText(file); // 以文本形式读取文件
  reader.onload = function(e) {
  //var content = e.target.result; // 读取文件内容
  logContent= e.target.result; // 读取文件内容
      // 运行提取并分组函数
  groupedData = extractAndGroupById(logContent);

    // 输出结果
    for (const id in groupedData) {
        console.log(`CAN ID: ${id}`);
        groupedData[id].forEach(entry => {
            console.log(`  Time: ${entry.time}, Data: ${entry.data}`);
        });
    }



      //console.log(content); // 打印到控制台
      //let hexData = extractAllHexData(content);
      //console.log(hexData);
      // 转换为反序的二进制字符串数组
      //binaryStrings = hexData.map(hexString => hexStringToReversedBinary(hexString));
      // 输出32位的反序二进制字符串数组
      //console.log(binaryStrings);
      // 这里可以对content进行进一步处理，例如显示在页面上或发送到服务器
  };

});

function hexStringToReversedBinary(hexString) {
  // 移除可能的空格和双引号
  hexString = hexString.replace(/"/g, '').replace(/\s+/g, '');
  
  // 分割字符串为数组，并将每个十六进制数转换为二进制字符串
  const hexArray = hexString.match(/.{1,2}/g);
  const binaryArray = hexArray.map(hex => parseInt(hex, 16).toString(2).padStart(8, '0'));
  
  // 反序拼接二进制字符串
  let binaryString = binaryArray.reverse().join('');
  
  // 确保二进制字符串是32位长
  binaryString = binaryString.padStart(32, '0');
  
  return binaryString;
}

function extractAllHexData(content) {
  // 定义一个正则表达式来匹配每一行中的所有16进制数据
  const hexPattern = /Tx\s+d\s+8\s+([0-9A-F ]+)/g;
  // 定义一个数组来存储提取的16进制数据
  let hexDataArray = [];

  // 使用正则表达式匹配并提取数据
  let match;
  while ((match = hexPattern.exec(content)) !== null) {
    // 移除空格，只保留16进制字符
    let hexData = match[1].replace(/\s/g, '');
    // 将匹配到的16进制数据添加到数组中
    hexDataArray.push(hexData);
  }

  // 返回提取的16进制数据数组
  return hexDataArray;
}


Office.onReady(() => {
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;
  document.getElementById("tang").onclick = genTang;
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      // 获取当前工作簿的所有工作表
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
  
      await context.sync();
  
      // 检查工作表数量并打印所有工作表的名称
      if (sheets.items.length > 1) {
          console.log(`There are ${sheets.items.length} worksheets in the workbook:`);
      } else {
          console.log(`There is one worksheet in the workbook:`);
      }
  
      sheets.items.forEach(function (sheet) {
          console.log(sheet.name);
      });
  
      // 检查是否存在名为"BinaryLog"的工作表
      let binaryLogSheet = sheets.getItemOrNullObject("BinaryLog");
      await context.sync();
  
      if (!binaryLogSheet.isNullObject) {
          // 如果存在，则删除它
          binaryLogSheet.delete();
      }
  
      // 添加一个新的名为"BinaryLog"的工作表
      let sheetBinaryLog= sheets.add("BinaryLog");
      const numberOfRows = binaryStrings.length;
      console.log(`Row indices with ID=64: ${numberOfRows}`);
      const startCellLog = "A2"; // Starting from cell A1
      const endCellLog = `A${numberOfRows+1}`; // Calculate the ending cell based on the array length
      const rangeAddressLog = `${startCellLog}:${endCellLog}`; // Define the range address
      const rangeLog = sheetBinaryLog.getRange(rangeAddressLog);
      let stringLogArray= binaryStrings.map(num => [num]);
      // Set the values of the range with the stringArray
      rangeLog.numberFormat = '@'; // '@' sets the format to Text
      rangeLog.values = stringLogArray;
      await context.sync();
  });
  } catch (error) {
    console.error(error);
  }
}

export async function genTang() {
  try {
    await Excel.run(async (context) => {
      // 获取当前工作簿的所有工作表
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
  
      await context.sync();
  
      // 检查工作表数量并打印所有工作表的名称
      if (sheets.items.length > 1) {
          console.log(`There are ${sheets.items.length} worksheets in the workbook:`);
      } else {
          console.log(`There is one worksheet in the workbook:`);
      }
  
      sheets.items.forEach(function (sheet) {
          console.log(sheet.name);
      });
  
      // 检查是否存在名为"BinaryLog"的工作表
      let binaryLogSheet = sheets.getItemOrNullObject("TANG");
      await context.sync();
  
      if (!binaryLogSheet.isNullObject) {
          // 如果存在，则删除它
          binaryLogSheet.delete();
      }
  
      // 添加一个新的名为"BinaryLog"的工作表
      let sheetTang= sheets.add("TANG");
 // 计算XOR
      let resultXORArray = xorAdjacentElementsDirect(binaryStrings);
      console.log(`Result XOR Array:/${resultXORArray}`);
      let valTang=sumBinaryColumns(resultXORArray);
      console.log(valTang);
      const rangeBitheader = sheetTang.getRange("A1:BL1");
      const tangValuerange = sheetTang.getRange("A2:BL2");
      // 填充单元格
      let bitLabels = [];
      for (let i = 63; i >= 0; i--) {
        bitLabels.push(`bit${i}`);
      }
      
      // 设置单元格值
      rangeBitheader.values = [bitLabels];
      tangValuerange.values=[splitStringIntoSubarrays(valTang).map(Number)];
      console.log(splitStringIntoSubarrays(valTang));


      let dataTangRange = sheetTang.getRange("A1:BL2");
      let chartTang = sheetTang.charts.add(
      Excel.ChartType.line, 
      dataTangRange, 
      Excel.ChartSeriesBy.auto);

      chartTang.title.text = "TANG";
      chartTang.legend.position = Excel.ChartLegendPosition.right;
      chartTang.legend.format.fill.setSolidColor("white");
      chartTang.dataLabels.format.font.size = 15;
      chartTang.dataLabels.format.font.color = "black";



      await context.sync();
  });
  } catch (error) {
    console.error(error);
  }
}

function xorAdjacentElementsDirect(binaryStringArray) {
  // 获取字符串的长度
  const length = binaryStringArray[0].length;
  // 创建一个新数组，长度为原数组长度减1
  let xorArray = [];

  // 遍历原数组，除了最后一个元素
  for (let i = 0; i < binaryStringArray.length - 1; i++) {
    // 遍历每个字符串的每个字符
    let xorString=xorBinaryStrings(binaryStringArray[i],binaryStringArray[i+1]);
    // 将完整的XOR结果字符串添加到新数组中
    xorArray.push(xorString);
  }

  return xorArray;
}

function xorBinaryStrings(str1, str2) {
  // 确保两个字符串长度相同，较短的字符串前面补0
  let maxLength = Math.max(str1.length, str2.length);
  let xorResult = '';
  for (let i = 0; i < maxLength; i++) {
    // 对应位进行XOR操作
    xorResult += (parseInt(str1[i], 10) ^ parseInt(str2[i], 10)).toString();
  }
  return xorResult;
}

function sumBinaryColumns(binaryArray) {
  // 确定数组中最长的字符串长度
  let maxLength = Math.max(...binaryArray.map(str => str.length));
  
  // 初始化每列的和为0，长度为最长字符串的长度
  let sum = new Array(maxLength).fill(0);

  // 遍历数组的每一行
  for (let row of binaryArray) {
    // 遍历每一列
    for (let col = 0; col < maxLength; col++) {
      // 如果当前位置有值（即不是超出当前行字符串长度的填充0），则累加到对应的列
      if (col < row.length) {
        sum[col] += parseInt(row[col], 2);
      }
    }
  }

  // 将每列的和转换为十进制字符串并返回
  return sum.join(',');
}

function splitStringIntoSubarrays(str) {
  // 使用split方法按逗号分割字符串
  let numbersArray = str.split(',');
  // 使用map方法将每个分割后的字符串放入一个单独的数组中
  let subarrays = numbersArray.map(number => [number]);
  return subarrays;
}

// 函数：过滤并提取时间序列数据，并按照ID分组
function extractAndGroupById(log) {
  const lines = log.trim().split('\n');  // 按行分割日志内容
  const groupedById = {};

  lines.forEach(line => {
      // 使用正则表达式匹配有效的CAN Bus日志数据行
      const match = line.match(/^\s*([\d.]+)\s+1\s+([\dA-F]+)\s+Rx\s+d\s+\d\s+([\dA-F\s]+)/i);
      if (match) {
          const time = parseFloat(match[1]);  // 提取时间戳
          const id = match[2];                // 提取CAN ID
          const data = match[3].trim();       // 提取数据字段

          // 如果该ID还没有记录，初始化一个空数组
          if (!groupedById[id]) {
              groupedById[id] = [];
          }
          // 将该条数据添加到对应ID的数组中
          groupedById[id].push({ time, data });
      }
  });

  return groupedById;
}