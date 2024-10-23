/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
let logContent;
let groupedData;



Office.onReady(() => {
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("tang").onclick = genTang;
});

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
      const ids = Object.keys(groupedData);
      console.log("Extracted IDs:", ids);
// Assuming groupedData is already populated
      Object.keys(groupedData).forEach((id,index) => {
        const dataArray = groupedData[id].map(entry => entry.data);
        let resultXORArray = xorAdjacentElementsDirect(dataArray);
        let valTang=sumBinaryColumns(resultXORArray);
        console.log(`id: ${id} Tang is ${valTang}`);
         // 获取 resultXORArray 的长度
        const arrayLength = groupedData[id][0].data.length;
        console.log(`Array Length is ${arrayLength}`);
        let startRow = 1 + index*3; // 起始行号
        let endRow = 1 + index*3; // 结束行号
    
        let startColIndex = 0; 
        let endColIndex = arrayLength-1; 
    
        // 使用getExcelColumnLabel函数将列索引转换为列字母
        let startColLetter = getExcelColumnLabel(startColIndex);
        let endColLetter = getExcelColumnLabel(endColIndex);
    
        // 拼接rangeAddress
        let idAdddress = `${startColLetter}${startRow}`;
        let rangeAddress = `${startColLetter}${startRow+1}:${endColLetter}${endRow+1}`;
        let valueAddress = `${startColLetter}${startRow+2}:${endColLetter}${endRow+2}`;
        let chartDataAddress = `${startColLetter}${startRow+1}:${endColLetter}${endRow+2}`;
        console.log(rangeAddress);  // 输出每次循环生成的rangeAddress
        console.log(valueAddress);  // 输出每次循环生成的valueAddress
        let bitLabels = [];
        for (let i = endColIndex; i >= 0; i--) {
          bitLabels.push(`bit${i}`);
        }
        const idheader = sheetTang.getRange(idAdddress);
        const rangeBitheader = sheetTang.getRange(rangeAddress);
        const tangValuerange = sheetTang.getRange(valueAddress);
        const chartRange = sheetTang.getRange(chartDataAddress);
        idheader.values=[[`id = ${id}`]];
        rangeBitheader.values = [bitLabels];
        tangValuerange.values=[splitStringIntoSubarrays(valTang).map(Number)];
        let chart = sheetTang.charts.add(
          Excel.ChartType.line, 
          chartRange, 
          Excel.ChartSeriesBy.auto);
    
        chart.title.text = `id = ${id}`;
        chart.legend.position = Excel.ChartLegendPosition.right;
        chart.legend.format.fill.setSolidColor("white");
        chart.dataLabels.format.font.size = 15;
        chart.dataLabels.format.font.color = "black";        
      });

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
          const data = hexToBinary(match[3].trim().replace(/\s+/g, ''));       // 提取数据字段

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

function hexToBinary(hexStr) {
  let binaryStr = '';
  
  // 每两个字符表示一个字节，因此按字节对字符串进行分割并倒序
  for (let i = hexStr.length - 2; i >= 0; i -= 2) {
      // 提取每个字节（两个16进制字符）
      let byte = hexStr.slice(i, i + 2);
      
      // 将字节转换为二进制并且填充为8位
      let bin = parseInt(byte, 16).toString(2).padStart(8, '0');
      
      // 拼接到最终的二进制字符串中
      binaryStr += bin;
  }
  
  return binaryStr;
}

// 将列号转换为Excel列字母，如0 -> A, 1 -> B, ..., 26 -> AA, 27 -> AB
function getExcelColumnLabel(index) {
  let label = '';
  while (index >= 0) {
      label = String.fromCharCode((index % 26) + 65) + label;
      index = Math.floor(index / 26) - 1;
  }
  return label;
}

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
      console.time("Processing log");
      groupedData = extractAndGroupById(logContent);
      console.timeEnd("Processing log");
      

  };

});
