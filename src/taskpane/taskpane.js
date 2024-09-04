/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
document.getElementById('fileInput').addEventListener('change', function(event) {
  var file = event.target.files[0]; // 获取用户选择的文件
  if (!file) {
      return;
  }

  var reader = new FileReader(); // 创建FileReader对象
  reader.readAsText(file); // 以文本形式读取文件
  reader.onload = function(e) {
      var content = e.target.result; // 读取文件内容
      console.log(content); // 打印到控制台
      // 这里可以对content进行进一步处理，例如显示在页面上或发送到服务器
  };

});


Office.onReady(() => {
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}
