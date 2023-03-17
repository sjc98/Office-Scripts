Excel.run(function (context) {
  var shapes = context.workbook.worksheets.getItem("Leaders").shapes;
  var shape = sheet.shapes.getItem("Main");
  var stringResult = shape.getAsImage(Excel.PictureFormat.png);

  return context.sync().then(function () {
    console.log(stringResult.value);
    // Instead of logging, your add-in may use the base64-encoded string to save the image as a file or insert it in HTML.
  });
}).catch(errorHandlerFunction);