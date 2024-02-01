var rootDir =
  "E:/Workspace/Automation/PhotoShop/Generate_Logo_2023_5_18/script";
app.preferences.rulerUnits = Units.PIXELS;

function hexToRGB(hex) {
  var solidColor = new SolidColor();
  if (hex == "Transparent") {
    solidColor.rgb.red = 0;
    solidColor.rgb.green = 0;
    solidColor.rgb.blue = 0;
    solidColor.opacity = 0;
    return solidColor;
  }
  hex = hex.replace("#", "");

  var r = parseInt(hex.substring(0, 2), 16);
  var g = parseInt(hex.substring(2, 4), 16);
  var b = parseInt(hex.substring(4, 6), 16);

  solidColor.rgb.red = r;
  solidColor.rgb.green = g;
  solidColor.rgb.blue = b;
  return solidColor;
}

function OpenFile(filePath) {
  var file = new File(filePath);
  if (file.open("r")) {
    var fileData = file.read();
    file.close();
    return fileData;
  } else {
    $.writeln("An error occurred: " + file.error);
    return null;
  }
}

function CSVtoArray(str) {
  var delimiter = ",";
  var quotes = '"';
  var newElements = [];
  for (var i = 0; i < str.length; ++i) {
    var startIndex = 0;
    if (str[i] == quotes) {
      var temp = "";
      i++;
      while (i < str.length) {
        if (str[i] == quotes && i + 1 < str.length && str[i + 1] == quotes) {
          // double quotes
          temp += '""';
          i += 2;
          continue;
        }
        if (str[i] == quotes && i + 1 < str.length && str[i + 1] == delimiter) {
          i += 1;
          break;
        }
        temp += str[i];
        i++;
      }
      newElements.push(temp);
    } else {
      startIndex = i;
      var temp = "";
      while (str[i] != delimiter && i < str.length) {
        temp += str[i];
        i++;
      }
      newElements.push(temp);
    }
  }

  return newElements;
}

function GetSimilarityStrings(s1, s2) {
  var longer = s1;
  var shorter = s2;
  if (s1.length < s2.length) {
    longer = s2;
    shorter = s1;
  }
  var longerLength = longer.length;
  if (longerLength == 0) {
    return 1.0;
  }
  return (
    (longerLength - editDistance(longer, shorter)) / parseFloat(longerLength)
  );
}

function editDistance(s1, s2) {
  s1 = s1.toLowerCase();
  s2 = s2.toLowerCase();

  var costs = new Array();
  for (var i = 0; i <= s1.length; i++) {
    var lastValue = i;
    for (var j = 0; j <= s2.length; j++) {
      if (i == 0) costs[j] = j;
      else {
        if (j > 0) {
          var newValue = costs[j - 1];
          if (s1.charAt(i - 1) != s2.charAt(j - 1))
            newValue = Math.min(Math.min(newValue, lastValue), costs[j]) + 1;
          costs[j - 1] = lastValue;
          lastValue = newValue;
        }
      }
    }
    if (i > 0) costs[s2.length] = lastValue;
  }
  return costs[s2.length];
}

function GetMatchingFont(fontFamily, fontStyle) {
  var maxSim = 0;
  var maxI = 0;
  var fonts = app.fonts;
  for (var i = 0; i < fonts.length; i++) {
    var font = fonts[i];
    var sim1 = GetSimilarityStrings(font.family, fontFamily);
    var sim2 = GetSimilarityStrings(font.style, fontStyle);
    if (maxSim < sim1 + sim2) {
      maxSim = sim1 + sim2;
      maxI = i;
    }
  }
  return fonts[maxI];
}

function GetParagrahValue(paragraphString) {
  paragraphString = paragraphString.toLowerCase();
  switch (paragraphString) {
    case "left align text":
      return Justification.LEFT;
    case "center text":
      return Justification.CENTER;
    case "right align text":
      return Justification.RIGHT;
    case "justify last left":
      return Justification.LEFTJUSTIFIED;
    case "justify last centered":
      return Justification.CENTERJUSTIFIED;
    case "justify all":
      return Justification.FULLYJUSTIFIED;
  }
  return Justification.LEFT;
}

function selectLayerById(id) {
  var desc1 = new ActionDescriptor();
  var ref1 = new ActionReference();
  ref1.putIdentifier(charIDToTypeID("Lyr "), id);
  desc1.putReference(charIDToTypeID("null"), ref1);
  executeAction(charIDToTypeID("slct"), desc1, DialogModes.NO);
}

function GetXLogoFileColumn(header, headerData) {
  for (var i = 0; i < headerData.length; i++) {
    if (header == headerData[i]) return i;
  }
  return -1;
}

function GetXLogoSize(filename) {
  var arr = filename.split(/[_x.]/);
  var width = parseInt(arr[1]);
  var height = parseInt(arr[2]);
  return {
    width: width,
    height: height,
  };
}

function GenerateSaveFileName(productType, saying, colorFileName) {
  var currentDate = new Date();
  var datetimeString =
    currentDate.getDate() +
    "-" +
    (currentDate.getMonth() + 1) +
    "-" +
    currentDate.getFullYear();
  var replacedSaying = saying.replace(/[^\w\s]/gi, "");

  var folderPath = rootDir + "/exports/" + datetimeString + "/" + productType;
  var folder = new Folder(folderPath);
  // Check if the folder doesn't exist
  if (!folder.exists) {
    // Create the folder
    folder.create();
  }
  var fileName = encodeURI(replacedSaying + "_" + colorFileName + ".png");
  return folderPath + "/" + fileName;
}

function GetMasterKey(csvType, csvData) {
  var data = csvData;
  if (typeof data == "string") data = CSVtoArray(data);

  switch (csvType) {
    case "saying":
      return parseFloat(data[2]);
    case "product":
      return parseFloat(data[1]);
  }
}

function ReplaceSayingText(sayingText) {
  sayingText = sayingText.replace(/"""/g, '"');
  sayingText = sayingText.replace(/""/g, '"');
  var result = sayingText.replace(/\\r\\n/g, "\r");

  return result;
}

function MakeProduct(productHeader, productRow, colorRow, sayingRow) {
  var sayingData = CSVtoArray(sayingRow);

  var productData = CSVtoArray(productRow);
  var colorData = CSVtoArray(colorRow);

  var productHeaderData = CSVtoArray(productHeader);
  var sayingText = ReplaceSayingText(sayingData[1]);
  var sayingFileName = sayingData[0];
  var productType = productData[0];
  var canvasWidth = parseInt(productData[2]);
  var canvasHeight = parseInt(productData[3]);
  var canvasDpi = parseInt(productData[4]);
  var canvasColor = productData[5];
  var fontName = productData[6];
  var fontSize = parseFloat(productData[7]);
  var lineHeight = parseFloat(productData[8]);
  var lineYAxis = parseFloat(productData[9]);
  var fontWeight = parseInt(productData[10]);
  var fontWidth = parseInt(productData[11]);
  var fontSlant = parseInt(productData[12]);
  var fontStyle = productData[13];
  var verticalScale = parseInt(productData[14]);
  var xAxis = parseInt(productData[15]);
  var paragrapth = productData[16];
  var xLogoXAxis = parseInt(productData[17]);
  var xLogoYAxis = parseInt(productData[18]);
  var scaleValue = 1.0155;
  var fontColor = colorData[0];
  var colorFileName = colorData[2];
  var xLogoTableHeaderName = colorData[1];
  var xLogoFileColumn = GetXLogoFileColumn(
    xLogoTableHeaderName,
    productHeaderData
  );

  // Get xlogo file name by color column header
  if (xLogoFileColumn < 0) return;

  var xLogoFileName = productData[xLogoFileColumn];
  var xLogoSize = GetXLogoSize(xLogoFileName);

  var bgColor = hexToRGB(canvasColor);
  var doc = app.documents.add(canvasWidth, canvasHeight, canvasDpi);

  var textLayer = doc.artLayers.add();
  textLayer.kind = LayerKind.TEXT;

  var aDoc = app.activeDocument;
  var theLayers = aDoc.layers;
  aDoc.activeLayer = theLayers[theLayers.length - 1];
  var aLayer = aDoc.activeLayer;

  if (aLayer.isBackgroundLayer) {
    if (canvasColor == "Transparent") {
      aLayer.remove();
    } else {
      app.activeDocument.selection.fill(
        bgColor,
        ColorBlendMode.NORMAL,
        100,
        false
      );
    }
  }

  // Set the new font attributes
  var matchedFont = GetMatchingFont(fontName, fontStyle);
  textLayer.textItem.font = matchedFont.postScriptName;
  // Set the font weight

  // change line height
  textLayer.textItem.useAutoLeading = false;
  textLayer.textItem.leading = lineHeight;
  textLayer.textItem.color = hexToRGB(fontColor);
  textLayer.textItem.contents = sayingText;
  textLayer.textItem.size = fontSize;
  textLayer.textItem.kind = TextType.PARAGRAPHTEXT;
  textLayer.textItem.verticalScale = verticalScale;
  textLayer.textItem.horizontalScale = 100 * scaleValue;
  var bounds = textLayer.bounds;
  var width = bounds[2].value - bounds[0].value;
  var adjustment = (4 * fontSize) / 3;
  $.writeln("width: " + width);
  var xPos = 0;
  if (!xAxis) {
    xPos = -10;
    textLayer.textItem.justification = GetParagrahValue(paragrapth);
    // Change Width, height
    textLayer.textItem.width = new UnitValue(
      ((canvasWidth + 20) * 72) / canvasDpi,
      "px"
    );
    textLayer.textItem.height = new UnitValue(
      (canvasHeight * 72) / canvasDpi,
      "px"
    );
  } else {
    xPos = xAxis - width / 2 - 10;
    textLayer.textItem.justification = GetParagrahValue(paragrapth);
  }

  textLayer.textItem.position = [xPos, lineYAxis + adjustment];

  // Import XLogo and set position
  var xlogoFile = new File(rootDir + "/assets/xlogos/" + xLogoFileName);
  app.load(xlogoFile); //load it into documents
  backFile = app.activeDocument; //prepare your image layer as active document
  backFile.resizeImage(xLogoSize.width, xLogoSize.height); //resize image into given size i.e 640x480
  backFile.selection.selectAll();
  backFile.selection.copy(); //copy image into clipboard
  backFile.close(SaveOptions.DONOTSAVECHANGES); //close image without saving changes
  doc.paste(); //paste selection into your document
  var xLogoLayer = doc.layers[0];
  xLogoLayer.name = "XLogo Image"; //set your layer's namevar imageTransform = imageLayer.transform;
  xLogoLayer.translate(
    xLogoXAxis - xLogoLayer.bounds[0].value,
    xLogoYAxis - xLogoLayer.bounds[1].value
  );

  var saveFilePath = GenerateSaveFileName(
    productType,
    sayingFileName,
    colorFileName
  );

  var outputFile = new File(saveFilePath); // Replace with the desired output file path
  if (outputFile.exists) {
    // Delete the existing file
    outputFile.remove();
  }

  var exportOptions = new ExportOptionsSaveForWeb();
  exportOptions.format = SaveDocumentType.PNG;
  exportOptions.PNG8 = false; // Use PNG-24 format
  exportOptions.transparency = true; // Preserve transparency
  exportOptions.optimized = true; // Enable optimization

  // Adjust the quality and settings as needed
  exportOptions.quality = 100; // Set the desired quality level (0-100)
  exportOptions.optimization = true; // Apply optimization settings
  exportOptions.includeProfile = false; // Exclude color profile
  exportOptions.interlaced = false; // Disable interlacing
  doc.exportDocument(
    new File(outputFile),
    ExportType.SAVEFORWEB,
    exportOptions
  );

  $.sleep(10);
  app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);
  $.sleep(10);
}

var sayingCsvData = OpenFile(rootDir + "/assets/sayings.csv").split("\n");
var productNameCsvData = OpenFile(rootDir + "/assets/products.csv").split("\n");
var colorCsvData = OpenFile(rootDir + "/assets/colors.csv").split("\n");

var startTime = new Date().getTime();

for (var sayingIndex = 1; sayingIndex < sayingCsvData.length; sayingIndex++) {
  var sayingRow = sayingCsvData[sayingIndex];
  $.writeln("sayingRow: " + sayingRow);

  var masterkey = GetMasterKey("saying", sayingRow);
  if (!masterkey) alert(sayingRow);
  for (
    var productNameIndex = 1;
    productNameIndex < productNameCsvData.length;
    productNameIndex++
  ) {
    var csvFileName = productNameCsvData[productNameIndex];
    if (csvFileName == "") break;
    var productDataArray = OpenFile(
      rootDir + "/assets/products/" + csvFileName
    ).split(/\n/);
    var productHeader = productDataArray[0];
    for (var row = 1; row < productDataArray.length; row++) {
      var productRow = productDataArray[row];
      if (masterkey != GetMasterKey("product", productRow)) continue;
      for (var colorIndex = 1; colorIndex < colorCsvData.length; colorIndex++) {
        var colorRow = colorCsvData[colorIndex];
        MakeProduct(productHeader, productRow, colorRow, sayingRow);
      }
    }
  }
}

var endTime = new Date().getTime();
var exeTime = endTime - startTime;
var exeSeconds = Math.floor((exeTime % 60000) / 1000);
var exeMinutes = Math.floor((exeTime % 3600000) / 60000);
var exeHours = Math.floor(exeTime / 3600000);
alert(
  "Execution time: " + exeHours + " h " + exeMinutes + " m " + exeSeconds + " s"
);
