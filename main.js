#include "extendscript.csv.jsx";

var parentFold = (new File($.fileName)).parent;
var outputFold = parentFold + '/output/';

var docRef = app.activeDocument;

var frontLayerSet = docRef.layerSets.getByName('正面');
var backLayerSet = docRef.layerSets.getByName('背面');



var DEBUG = false;
var data = CSV.toJSON('', true, ',');

for (var index in data) {
    var catObj = data[index];
    for (var key in catObj) {
        if (key == '序号') {
            continue;
        }
        changeTextLayerContent(docRef, key + 'T', catObj[key]);
    }
    var grade = catObj['等级'];
    switchGradeLayerSet(grade);
    // 正面
    showFrontLayerSet(true);
    saveJPG(outputFold + index + 'A.jpg', 12);

    // 背面
    showFrontLayerSet(false);
    saveJPG(outputFold + index + 'B.jpg', 12);
}




/**
 * Change text content of a specific named Text Layer to a new text string.
 *
 * @param {Object} doc - A reference to the document to change.
 * @param {String} layerName - The name of the Text Layer to change.
 * @param {String} newTextString - New text content for the Text Layer.
 */

function changeTextLayerContent(doc, layerName, newTextString) {
    for (var i = 0, max = doc.layers.length; i < max; i++) {
        var layerRef = doc.layers[i];
        if (layerRef.typename === "ArtLayer") {
            if (layerRef.name === layerName && layerRef.kind === LayerKind.TEXT) {
                layerRef.textItem.contents = newTextString;
            }
        } else {
            changeTextLayerContent(layerRef, layerName, newTextString);
        }
    }
}


/**
 * A quick way to change text content.
 *
 * @param {Layerset} layerSet- parent layer set.
 * @param {String} layerName - The name of the Text Layer to change.
 * @param {String} newTextString - New text content for the Text Layer.
 */

function changeTextLayerContentByName(layerSet, layerName, newTextString) {
    var textLayer = layerSet.artLayers.getByName(layerName);
    textLayer.textItem.contents = newTextString;
}


function showFrontLayerSet(isFrontVisible) {
    frontLayerSet.visible = isFrontVisible;
    backLayerSet.visible = !isFrontVisible;
}

// Pay too much to code clean. Need change.
function switchGradeLayerSet(grade) {
    var gradeList = ['A', 'B', 'C', 'S'];
    for (i = 0; i < gradeList.length; i++) {
        var gradeEach = gradeList[i];
        var gradeLayer = frontLayerSet.layerSets.getByName(gradeEach);
        gradeLayer.visible = gradeEach == grade ? true : false;
    }

}

function saveJPG(saveFile, jpegQuality) {

    saveFile = (saveFile instanceof File) ? saveFile : new File(saveFile);

    jpegQuality = jpegQuality || 10;

    var jpgSaveOptions = new JPEGSaveOptions();

    jpgSaveOptions.embedColorProfile = true;

    jpgSaveOptions.formatOptions = FormatOptions.STANDARDBASELINE;

    jpgSaveOptions.matte = MatteType.NONE;

    jpgSaveOptions.quality = jpegQuality;

    activeDocument.saveAs(saveFile, jpgSaveOptions, true, Extension.LOWERCASE);

}