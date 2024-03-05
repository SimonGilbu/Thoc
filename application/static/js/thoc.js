"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
// GLOBAL CONST
const _POSTSCRIPT_POINT = 0.0352777778;
// SECTION: Utility
/*
    setMessage(text: string)
    
    the function sets the text as a message on the specified "message" element
    in the taskpane html-view, should take in a string that can be displayed to
    the user
*/
function setMessage(text) {
    const msgBox = document.getElementById("setMessageDisplayElement");
    if (msgBox) {
        msgBox.innerHTML = text;
    }
}
// SECTION: Size and position:
function changeShapePositionTop(ev) {
    PowerPoint.run((context) => __awaiter(this, void 0, void 0, function* () {
        var _a;
        const htmlInput = ev.target;
        // each point is 0.35mm
        // input is by default cm, so we need to do some conversion
        // to make the math math
        let changeValue = (_a = parseFloat(htmlInput.value)) !== null && _a !== void 0 ? _a : 0;
        if (changeValue != 0) {
            changeValue = changeValue / _POSTSCRIPT_POINT;
        }
        if (!changeValue) {
            console.log("cant change the value");
            return;
        }
        const shape = context.presentation.getSelectedShapes().getItemAt(0);
        shape.top = changeValue;
    }));
}
function changeShapePositionLeft() {
    PowerPoint.run((context) => __awaiter(this, void 0, void 0, function* () {
        const shape = context.presentation.getSelectedShapes().getItemAt(0);
    }));
}
function changeShapeSizeHeight() {
    PowerPoint.run((context) => __awaiter(this, void 0, void 0, function* () {
        const shape = context.presentation.getSelectedShapes().getItemAt(0);
    }));
}
function changeShapeSizeWidth() {
    PowerPoint.run((context) => __awaiter(this, void 0, void 0, function* () {
        const shape = context.presentation.getSelectedShapes().getItemAt(0);
    }));
}
// SECTION: main
function registerOfficeEventHandlers() {
    // Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, MyHandler);
}
function registerOnChangeHandlers() {
    // Change handlers for size/position of the active shape;
    const shapePositionTopInput = document.getElementById("shapePositionTopInput");
    if (shapePositionTopInput) {
        shapePositionTopInput.onchange = (event) => { changeShapePositionTop(event); };
    }
    const shapePositionLeftInput = document.getElementById("shapePositionLeftInput");
    if (shapePositionLeftInput) {
        shapePositionLeftInput.onchange = changeShapePositionLeft;
    }
    const shapeSizeHeightInput = document.getElementById("shapeSizeHeightInput");
    if (shapeSizeHeightInput) {
        shapeSizeHeightInput.onchange = changeShapeSizeHeight;
    }
    const shapeSizeWidthInput = document.getElementById("shapeSizeWidthInput");
    if (shapeSizeWidthInput) {
        shapeSizeWidthInput.onchange = changeShapeSizeWidth;
    }
}
function registerOnClickHandlers() {
    const renameShapesButton = document.getElementById("renameShapesButton");
    if (renameShapesButton) { }
}
Office.onReady(() => {
    registerOfficeEventHandlers();
    registerOnChangeHandlers();
    registerOnClickHandlers();
});
