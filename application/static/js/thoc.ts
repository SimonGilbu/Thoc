// GLOBAL CONST
const _POSTSCRIPT_POINT = 0.0352777778;


// SECTION: Utility

/*
    setMessage(text: string)
    
    the function sets the text as a message on the specified "message" element
    in the taskpane html-view, should take in a string that can be displayed to 
    the user
*/
function setMessage(text: string) {
    const msgBox = document.getElementById("setMessageDisplayElement");
    if (msgBox) {
        msgBox.innerHTML = text;
    }
}

// SECTION: Size and position:
function changeShapePositionTop(ev: Event) {
    PowerPoint.run(async (context) => {
        const htmlInput = ev.target as HTMLInputElement;
        let changeValue = parseFloat(htmlInput.value);
        // apply the conversion from CM to PostScript Points
        changeValue = changeValue / _POSTSCRIPT_POINT;
        const shape = context.presentation.getSelectedShapes().getItemAt(0);
        shape.top = changeValue;
    });
}
function changeShapePositionLeft() {
    PowerPoint.run(async (context) => {
        const shape = context.presentation.getSelectedShapes().getItemAt(0);
    });
}
function changeShapeSizeHeight() {
    PowerPoint.run(async (context) => {
        const shape = context.presentation.getSelectedShapes().getItemAt(0);
    });
}
function changeShapeSizeWidth() {
    PowerPoint.run(async (context) => {
        const shape = context.presentation.getSelectedShapes().getItemAt(0);
    });
}


// SECTION: main

function registerOfficeEventHandlers() {
    // Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, MyHandler);
}

function registerOnChangeHandlers() {
    // Change handlers for size/position of the active shape;
    const shapePositionTopInput = document.getElementById("shapePositionTopInput");
    if (shapePositionTopInput) {
        shapePositionTopInput.onchange = (event) => { changeShapePositionTop(event) };
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
})
