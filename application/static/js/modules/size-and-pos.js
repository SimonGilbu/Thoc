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
// module functions are namespaced to prevent collisions
var sizeAndPos;
(function (sizeAndPos) {
    /*
        changeShapePositionTop(ev: Event)
        changeShapePositionLeft(ev: Event)
        changeShapeSizeHeight(ev: Event)
        changeShapeSizeWidth(ev: Event)

        Gets the change-event and converts the input value from cm to
        points in order to align the selected item in the ppt presentation,
        
        allows us to set and change the height, bottom left and right pos
        of the shape ^~^

        TODO: this should use a class, type or interface, and feed it
        to a chane size-pos operation, to minimize the calls to powerpoint for
        context-syncing.
    */
    function changeShapePositionTop(ev) {
        PowerPoint.run((context) => __awaiter(this, void 0, void 0, function* () {
            const htmlInput = ev.target;
            let changeValue = parseFloat(htmlInput.value);
            changeValue = changeValue / _POSTSCRIPT_POINT;
            const shape = context.presentation.getSelectedShapes().getItemAt(0);
            shape.top = changeValue;
        }));
    }
    function changeShapePositionLeft(ev) {
        PowerPoint.run((context) => __awaiter(this, void 0, void 0, function* () {
            const htmlInput = ev.target;
            let changeValue = parseFloat(htmlInput.value);
            changeValue = changeValue / _POSTSCRIPT_POINT;
            const shape = context.presentation.getSelectedShapes().getItemAt(0);
            shape.left = changeValue;
        }));
    }
    function changeShapeSizeHeight(ev) {
        PowerPoint.run((context) => __awaiter(this, void 0, void 0, function* () {
            const htmlInput = ev.target;
            let changeValue = parseFloat(htmlInput.value);
            changeValue = changeValue / _POSTSCRIPT_POINT;
            const shape = context.presentation.getSelectedShapes().getItemAt(0);
            shape.height = changeValue;
        }));
    }
    function changeShapeSizeWidth(ev) {
        PowerPoint.run((context) => __awaiter(this, void 0, void 0, function* () {
            const htmlInput = ev.target;
            let changeValue = parseFloat(htmlInput.value);
            changeValue = changeValue / _POSTSCRIPT_POINT;
            const shape = context.presentation.getSelectedShapes().getItemAt(0);
            shape.width = changeValue;
        }));
    }
    /*
        selectionChangeHanlder(ev: Office.EventType)
        
        the method fetches the information from the currently selected object
        and assignts its values to the properties panel that has been created
        in the corespeonding .jinja file.
    */
    function selectionChangeHanlder(ev) {
        PowerPoint.run((context) => __awaiter(this, void 0, void 0, function* () {
            const shapePositionTopInput = document.getElementById("shapePositionTopInput");
            const shapePositionLeftInput = document.getElementById("shapePositionLeftInput");
            const shapeSizeHeightInput = document.getElementById("shapeSizeHeightInput");
            const shapeSizeWidthInput = document.getElementById("shapeSizeWidthInput");
            // sync the context from PPT
            const shapes = context.presentation.getSelectedShapes();
            const selectionCount = shapes.getCount();
            yield context.sync();
            if (selectionCount.value <= 0) {
                shapePositionTopInput.value = "";
                shapePositionLeftInput.value = "";
                shapeSizeHeightInput.value = "";
                shapeSizeWidthInput.value = "";
                return;
            }
            const shape = shapes.getItemAt(0);
            shape.load();
            yield context.sync();
            // math to convert from postsript point to CM
            let top = shape.top * _POSTSCRIPT_POINT;
            let left = shape.left * _POSTSCRIPT_POINT;
            let height = shape.height * _POSTSCRIPT_POINT;
            let width = shape.width * _POSTSCRIPT_POINT;
            shapePositionTopInput.value = top.toFixed(2);
            shapePositionLeftInput.value = left.toFixed(2);
            shapeSizeHeightInput.value = height.toFixed(2);
            shapeSizeWidthInput.value = width.toFixed(2);
        }));
    }
    function registerEventHandlers() {
        // selection handler to propegate atribute panel
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, selectionChangeHanlder);
        // change handlers for size and position;
        const topInput = document.getElementById("shapePositionTopInput");
        topInput.onchange = changeShapePositionTop;
        const leftInput = document.getElementById("shapePositionLeftInput");
        leftInput.onchange = changeShapePositionLeft;
        const heightInput = document.getElementById("shapeSizeHeightInput");
        heightInput.onchange = changeShapeSizeHeight;
        const widthInput = document.getElementById("shapeSizeWidthInput");
        widthInput.onchange = changeShapeSizeWidth;
    }
    // loud our functions for the module once office is ready
    Office.onReady(() => {
        registerEventHandlers();
    });
})(sizeAndPos || (sizeAndPos = {}));
