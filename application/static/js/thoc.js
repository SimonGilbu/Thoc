Office.onReady(function () {
    addSelectionChangedEventHandler();
});
function addSelectionChangedEventHandler() {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, MyHandler);
}
function MyHandler(eventArgs) {
    setMessage('Event raised: ' + eventArgs.type);
}
function setMessage(text) {
    var msgBox = document.getElementById("activeSelectionName");
    if (msgBox) {
        msgBox.innerHTML = text;
    }
}
