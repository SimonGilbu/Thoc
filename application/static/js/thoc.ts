Office.onReady(() => {
   addSelectionChangedEventHandler();
})


function addSelectionChangedEventHandler() {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, MyHandler);
}

function MyHandler(eventArgs: Office.DocumentSelectionChangedEventArgs) {
    setMessage('Event raised: ' + eventArgs.type);
}

function setMessage(text: string) {
    const msgBox = document.getElementById("activeSelectionName");
    if (msgBox) {
        msgBox.innerHTML = text;
    }
}