(function () {
    "use strict";
    Office.onReady()
        .then(function () {
            $(document).ready(function () {
                $('#ok-button').click(sendStringToParentPage);
                $('#cancel-button').click(sendStringToParentPageCancel);
                var item = JSON.parse(localStorage.getItem("confirmationDialog"));
                // setMessage(item);
            });
        });

    function setMessage(item) {
        if (item && item.length > 0) {
            const fieldsMessage = item.reduce((accum, val) => {
                return accum + `<br>${val.display}<br>`
            }, "")
            document.getElementById("messageDialog").innerHTML +=
                `<br><br>The following fields are may have been added or removed. <br> ${fieldsMessage}`
        }
    }

    function sendStringToParentPage() {
        const obj = {
            type: "yes"
        }
        Office.context.ui.messageParent(JSON.stringify(obj));
    }
    function sendStringToParentPageCancel() {
        const obj = {
            type: "no"
        }
        Office.context.ui.messageParent(JSON.stringify(obj));
    }
}());