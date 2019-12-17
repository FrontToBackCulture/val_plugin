(function () {
    "use strict";
    let excelHeaders = [];
    let mapSettings = [];
    let tableDetails = [];
    Office.onReady()
        .then(function () {
            $(document).ready(function () {
                $('#ok-button').click(sendStringToParentPage);
                $('#removeMapping').click(removeMapping);
                $('#mapAll').click(mapAll);
                excelHeaders = localStorage.getItem("headerSet");
                if (excelHeaders && typeof excelHeaders == 'string') {
                    excelHeaders = JSON.parse(excelHeaders)
                }

                mapSettings = JSON.parse(localStorage.getItem("mapSettings"));
                tableDetails = JSON.parse(localStorage.getItem("tableDetails"))
                console.log(tableDetails)
                setUIForMapping()

            });
        });

    function setUIForMapping() {
        var mapList = document.getElementById("excel-list");
        var valOptionList = document.getElementById("val-list");
        mapList.innerHTML += `<span class="ms-ListItem-secondaryText">ExcelHeader</span> `
        valOptionList.innerHTML += `<span class="ms-ListItem-secondaryText">VAL Table Field</span> `
        _.forEach(excelHeaders, header => {
            mapList.innerHTML += `<li class="ms-ListItem" tabindex="0">  <span class="ms-ListItem-secondaryText">${header}</span> </li>`
        })

        _.forEach(excelHeaders, header => {
            valOptionList.innerHTML += `<li class="val-item" tabindex="0">  <select id="select_${header}" class="headerText"></select> </li>`
        })

        // let row = mapTable.insertRow(0);
        // _.forEach(excelHeaders, (head, index) => {
        //     var cell = row.insertCell(index);
        //     cell.innerHTML = head
        // })

        // let row2 = mapTable.insertRow(1);
        // _.forEach(excelHeaders, (head, index) => {
        //     var cell = row2.insertCell(index);
        //     cell.innerHTML = `<select id="select_${head}" class="headerText"></select>`
        // })

        tableDetails.fields = _.filter(tableDetails.fields, field => {
            if (field.column_name != "updated_date") {
                return field;
            }
        })
        let optionSet = _.cloneDeep(tableDetails.fields);
        optionSet.unshift({ display: "Not Mapped", column_name: "None" })
        _.map(excelHeaders, (header, index) => {
            var select = document.getElementById(`select_${header}`);
            console.log(select)
            _.map(optionSet, head => {
                select.options[select.options.length] = new Option(head.display, head.column_name)
            })
        })

        // // checkForMapping 
        // console.log(mapSettings)
        // console.log(optionSet)
        if (mapSettings && mapSettings.length > 0) {
            _.map(mapSettings, set => {
                if (set.valField && set.valField != "None") {
                    let optionSetter = _.find(tableDetails.fields, { 'column_name': set.valField })
                    if (optionSetter) {
                        let select = document.getElementById(`select_${set.header}`);
                        select.value = optionSetter.column_name;
                    }
                }
            })
        }
    }

    function sendStringToParentPage() {
        let mappingArr = []
        excelHeaders.map((head, index) => {
            var select = document.getElementById(`select_${head}`).value;
            mappingArr.push({ header: head, valField: select })
        })
        console.log(mappingArr)
        mappingArr = JSON.stringify(mappingArr)
        Office.context.ui.messageParent(mappingArr);
    }

    function removeMapping() {
        var valOptionList = document.getElementById("val-list");
        console.log(valOptionList)
        excelHeaders.map(headers => {
            let select = document.getElementById(`select_${headers}`);
            // select.value = { display: "Not Mapped", column_name: "None" };
            select.value = "None"
        })
    }

    function mapAll() {
        excelHeaders.map(headers => {
            let optionSetter = _.find(tableDetails.fields, { 'display': headers })
            let select = document.getElementById(`select_${headers}`);
            if (optionSetter) {
                select.value = optionSetter.column_name
            }
        })

    }






}());