const forecastsMax = 4;
const tableHeaders = {
    'groupName': 'Название группы',
    'groupNumber': 'Номер группы',
    'groupPhrases': [{
        'phrase': 'Ключевая фраза',
        'forecastRates': [
            "Прогноз ставки 1",
            "Прогноз ставки 2",
            "Прогноз ставки 3",
            "Прогноз ставки 4",
            "Прогноз ставки 5",
        ],
        'writtenOffPrice': [
            "Спис. цена, руб. 1",
            "Спис. цена, руб. 2",
            "Спис. цена, руб. 3",
            "Спис. цена, руб. 4",
            "Спис. цена, руб. 5",
        ],
        'forecastTraffic': 'Прогноз трафика'
    }]
}

exportToExcel = {
    "exportData": [],
    "exportToExcelBtn": null,
    "tableToExport": null,
    "tableBody": null,
    "init": function () {
        $("html").append("<table style='display: none' id='tableToExport'>");
        $("html").append("<button type='button' id='exportToExcelBtn'>Скачать XLS");
        this.tableToExport = $("#tableToExport");
        this.tableToExport.append("<tbody>");
        this.tableBody = $("tbody", this.tableToExport);

        this.exportToExcelBtn = $("#exportToExcelBtn");
        this.exportToExcelBtn.css({
            'position': 'sticky',
            'bottom': '40px',
            'left': '10px',
            'cursor': 'pointer',
        });
        this.exportToExcelBtn.click(() => this.process());
    },
    "process": function () {
        this.exportToExcelBtn.prop("disabled", true);
        this.exportData = [];
        this.addTableRows(tableHeaders);
        this.prepareData();
        this.mappingData();
        this.downloadExcelFile();
        this.exportToExcelBtn.prop("disabled", false);
    },
    "addTableRows": function (data) {
        let _this = this;
        data.groupPhrases.map(function (e){
            let oneRow = $("<tr></tr>")
                .append("<td>" + data.groupName)
                .append("<td>" + data.groupNumber)
                .append("<td>" + e.phrase);
            for (let i = 0; i < forecastsMax; i++) {
                oneRow.append("<td>" + (e.forecastRates[i] || ''))
                    .append("<td>" + (e.writtenOffPrice[i] || ''));
            }
            oneRow.append("<td>" + e.forecastTraffic);
            _this.tableBody.append(oneRow);
        })
    },
    "prepareData": function () {
        let _this = this;
        $(".b-campaign-group").each(function (){
            let groupInfo = {
                'groupName': $(".b-campaign-group__group-title", $(this)).text(),
                'groupNumber': $(".b-campaign-group__group-number", $(this)).text(),
                'groupPhrases': [],
            };
            $(".b-group-phrase", $(this)).each(function () {
                let phraseInfo = {
                    'phrase': $(".b-phrase-key-words", $(this)).text(),
                    'forecastRates': [],
                    'writtenOffPrice': [],
                    'forecastTraffic': $(".b-group-phrase__current-traffic-volume", $(this)).text()
                }
                $(".b-group-phrase__price-value_type_bid", $(this)).each(function (){
                    let textValue = $(this).text().split(' ').join('');
                    phraseInfo.forecastRates.push(parseFloat(textValue) || textValue);
                });
                $(".b-group-phrase__price-value_type_amnesty", $(this)).each(function (){
                    let textValue = $(this).text().split(' ').join('');
                    phraseInfo.writtenOffPrice.push(parseFloat(textValue) || textValue)
                });
                groupInfo.groupPhrases.push(phraseInfo);
            });
            _this.exportData.push(groupInfo);
        });
    },
    "mappingData": function () {
        let _this = this;
        this.exportData.map((e) => _this.addTableRows(e));
    },
    "downloadExcelFile": function () {
        let _this = this;
        let filename = $("title").text();
        let wb = XLSX.utils.table_to_book(
            _this.tableToExport[0], {sheet:"List 1"}
        );
        return XLSX.writeFile(wb, filename + '.xls');
    }
}
$(document).ready(exportToExcel.init());