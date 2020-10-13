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
        this.tableToExport = $("#tableToExport");
        this.tableToExport.append("<tbody>");
        this.tableBody = $("tbody", this.tableToExport);
        this.tableBody.append(this.addTableRow(tableHeaders));

        $("html").append("<button type='button' id='exportToExcelBtn'>Скачать XLS");
        this.exportToExcelBtn = $("#exportToExcelBtn");
        this.exportToExcelBtn.css({
            'position': 'sticky',
            'bottom': '40px',
            'left': '10px',
            'cursor': 'pointer',
        });
        this.exportToExcelBtn.click(() => this.prepareData());
    },
    "addTableRow": function (data) {
        let html = '';
        data.groupPhrases.map(function (e){
            html = $("<tr></tr>");
            html.append("<td>" + data.groupName)
                .append("<td>" + data.groupNumber);
            html.append("<td>" + e.phrase);
            for (let i = 0; i < forecastsMax; i++) {
                html.append("<td>" + (e.forecastRates[i] || ''));
                html.append("<td>" + (e.writtenOffPrice[i] || ''));
            }
            html.append("<td>" + e.forecastTraffic);
        })
        return html;
    },
    "prepareData": function () {
        let _this = this;
        this.exportToExcelBtn.prop("disabled", true);
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

        this.mappingData();
    },
    "mappingData": function () {
        let _this = this;
        this.exportData.map(function (e){
            _this.tableBody.append(_this.addTableRow(e));
        })
        _this.downloadExcelFile();
        _this.exportToExcelBtn.prop("disabled", false);
    },
    "downloadExcelFile": function () {
        let _this = this;
        let filename = $("title").text();

        let wb = XLSX.utils.table_to_book(
            _this.tableToExport[0], {sheet:"List 1"}
        );

        return  XLSX.writeFile(wb, filename + '.xls');
    }
}
exportToExcel.init();