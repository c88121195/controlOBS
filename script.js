Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // Excel 已經準備好
        console.log("Excel is ready!");
        // 在這裡呼叫您的函數來操作已打開的 Excel 檔案
        callOpenExcel();


        $("#register-event-handlers").click(() => { tryCatch(registerEventHandlers); });

        var obs = new OBSWebSocket();

        async function registerEventHandlers() {
            await Excel.run(async (context) => {
                //connect to OBS Websocket localhost
                //Get websocket connection info
                //Enter the websocketIP address
                const websocketIP = document.getElementById("IP").value;
                //Enter the OBS websocket port number
                const websocketPort = document.getElementById("Port").value;
                //Enter the OBS websocket server password
                const websocketPassword = document.getElementById("PW").value;

                console.log(`ws://${websocketIP}:${websocketPort}`);
                // connect to OBS websocket
                try {
                    const { obsWebSocketVersion, negotiatedRpcVersion } = await obs.connect(
                        `ws://${websocketIP}:${websocketPort}`,
                        websocketPassword,
                        {
                            rpcVersion: 1
                        }
                    );
                    console.log(`Connected to server ${obsWebSocketVersion} (using RPC ${negotiatedRpcVersion})`);
                } catch (error) {
                    console.error("Failed to connect", error.code, error.message);
                }
                obs.on("error", (err) => {
                    console.error("Socket error:", err);
                });

                //Get all the Text Sources from OBS
                const textSource = await obs.call("GetInputList", { inputKind: "text_gdiplus_v2" });

                // Add a 'on changed' event handler for the workbook.
                let tables = context.workbook.tables;
                tables.onChanged.add(onChange);
                console.log("A handler has been registered for the table collection onChanged event");
                await context.sync();
                setup();
            });
        }

        async function setup() {
            await Excel.run(async (context) => {
                //Delete the 'Sample' sheet
                // context.workbook.worksheets.getItemOrNullObject("Sample").delete();
                //Add a sheet named 'Sample
                const sheet = context.workbook.worksheets.add("Sample");

                //Create a Table to store the Text sources
                createSourceTable(sheet);
                //sync changes
                await context.sync();
                //Get the current content for each Text source
                getSourceSettings();

                let format = sheet.getRange().format;
                format.autofitColumns();
                format.autofitRows();

                sheet.activate();
                await context.sync();
            });
        }
        /** 
        * @param {Excel.Wordsheet} sheet
        */

        async function createSourceTable(sheet) {
            await Excel.run(async (context) => {
                //get the 'Sample' worksheet
                let sheet = context.workbook.worksheets.getItem("Sample");

                //create a table names SourceTable with 2 columns
                let sourceTable = sheet.tables.add("A18:B18", true);
                sourceTable.name = "SourceTable";
                sourceTable.getHeaderRowRange().values = [["inputName", "Setting"]];

                //Get a list of Text sources from OBS
                let textSources = await obs.call("GetInputList", { inputKind: "text_gdiplus_v2" });

                //transform the list of Text Sources to fit in the table
                textSources = Object.values(textSources).flat(1);
                let newData = textSources.map((item) => [item.inputName, item.inputKind]);

                //add the Text sources to the table
                sourceTable.rows.add(null, newData);

                //adjust table column widths
                sheet.getUsedRange().format.autofitColumns();
                sheet.getUsedRange().format.autofitRows();

                sheet.activate();
                await context.sync();
            });
        }

        async function getSourceSettings() {
            await Excel.run(async (context) => {
                //get the 'Sample' worksheet amd Source Table
                const sheet = context.workbook.worksheets.getItem("Sample");
                const sourceTable = sheet.tables.getItem("SourceTable");

                //get values from the Source Table
                const bodyRange = sourceTable.getDataBodyRange().load("values");
                const tableAddress = sourceTable.getDataBodyRange().load("address");

                //to read the table values, use the sync() method
                await context.sync();

                //read the Source table values
                const bodyValues = bodyRange.values;
                sheet.getRange(tableAddress.address).values = bodyValues;
                let sourceSetting;
                //for each Text Source get the current content
                for (let i = 0; i < bodyValues.length; i++) {
                    //get settings from OBS
                    sourceSetting = await obs.call("GetInputSettings", { inputName: bodyValues[i][0] });
                    //transform the settings to fit in the table
                    sourceSetting = Object.values(sourceSetting).flat(1);
                    console.log(sourceSetting);
                    let newData = sourceSetting.map((item) => [item.text]);
                    console.log(newData);
                    bodyValues[i][1] = newData[1][0];
                }
            });
        }

        async function onChange(event) {
            await Excel.run(async (context) => {
                //Get changed cell value and send it to OBS
                let table = context.workbook.tables.getItem(event.tableId);
                let worksheet = context.workbook.worksheets.getItem(event.worksheetId);
                let tablename = table.load("name");

                await context.sync();

                console.log(
                    "Handler for table collection onChanged event has been triggered. Data changed address: " + event.address
                );

                const newValue = worksheet.getRange(event.address);
                const inputName = newValue.getOffsetRange(0, -1);
                newValue.load("text");
                inputName.load("text");

                await context.sync();

                let textValue = newValue.text.toString();
                let inputValue = inputName.text.toString();

                console.log("Table Id : " + event.tableId);
                console.log("Table Name : " + table.name);

                //set OBS source text
                await obs.call("SetInputSettings", {
                    inputName: inputValue,
                    inputSettings: {
                        text: textValue
                    }
                });
            });
        }

        /** Default helper for invoking an action and handling errors. */
        async function tryCatch(callback) {
            try {
                await callback();
            } catch (error) {
                // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
                console.error(error);
            }
        }
    }
});
// @types/office-js

// core - js@2.4.1 / client / core.min.js
// @types/core-js

// jquery @3.1.1
// @types/jquery@3.3.1

// obs - websocket - js