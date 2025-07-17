jQuery.sap.require("materialdocument.model.jszip");
jQuery.sap.require("materialdocument.model.xlsx");
sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/core/library",
    "sap/ui/export/library",
    "sap/ui/export/Spreadsheet",
    "sap/m/Dialog",
    "sap/m/library",
    "sap/m/Button",
    "sap/m/Text",
    "sap/ui/model/Sorter",
    "sap/ui/model/Filter",
    "sap/m/SearchField",
    "sap/ui/table/Column",
    "sap/m/Column",
    "sap/m/Label",
    "sap/ui/model/type/String",
    "sap/ui/comp/library",
    "sap/ui/model/FilterOperator",
    "sap/ui/core/Fragment",
    "sap/ui/core/message/Message",
    "sap/m/MessageBox"
], (Controller,coreLibrary, exportLibrary, Spreadsheet, Dialog, mobileLibrary, Button, Text, Sorter, Filter,
    SearchField, UIColumn, MColumn, Label, TypeString, compLibrary, FilterOperator, Fragment, Message,MessageBox) => {
    "use strict";
    var uploadModel = new sap.ui.model.json.JSONModel();
    var oBusyDialogAdd;
    var SucessData = [];
    var EdmType = exportLibrary.EdmType;
    return Controller.extend("materialdocument.controller.Upload_Material", {
        onInit() {
            this._excelStatusMap = [];
        },
        onUpload: function(oEvent){
            var that = this;
            this.getView().byId('logStrip').setVisible(false);
            that._import(oEvent.getParameter("files") && oEvent.getParameter("files")[0]);
        },
        _import : function(file){
            var that = this;
            var excelData = {};
            var uploadData = {};
            var TableuploadData = [];
            if (file && window.FileReader) {
                var reader = new FileReader();
                reader.onload = function (e) {
                    var data = e.target.result;
                    var workbook = XLSX.read(data, {
                        type: 'binary'
                    });
                    workbook.SheetNames.forEach(function (sheetName) {
                        // Here is your object for every sheet in workbook
                        excelData = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);

                    });
                    for (let i = 0; i < excelData.length; i++) {
                        var currentData = excelData[i];
                            if (currentData) {
                                var uploadData = {};
                                uploadData.DocumentNo = excelData[i].DocumentNo;
                                uploadData.PostingDate = "2025-06-18T00:00:00"; //excelData[i].PostingDate;
                                uploadData.GoodsMovementCode = "01";//excelData[i].GoodsMovementCode;
                                uploadData.Material = excelData[i].Material;
                                uploadData.Plant = excelData[i].Plant;
                                uploadData.StorageLocation = excelData[i].StorageLocation;
                                uploadData.GoodsMovementType = excelData[i].GoodsMovementType;
                                uploadData.PurchaseOrder = excelData[i].PurchaseOrder;
                                uploadData.PurchaseOrderItem = excelData[i].PurchaseOrderItem;
                                uploadData.GoodsMovementRefDocType = excelData[i].GoodsMovementRefDocType;
                                uploadData.EntryUnit = excelData[i].EntryUnit;
                                uploadData.QuantityInEntryUnit = excelData[i].QuantityInEntryUnit;
                                TableuploadData.push(uploadData);
                            }
                    }

                    //set submit button enable
                    if (TableuploadData.length > 0) {
                       // that.getView().byId("id_dialoguploadcreate").setEnabled(true);
                    }
                    // Setting the data to the Upload Table model 
                    uploadModel.setData(TableuploadData);
                    that.getView().setModel(uploadModel, "MaterialUpload");
                    uploadModel.refresh(true);

                };
                reader.onerror = function (ex) {
                    console.log(ex);
                };
                reader.readAsBinaryString(file);
            }
        },
        onDownloadResult: function () {
            const data = this._excelStatusMap || [];
        
            if (data.length === 0) {
                MessageBox.warning("No data available for export.");
                return;
            }
        
            const oSettings = {
                workbook: {
                    columns: this._downloadStatusColumnConfig()
                },
                dataSource: data,
                fileName: 'Material_Upload_Status.xlsx'
            };
        
            const oSheet = new Spreadsheet(oSettings);
            oSheet.build().finally(() => {
                oSheet.destroy();
            });
        },
        _downloadStatusColumnConfig: function () {
            return [
                { label: 'Document No', type: EdmType.String, property: 'DocumentNo' },
                { label: 'Posting Date', type: EdmType.String, property: 'PostingDate' },
                { label: 'Goods Movement Code', type: EdmType.String, property: 'GoodsMovementCode' },
                { label: 'Material', type: EdmType.String, property: 'Material' },
                { label: 'Plant', type: EdmType.String, property: 'Plant' },
                { label: 'Storage Location', type: EdmType.String, property: 'StorageLocation' },
                { label: 'Goods Movement Type', type: EdmType.String, property: 'GoodsMovementType' },
                { label: 'Purchase Order', type: EdmType.String, property: 'PurchaseOrder' },
                { label: 'Purchase Order Item', type: EdmType.String, property: 'PurchaseOrderItem' },
                { label: 'Goods Movement Ref Doc Type', type: EdmType.String, property: 'GoodsMovementRefDocType' },
                { label: 'Entry Unit', type: EdmType.String, property: 'EntryUnit' },
                { label: 'Quantity In Entry Unit', type: EdmType.String, property: 'QuantityInEntryUnit' },
                { label: 'Status', type: EdmType.String, property: 'Status' }
            ];
        },                
        postMaterial: async function (oEvent) {
            const oBusyDialogAdd = new sap.m.BusyDialog({
                title: "Processing.....",
                text: "Migrating Material Documents... Please Wait"
            });
            oBusyDialogAdd.open();
        
            const oModel = this.getOwnerComponent().getModel();
            const oTable = this.getView().byId("id_table");
            const oEntries = oTable.getSelectedItems();
        
            if (oEntries.length === 0) {
                MessageBox.warning("Please select at least one item to post.");
                oBusyDialogAdd.close();
                return;
            }
        
            const groupedPayloads = {};
        
            for (let i = 0; i < oEntries.length; i++) {
                const oCells = oEntries[i].getCells();
                const docId = oCells[0].getText();
        
                if (!groupedPayloads[docId]) {
                    groupedPayloads[docId] = {

                        PostingDate: oCells[1].getText(),
                        GoodsMovementCode: oCells[2].getText(),
                        to_MaterialDocumentItem: []
                    };
                }
        
                const item = {
                    Material: oCells[3].getText(),
                    Plant: oCells[4].getText(),
                    StorageLocation: oCells[5].getText(),
                    GoodsMovementType: oCells[6].getText(),
                    PurchaseOrder: oCells[7].getText(),
                    PurchaseOrderItem: oCells[8].getText(),
                    GoodsMovementRefDocType: oCells[9].getText(),
                    EntryUnit: oCells[10].getText(),
                    QuantityInEntryUnit: oCells[11].getText()
                };
        
                groupedPayloads[docId].to_MaterialDocumentItem.push(item);
            }
        
            const documentIds = Object.keys(groupedPayloads);
            console.log(groupedPayloads);
            this._excelStatusMap = []; 
            for (const docId of documentIds) {
                const payload = groupedPayloads[docId];
                const materialItems = payload.to_MaterialDocumentItem;
                try {
                    const response = await this._triggerMaterialDocument(oModel, payload);
                    const statusMsg = response && response.statusCode === "201"
                        ? "Success: Material Document No - " + (response.data?.MaterialDocument || docId)
                        : "Failed: Unknown Error";
                
                    // Add status per item
                    materialItems.forEach(item => {
                        this._excelStatusMap.push({
                            DocumentNo: docId,
                            PostingDate: payload.PostingDate,
                            GoodsMovementCode: payload.GoodsMovementCode,
                            ...item,
                            Status: statusMsg
                        });
                    });
                
                    if (response && response.statusCode === "201") {
                      //  MessageBox.success(statusMsg);
                    } else {
                        //MessageBox.error("Failed to create Material Document for ID: " + docId);
                    }
                } catch (error) {
                    const errMsg = error?.responseText
                        ? JSON.parse(error.responseText)?.error?.message?.value
                        : error.message;
                
                    materialItems.forEach(item => {
                        this._excelStatusMap.push({
                            DocumentNo: docId,
                            PostingDate: payload.PostingDate,
                            GoodsMovementCode: payload.GoodsMovementCode,
                            ...item,
                            Status: "Failed: " + errMsg
                        });
                    });
                
                    //MessageBox.error("Error for Document ID " + docId + ": " + errMsg);
                }
            }
            this.byId("logStrip").setVisible(true);

            oBusyDialogAdd.close();
        }
        ,
         //var oUrl = this.getBaseURL() + `/sap/opu/odata/sap/API_MATERIAL_DOCUMENT_SRV/$metadata`;
               /* var oSettings = {
                           "url": this.getBaseURL() + `/sap/opu/odata/sap/API_MATERIAL_DOCUMENT_SRV/$metadata`,
                            "method": "GET",
                            async: false,
                            "headers": {
                                "X-CSRF-Token" : "Fetch"
                            },
                            "Accept": "application/json",
                            "contentType": "application/json"
                }; */
                /*console.log(oSettings);
                let oToken = await that._triggerCSRF(oUrl);
                var oSettingsURL = {
                    "url": this.getBaseURL() + `/sap/opu/odata/sap/API_MATERIAL_DOCUMENT_SRV/A_MaterialDocumentHeader`,
                     "method": "POST",
                     async: false,
                     "headers": {
                        "X-CSRF-Token" : oToken
                     },
                     "Accept": "application/json",
                     "contentType": "application/json",
                     "data": JSON.stringify(oPayloadEntries) 
                }; */
               /* await that._triggerMaterial(oSettingsURL).then((oResponse) => {
                    console.log(oResponse);
                }); */
        /*getBaseURL: function () {
            return sap.ui.require.toUrl("materialdocument");
        },*/
        /*_triggerMaterial : async function (oSettingsUrl){
            return new Promise(function (resolve, reject) {
                $.ajax(oSettingsUrl).done(function (oData,oResponse) {
                    resolve(oResponse);
                }).fail(async function (oError) {
                    reject(oError);
                });
            });
        },*/
        /*_triggerCSRF : async function (oSettingsUrl){
            return new Promise(function (resolve, reject) {
                $.ajax({ url: oSettingsUrl ,
                    type: "GET",
                    async: false ,
                    headers: {
                    },
                    beforeSend: function(xhr){ xhr.setRequestHeader("X-CSRF-Token", "Fetch") ; },
                    complete: function(xhr){ 
                        var token = xhr.getResponseHeader("X-CSRF-Token");
                        resolve(token);
                    }
                })
            });
        },*/
        _triggerMaterialDocument : async function (oModel,oPayloadEntries){
            return new Promise(function (resolve, reject) {
                oModel.create("/A_MaterialDocumentHeader", oPayloadEntries, {
                    success: function (data, response) {
                        resolve(response);
                        SucessData = data;
                        console.log(data);
                    },
                    error: function (error) {          
                       reject(error);
                    }
                });
            });
        },
        downloadButton: function(oEvent){
            var aCols, oSettings, oRowBinding, oSheet, oEmployeeModel;
                oRowBinding = this.getView().getModel("templateModel").getProperty('/Template');
                aCols = this._downloadColumnConfig();

                oSettings = {
                    workbook: {
                        columns: aCols,
                        hierarchyLevel: 'Level'
                    },
                    dataSource: oRowBinding,
                    fileName: 'Marerial_Document_Upload.xlsx'
                };

                oSheet = new Spreadsheet(oSettings);
                oSheet.build().finally(function () {
                    oSheet.destroy();
                });
        },
        _downloadColumnConfig: function () {
            var aCols = [];

            aCols.push({
                label: 'DocumentNo',
                type: EdmType.String,
                property: 'DocumentNo'
            });


            aCols.push({
                label: 'PostingDate',
                type: EdmType.String,
                property: 'PostingDate'
            });

            aCols.push({
                label: 'GoodsMovementCode',
                type: EdmType.String,
                property: 'GoodsMovementCode'
            });

            aCols.push({
                label: 'Material',
                type: EdmType.String,
                property: 'Material'
            });

            aCols.push({
                label: 'Plant',
                type: EdmType.String,
                property: 'Plant'
            });

            aCols.push({
                label: 'StorageLocation',
                type: EdmType.String,
                property: 'StorageLocation'
            });

            aCols.push({
                label: 'GoodsMovementType',
                type: EdmType.String,
                property: 'GoodsMovementType'
            });

            aCols.push({
                label: 'PurchaseOrder',
                type: EdmType.String,
                property: 'PurchaseOrder'
            });

            aCols.push({
                label: 'PurchaseOrderItem',
                type: EdmType.String,
                property: 'PurchaseOrderItem'
            });

            aCols.push({
                label: 'GoodsMovementRefDocType',
                type: EdmType.String,
                property: 'GoodsMovementRefDocType'
            });

            aCols.push({
                label: 'EntryUnit',
                type: EdmType.String,
                property: 'EntryUnit'
            });

            aCols.push({
                label: 'QuantityInEntryUnit',
                type: EdmType.String,
                property: 'QuantityInEntryUnit'
            });

            return aCols;
        }
    });
});