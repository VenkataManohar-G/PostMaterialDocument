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
        },
        onUpload: function(oEvent){
            var that = this;
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
        postMaterial: async function(oEvent){
            oBusyDialogAdd = new sap.m.BusyDialog({
                title: "Processing.....",
                text: "Migrating Material Documents.. Please Wait"
            });
            oBusyDialogAdd.open();
            var oModel = this.getOwnerComponent().getModel();
            var that = this;
            var oTable = this.getView().byId("id_table");
            var oEntries = oTable.getSelectedItems();
            var oPayloadEntries = {};
            oPayloadEntries.to_MaterialDocumentItem = []
            var Material = {};
            for( var i = 0; i<oEntries.length; i++ ){
                //test
                var oCells = oEntries[i].getCells();
                oPayloadEntries.PostingDate = oCells[1].getText();
                oPayloadEntries.GoodsMovementCode = oCells[2].getText();
                Material.Material = oCells[3].getText();
                Material.Plant = oCells[4].getText();
                Material.StorageLocation = oCells[5].getText();
                Material.GoodsMovementType = oCells[6].getText();
                Material.PurchaseOrder = oCells[7].getText();
                Material.PurchaseOrderItem = oCells[8].getText();
                Material.GoodsMovementRefDocType = oCells[9].getText();
                Material.EntryUnit = oCells[10].getText();
                Material.QuantityInEntryUnit = oCells[11].getText();
                oPayloadEntries.to_MaterialDocumentItem.push(Material);
                console.log(oPayloadEntries);
                console.log(JSON.stringify(oPayloadEntries));
            }
            if (oPayloadEntries){
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

                try {
                    let response = await that._triggerMaterialDocument(oModel,oPayloadEntries);
                    if(response){
                        try {
                            if (response.statusCode === '201'){
                                var MessageText = 'Material Document No-' + SucessData.MaterialDocument  + '  ' +  'created successfully' ;
                                MessageBox.success(MessageText);
                            }
                        } catch (error) {
                            MessageBox.error('Error while creating Material Document');
                        }
                        oBusyDialogAdd.close();
                    }else{
                        MessageBox.error('Error while creating Material Document');
                    }
                } catch (error) {
                    var errorLog = JSON.parse(error.responseText);
                    var Message = errorLog.error.message.value;
                    MessageBox.error(Message);
                    oBusyDialogAdd.close();
                }
                
                
            }
        },
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