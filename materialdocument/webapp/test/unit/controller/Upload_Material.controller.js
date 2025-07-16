/*global QUnit*/

sap.ui.define([
	"materialdocument/controller/Upload_Material.controller"
], function (Controller) {
	"use strict";

	QUnit.module("Upload_Material Controller");

	QUnit.test("I should test the Upload_Material controller", function (assert) {
		var oAppController = new Controller();
		oAppController.onInit();
		assert.ok(oAppController);
	});

});
