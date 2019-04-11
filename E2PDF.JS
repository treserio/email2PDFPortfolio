// These functions need to be placed in your privileged JS folder for Acrobat. Here's info on finding it: https://www.pdfscripting.com/public/Installing_Automation_Tools.cfm
// C:\Users\<username>\AppData\Roaming\Adobe\Acrobat\ if the folder is missing the following add them for Acrobat 11, it will change if it's DC instead "Privileged\11.0\JavaScripts".

// For a regular pdf this will import items as attachments, for a pdf portfolio it will import and provide fields to display additional data in a list format.
safeImportDataObject = app.trustedFunction(function(impDoc)
{
    app.beginPriv();
    // (name, location)
    this.importDataObject("Test.pdf", "/C/Users/tserio/Documents/Test.pdf");
    // app.alert("Import was Successful");
    app.endPriv();
});

// Set the field values for emails imported into pdf portpolios created by Adobe's pdfMaker add on for Outlook.
// tried to use objName with this.getDataObject(objName) but it never worked.
setDataObjFieldValues = app.trustedFunction(function(objName, emailFrom, emailTo, emailCC, emailSub, emailDate, numAttached)
{
    // elevate access
    app.beginPriv();
    
    // get most recent dataObject
    // var dataObj = this.getDataObject(objName); still unsure why this gives errors
    var dataObj = this.dataObjects[0];

    // set field values
    dataObj.setFieldValue("2", emailFrom);
    dataObj.setFieldValue("3", emailSub);
    dataObj.setFieldValue("4", emailDate);
    dataObj.setFieldValue("6", numAttached);
    dataObj.setFieldValue("8", emailTo);
    dataObj.setFieldValue("9", emailCC);
    
    // end access
    app.endPriv();
});