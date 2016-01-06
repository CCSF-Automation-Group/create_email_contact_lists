
function CheckAddresses(addressarr, itemtoadd) {
  for (var i = 0; i < addressarr.length; i++) {
    if (itemtoadd === addressarr[i])
      return false;
  }
  return true;
}
function ParseAddressforName(address) {
  var addresscomps = address.split("<");
  if (addresscomps.length <= 1 )
    return "X";
  return addresscomps[0];
}
function ParseAddressforAddr(address) {
  var addresscomps = address.split("<");
  if (addresscomps.length <=1)
    return addresscomps[0];
  addresscomps = addresscomps[1].split(">");
  return addresscomps[0];
}

function SaveContacts() {

  var gmailLabels  = "Address_Contact";
  var driveFolder  = "CCSFAUTO";
  var OutputName = "AddressSheet"

  var threads = GmailApp.search("in:" + gmailLabels, 0, 5);
  var addresses = [];
  if (threads.length > 0) {
    /* Gmail Label that contains the queue */
    var label = GmailApp.getUserLabelByName(gmailLabels) ?
        GmailApp.getUserLabelByName(gmailLabels) : GmailApp.createLabel(driveFolder);

    for (var t=0; t<threads.length; t++) {
      // Remove label to avoid duplicate actions!
      threads[t].removeLabel(label);
      var msgs = threads[t].getMessages();
      for (var i = 0; i < msgs.length; i++) {
        // getTo()
        if (CheckAddresses(addresses, msgs[i].getTo()))
            addresses.push(msgs[i].getTo());
        // getFrom()
        if (CheckAddresses(addresses, msgs[i].getFrom()))
            addresses.push(msgs[i].getFrom());
        // getCc()
        if (CheckAddresses(addresses, msgs[i].getCc()))
            addresses.push(msgs[i].getCc());
        // getBcc()
        if (CheckAddresses(addresses, msgs[i].getBcc()))
            addresses.push(msgs[i].getBcc());
        }
      }
    var file = DriveApp.getFilesByName(OutputName);
    if (file.hasNext()) {
      var Spreadsheet = SpreadsheetApp.open(file.next());
    } else {
      var Spreadsheet = SpreadsheetApp.create(OutputName);
    }
    var d = new Date();
    Spreadsheet.appendRow(["","" , "Added on: " + d.toDateString()]);
    for (var i = 0; i < addresses.length; i++) {
      Spreadsheet.appendRow([ParseAddressforName(addresses[i]), ParseAddressforAddr(addresses[i])])
                           .autoResizeColumn(1).autoResizeColumn(2);
      }
   }
}
