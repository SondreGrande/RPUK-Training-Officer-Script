function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('ODP Creation');
  
  menu.addItem('Generate ODP','startForm');
  menu.addSeparator();

  menu.addItem('User Guide','userguide');
  menu.addToUi();
}
