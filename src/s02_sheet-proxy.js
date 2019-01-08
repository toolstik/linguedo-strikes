function SheetProxy(name) {
    this.sheet = spreadSheet.getSheetByName(name);
    this.isEmpty;
    this.dataRange;
    this.values;
}

SheetProxy.prototype.getDataRange = function () {
    if (this.dataRange) return this.dataRange;

    var range = this.sheet.getDataRange();
    var numRows = range.getNumRows();

    if (numRows < 2) {
        this.isEmpty = true;
        return null;
    } else
        this.isEmpty = false;

    this.dataRange = range.offset(1, 0, numRows - 1);

    return this.dataRange;
}

SheetProxy.prototype.getValues = function () {
    if (this.values) return this.values;

    var range = this.getDataRange();

    return this.isEmpty ? [] : (this.values = range.getValues());
}

SheetProxy.prototype.getLastRow = function () {
    if (this.values) return this.values.length + 1;

    var range = this.getDataRange();

    return this.isEmpty ? 1 : range.getNumRows() + 1;
}

SheetProxy.prototype.getLastColumn = function () {
    if (this.values && this.values.length) return this.values[0].length;

    var range = this.getDataRange();

    if (this.isEmpty)
        return this.sheet.getLastColumn();

    return range.getNumColumns();
}

SheetProxy.prototype.resetRange = function () {
    this.dataRange = null;
}

SheetProxy.prototype.resetValues = function () {
    this.values = null;
}

SheetProxy.prototype.reset = function () {
    this.resetRange();
    this.resetValues();
}

SheetProxy.prototype.append = function (values) {
    console.time('SheetProxy.append');
    if (!values || !values.length) return;

    var rangeToFill = this.sheet.getRange(this.getLastRow() + 1, 1, values.length, this.getLastColumn());
    rangeToFill.setValues(values);

    if (this.values)
        this.values = this.values.concat(values);

    this.resetRange();
    console.timeEnd('SheetProxy.append');
}