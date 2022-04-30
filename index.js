module.exports = function transformData(worksheetData) {
  let cleanColumn = (x) => {
    // Intents to remove ATTR(), SUM(), AGGR(), AVG()
    var regExp = /[ATTR|SUM|AGGR|AVG]\(([^)]+)\)/;
    try {
      return regExp.exec(x)[1];
    } catch (error) {
      return x;
    }
  };

  let cleanMeasureColumn = (x) => {
    var regExp = /:(.*?):/;
    try {
      return regExp.exec(x)[1];
    } catch (error) {
      return x;
    }
  };

  const FIELDNAME = "_fieldName";
  const COLUMNNAME = "_columns";
  const DATANAME = "_data";
  const FORMATTEDVALUE = "_formattedValue";
  const VALUENAME = "_value";
  const NULLVALUE = "%null%";

  let rows = [];

  const measureNamesIndex = worksheetData[COLUMNNAME].map(
    (x) => x[FIELDNAME]
  ).indexOf("Measure Names");
  const hasMeasureNames = measureNamesIndex > -1;

  if (hasMeasureNames) {
    let originDimensions = worksheetData[COLUMNNAME].length - 2;
    let originColumns = worksheetData[COLUMNNAME].slice(
      0,
      originDimensions
    ).map((x) => x[FIELDNAME]);
    let measureColumns = Array.from(
      new Set(
        worksheetData[DATANAME].map(
          (x) =>
            x[originDimensions][FORMATTEDVALUE] ||
            cleanMeasureColumn(x[originDimensions][VALUENAME])
        )
      )
    ).reverse(); // the order is by default reversed, so this just goes back to normal

    let columns = [...originColumns, ...measureColumns];

    // Let's find the amount of rows in a group (the inner group)
    let rowRepeatCount = worksheetData[DATANAME].findIndex(
      (x) =>
        x[originDimensions][VALUENAME] !==
        worksheetData[DATANAME][0][originDimensions][VALUENAME]
    );

    columns = columns.map((x) => cleanColumn(x)); // Clean the columns from SUM(), ATTR() and so on....
    rows.columns = columns;

    for (
      let outerIndex = 0;
      outerIndex < worksheetData[DATANAME].length;
      outerIndex += rowRepeatCount * measureColumns.length
    ) {
      for (let innerIndex = 0; innerIndex < rowRepeatCount; innerIndex++) {
        let i = outerIndex + innerIndex;
        let rowData = [
          ...worksheetData[DATANAME][i]
            .slice(0, originDimensions)
            .map((x) => x[VALUENAME]),
          ...[...Array(measureColumns.length).keys()]
            .reverse() // the order is by default reversed, so this just goes back to normal
            .map(
              (ind) =>
                worksheetData[DATANAME][i + ind * rowRepeatCount][
                  originDimensions + 1
                ][VALUENAME]
            ),
        ];
        // Equals to Pythons dict(zip(rowData, columns))
        // https://riptutorial.com/javascript/example/8628/merge-two-array-as-key-value-pair
        let row = rowData.reduce(function (result, field, index) {
          result[columns[index]] = field === NULLVALUE ? null : field;
          return result;
        }, {});

        rows.push(row);
      }
    }
  } else {
    let columns = worksheetData[COLUMNNAME].map((x) => x[FIELDNAME]);
    columns = columns.map((x) => cleanColumn(x)); // Clean the columns from SUM(), ATTR() and so on....
    rows.columns = columns;

    for (var i = 0; i < worksheetData[DATANAME].length; i++) {
      var row = {};
      for (var col = 0; col < columns.length; col++) {
        let value = worksheetData[DATANAME][i][col].value;
        row[columns[col]] = value === NULLVALUE ? null : value;
      }
      rows.push(row);
    }
  }

  return rows;
};
