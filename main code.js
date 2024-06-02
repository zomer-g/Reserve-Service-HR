/*
 * This script extracts all unique names from the range C4:AT100 in the "Gantt chart" sheet,
 * sorts them by the categories in column A, and logs the list of names.
 * For each unique name, it creates a vocabulary containing the name, ID, and dates,
 * organized by the type of activity. The dates are displayed in "DD/MM/YY" format.
 * Dates are formatted as [date: DD/MM/YY, activity: XYZ] and sorted in chronological order.
 * The script writes the list of names in column A, the list of IDs in column B,
 * and the list of dates in column C in the "Summary" sheet, ignoring the activity.
 * Continuous dates are grouped and displayed as [start date]-[end date].
 * Additionally, this script extracts all unique dates and writes them in chronological order
 * to the "Kishur" sheet in column A. For each date, it writes all the names that appear on this date
 * and didn't appear on the previous date in column B, and all the names that appear on this date
 * and didn't appear on the next date in column C. Each name is followed by its ID in parentheses.
 * Finally, this script creates "Report1" based on the categories "קו" and "מפלג", copying columns A and B
 * from the "Gantt chart" for these categories, splitting column B into three columns based on spaces,
 * adding an ID column, and populating column C with the names and IDs corresponding to today's date.
 * This script also writes to the "Counter" sheet, listing names and IDs sorted by name in columns A and B,
 * and adding a column for each category with the number of times each name appears in that category.
 * A final column "All" counts the total number of appearances in all categories for each name.
 */

function extractAndLogUniqueNamesAndDetails() {
  try {
    console.log('Script started');

    // Define the spreadsheet and sheets
    var spreadsheetId = '[SPREADSHEET_ID]';
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var ganttSheet = spreadsheet.getSheetByName('Gantt chart');
    var idSheet = spreadsheet.getSheetByName('ID');
    var summarySheet = spreadsheet.getSheetByName('Summary');
    var kishurSheet = spreadsheet.getSheetByName('Kishur');
    var reportSheet = spreadsheet.getSheetByName('Report1');
    var counterSheet = spreadsheet.getSheetByName('Counter');
    
    console.log('Fetching data from the Gantt chart sheet');
    
    // Fetch the data from the Gantt chart sheet
    var ganttData = ganttSheet.getRange('A3:AT100').getValues();
    var namesRange = ganttSheet.getRange('C4:AT100').getValues();
    var dates = ganttData[0].slice(2); // Extract the dates from the header row
    var categories = ganttData.slice(1).map(function(row) { return row[0]; }); // Extract the types of activity

    console.log('Fetching data from the ID sheet');
    
    // Fetch the data from the ID sheet
    var idData = idSheet.getRange('A:D').getValues();

    // Create a map for IDs
    var idMap = {};
    idData.forEach(function(row) {
      var name = row[0];
      var id = row[3];
      if (name && id) {
        idMap[name] = id;
      }
    });

    console.log('Extracting categories and names');
    
    // Extract categories and names
    var names = namesRange.flat().filter(function(name) { return name !== ''; });
    
    // Get unique names
    var uniqueNames = [...new Set(names)];
    
    console.log('Unique Names:', uniqueNames);
    
    // Create a map for sorting names by categories
    var nameCategoryMap = {};
    uniqueNames.forEach(function(name) {
      for (var i = 0; i < namesRange.length; i++) {
        for (var j = 0; j < namesRange[i].length; j++) {
          if (namesRange[i][j] === name) {
            var category = ganttData[i + 1][0];
            if (!nameCategoryMap[category]) {
              nameCategoryMap[category] = [];
            }
            nameCategoryMap[category].push(name);
          }
        }
      }
    });

    // Sort categories
    var sortedCategories = Object.keys(nameCategoryMap).sort();
    
    console.log('Sorted Categories:', sortedCategories);

    // Create and log a vocabulary for each name
    var nameDetails = [];
    var processedNames = new Set();
    uniqueNames.forEach(function(name) {
      var vocabulary = {
        name: name,
        id: idMap[name] || 'ID not found',
        dates: []
      };

      namesRange.forEach(function(row, rowIndex) {
        row.forEach(function(cell, colIndex) {
          if (cell === name) {
            var date = new Date(dates[colIndex]);
            var formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yy');
            vocabulary.dates.push({ date: date, formatted: formattedDate });
          }
        });
      });

      // Sort dates in chronological order
      vocabulary.dates.sort(function(a, b) {
        return a.date - b.date;
      });

      // Group continuous dates
      var groupedDates = [];
      var startDate = null;
      var endDate = null;

      vocabulary.dates.forEach(function(item, index) {
        if (startDate === null) {
          startDate = item.date;
          endDate = item.date;
        } else if ((item.date - endDate) / (1000 * 60 * 60 * 24) === 1) {
          endDate = item.date;
        } else {
          if (startDate.getTime() === endDate.getTime()) {
            groupedDates.push(Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'dd/MM/yy'));
          } else {
            groupedDates.push(Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'dd/MM/yy') + '-' + Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'dd/MM/yy'));
          }
          startDate = item.date;
          endDate = item.date;
        }
      });

      if (startDate !== null) {
        if (startDate.getTime() === endDate.getTime()) {
          groupedDates.push(Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'dd/MM/yy'));
        } else {
          groupedDates.push(Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'dd/MM/yy') + '-' + Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'dd/MM/yy'));
        }
      }

      vocabulary.dates = groupedDates;

      nameDetails.push(vocabulary);
      processedNames.add(name);
    });

    console.log('Sorting the final list of names alphabetically');

    // Sort nameDetails by the name
    nameDetails.sort(function(a, b) {
      if (a.name < b.name) return -1;
      if (a.name > b.name) return 1;
      return 0;
    });

    console.log('Writing to the Summary sheet');

    // Clear the Summary sheet
    summarySheet.clear();

    // Write headers to the Summary sheet
    summarySheet.getRange('A1').setValue('Name');
    summarySheet.getRange('B1').setValue('ID');
    summarySheet.getRange('C1').setValue('Dates');

    // Write the list of names, IDs, and dates to the Summary sheet
    nameDetails.forEach(function(detail, index) {
      summarySheet.getRange(index + 2, 1).setValue(detail.name);
      summarySheet.getRange(index + 2, 2).setValue(detail.id);
      summarySheet.getRange(index + 2, 3).setValue(detail.dates.join(', '));
    });

    console.log('Logging the list of names with all dates collected for each name');
    nameDetails.forEach(function(detail) {
      console.log('Name:', detail.name, ', ID:', detail.id, ', Dates:', detail.dates);
    });

    console.log('Collecting all unique dates');

    // Collect all unique dates
    var allDates = [];
    namesRange.forEach(function(row, rowIndex) {
      row.forEach(function(cell, colIndex) {
        if (cell) {
          var date = new Date(dates[colIndex]);
          var formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yy');
          if (!allDates.includes(formattedDate)) {
            allDates.push(formattedDate);
          }
        }
      });
    });

    // Sort all dates in chronological order
    allDates.sort(function(a, b) {
      var dateA = new Date(a.split('/').reverse().join('/'));
      var dateB = new Date(b.split('/').reverse().join('/'));
      return dateA - dateB;
    });

    console.log('Writing all unique dates to the Kishur sheet');

    // Clear the Kishur sheet
    kishurSheet.clear();

    // Write headers to the Kishur sheet
    kishurSheet.getRange('A1').setValue('Dates');
    kishurSheet.getRange('B1').setValue('Start Date');
    kishurSheet.getRange('C1').setValue('End Date');

    // Write all unique dates to the Kishur sheet
    allDates.forEach(function(date, index) {
      kishurSheet.getRange(index + 2, 1).setValue(date);
    });

    // Create a map of names for each date
    var dateNamesMap = {};
    dates.forEach(function(date, colIndex) {
      var formattedDate = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'dd/MM/yy');
      if (!dateNamesMap[formattedDate]) {
        dateNamesMap[formattedDate] = [];
      }
      namesRange.forEach(function(row) {
        if (row[colIndex]) {
          var name = row[colIndex];
          var nameWithId = name + ' (' + (idMap[name] || 'ID not found') + ')';
          dateNamesMap[formattedDate].push(nameWithId);
        }
      });
    });

    // Write names to the Kishur sheet based on appearance criteria
    allDates.forEach(function(date, index) {
      var namesOnDate = dateNamesMap[date] || [];
      var namesOnPreviousDate = index > 0 ? dateNamesMap[allDates[index - 1]] || [] : [];
      var namesOnNextDate = index < allDates.length - 1 ? dateNamesMap[allDates[index + 1]] || [] : [];

      var startNames = namesOnDate.filter(function(name) {
        return !namesOnPreviousDate.includes(name);
      });
      var endNames = namesOnDate.filter(function(name) {
        return !namesOnNextDate.includes(name);
      });

      kishurSheet.getRange(index + 2, 2).setValue(startNames.join(', '));
      kishurSheet.getRange(index + 2, 3).setValue(endNames.join(', '));
    });

    console.log('Generating Report1');

    // Clear the Report1 sheet
    reportSheet.clear();

    // Write headers to the Report1 sheet
    var today = new Date();
    var formattedToday = Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd/MM/yy');
    reportSheet.getRange('A1').setValue(formattedToday);
    reportSheet.getRange('B1').setValue('Sub-category 1');
    reportSheet.getRange('C1').setValue('Sub-category 2');
    reportSheet.getRange('D1').setValue('Sub-category 3');
    reportSheet.getRange('E1').setValue('Name');
    reportSheet.getRange('F1').setValue('ID');

    // Write the categories and names for today to the Report1 sheet
    var reportRowIndex = 2;
    ganttData.slice(1).forEach(function(row, rowIndex) {
      var category = row[0];
      var subCategory = row[1];
      if (category === 'קו' || category === 'מפלג') {
        reportSheet.getRange(reportRowIndex, 1).setValue(category);

        var subCategories = subCategory.split(' ');
        for (var i = 0; i < subCategories.length; i++) {
          reportSheet.getRange(reportRowIndex, i + 2).setValue(subCategories[i]);
        }

        var dateIndex = dates.findIndex(function(date) {
          return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'dd/MM/yy') === formattedToday;
        });

        if (dateIndex !== -1) {
          var names = namesRange[rowIndex][dateIndex];
          if (names) {
            var nameList = names.split(',').map(function(name) {
              name = name.trim();
              return {
                name: name,
                id: idMap[name] || 'ID not found'
              };
            });

            nameList.forEach(function(item) {
              reportSheet.getRange(reportRowIndex, 5).setValue(item.name);
              reportSheet.getRange(reportRowIndex, 6).setValue(item.id);
              reportRowIndex++;
            });
          }
        }
      }
    });

    console.log('Writing to the Counter sheet');

    // Clear the Counter sheet
    counterSheet.clear();

    // Write headers to the Counter sheet
    counterSheet.getRange('A1').setValue('Name');
    counterSheet.getRange('B1').setValue('ID');

    // Write the list of names and IDs to the Counter sheet
    nameDetails.forEach(function(detail, index) {
      counterSheet.getRange(index + 2, 1).setValue(detail.name);
      counterSheet.getRange(index + 2, 2).setValue(detail.id);
    });

    // Write category headers
    var categoryHeaders = sortedCategories.map(function(category, index) {
      counterSheet.getRange(1, index + 3).setValue(category);
      return category;
    });

    counterSheet.getRange(1, categoryHeaders.length + 3).setValue('All');

    // Count the number of times each name appears in each category and in total
    nameDetails.forEach(function(detail, index) {
      var totalDays = 0;
      categoryHeaders.forEach(function(category, categoryIndex) {
        var categoryDays = nameCategoryMap[category].filter(function(name) {
          return name === detail.name;
        }).length;
        counterSheet.getRange(index + 2, categoryIndex + 3).setValue(categoryDays);
        totalDays += categoryDays;
      });
      counterSheet.getRange(index + 2, categoryHeaders.length + 3).setValue(totalDays);
    });

    console.log('Script completed successfully');
  } catch (error) {
    console.error('An error occurred:', error.message);
  }
}
