function isValidDate(value) {
  return Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value);
}
//get next day without using timezone
function getNextDay(date) {
  var date = new Date(date);
  // Add one day (automatically handles month-end)
  date.setDate(date.getDate() + 1);
  // Extract year, month, and day without timezone issues
  var year = date.getFullYear();
  var month = String(date.getMonth() + 1).padStart(2, '0'); // Ensure two digits
  var day = String(date.getDate()).padStart(2, '0'); // Ensure two digits
  var nextDay = `${year}-${month}-${day}`;
  return nextDay;
}
function columnToLetter(column) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}