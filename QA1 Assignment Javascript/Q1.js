const { Builder, By, Key, until } = require('selenium-webdriver');
const ExcelJS = require('exceljs');
const { DateTime } = require('luxon');
const { setTimeout } = require('timers/promises');
//read excel file
(async function main() {
  const excel = 'D1.xlsx';
  const sheetNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']; 
  const today = DateTime.local();
  const searchDayIndex = today.weekday; 
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(excel);
  const sheet = workbook.getWorksheet(sheetNames[searchDayIndex]);

  const driver = await new Builder().forBrowser('chrome').build();

  try {
    await driver.get('https://www.google.com');
    const searchBox = await driver.wait(until.elementLocated(By.name('q')), 3000);

    for (let rowIdx = 3; rowIdx <= 13; rowIdx++) {
      const searchQuery = sheet.getCell(rowIdx, 3).value;

      if (searchQuery) {
        await searchBox.clear();
        await searchBox.sendKeys(searchQuery);
        //wait for search
        await setTimeout(3000); 

        //Inspect path
        try {
          const suggestionsBox = await driver.wait(
            until.elementLocated(By.xpath("/html/body/div[1]/div[3]/form/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div/ul")),
            5000
          );
            
          const suggestionsElements = await suggestionsBox.findElements(By.xpath('./li[@class="sbct"]'));
          const suggestions = await Promise.all(suggestionsElements.map(el => el.getText()));
            // Find long and short option 
          const longestSuggestion = suggestions.reduce((longest, current) => current.length > longest.length ? current : longest, '');
          const shortestSuggestion = suggestions.reduce((shortest, current) => current.length < shortest.length ? current : shortest, suggestions[0]);

            //store data excel fiel
          sheet.getCell(rowIdx, 4).value = longestSuggestion;
          sheet.getCell(rowIdx, 5).value = shortestSuggestion;
        } catch (error) {
          console.log('Suggestions not found for:', searchQuery);
        }
      }
    }

    await workbook.xlsx.writeFile(excel);
  } finally {
    await driver.quit();
  }
})();
