const { Builder, By, Key } = require('selenium-webdriver');
const xlsx = require('xlsx');
const path = require('path');

const file = path.resolve('4BeatsQ1.xlsx');
const wb = xlsx.readFile(file);

const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const today = days[new Date().getDay()];
const ws = wb.Sheets[today];

const driver = new Builder().forBrowser('chrome').build();

(async function run() {
    const range = xlsx.utils.decode_range(ws['!ref']);
    
    for (let r = range.s.r + 1; r <= range.e.r; r++) {
        const cell = ws[xlsx.utils.encode_cell({ r: r, c: 0 })];
        const keyword = cell ? cell.v : '';

        if (!keyword) continue;

        await driver.get('http://www.google.com');
        const search = await driver.findElement(By.name('q'));
        await search.sendKeys(keyword, Key.ARROW_DOWN);
        await driver.sleep(1000);

        const suggestions = await driver.findElements(By.css('ul[role="listbox"] li span'));
        const texts = await Promise.all(suggestions.map(s => s.getText()));

        const longOpt = texts.length > 0 ? texts.reduce((a, b) => a.length > b.length ? a : b) : 'No suggestions found';
        const shortOpt = texts.length > 0 ? texts.reduce((a, b) => a.length < b.length ? a : b) : 'No suggestions found';

        ws[xlsx.utils.encode_cell({ r: r, c: 1 })] = { v: longOpt };
        ws[xlsx.utils.encode_cell({ r: r, c: 2 })] = { v: shortOpt };
    }

    xlsx.writeFile(wb, file);
    await driver.quit();
})();
