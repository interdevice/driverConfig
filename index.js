const { Builder, Key, By, until } = require('selenium-webdriver')
const fs = require('fs')
const lerXLS = require('xlsx')
const firefox = require('selenium-webdriver/firefox')
const proxy = require('selenium-webdriver/proxy')
const options = new firefox.Options()

options.setProfile('./SeleniumProfile')
const proxyServer='	140.227.62.35:58888'

options.setPreference("browser.download.dir", "C:\\Users\luizh\www\selenium")
options.setPreference("browser.download.folderList", 2)
options.setPreference("browser.helperApps.neverAsk.saveToDisk","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

const driver = new Builder()
    .forBrowser('firefox')
    .setFirefoxOptions(options)
    /* .setProxy(proxy.manual({
        http: proxyServer,
        https: proxyServer
    })) */
    .build()

async function busca() {
    var workbook = lerXLS.readFile('challenge.xlsx');
    var sheet_name_list = workbook.SheetNames;
    let data = []
    var xlData = lerXLS.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
    xlData.forEach((res) => {
        data.push(res)
     })
    
    await driver.get('http://rpachallenge.com/')
    start = await driver.findElement(By.xpath("//button[@class='waves-effect col s12 m12 l12 btn-large uiColorButton']"))
    await driver.wait(until.elementLocated(By.xpath("//a[@class=' col s12 m12 l12 btn waves-effect waves-light uiColorPrimary center']")),10000)
    download = await driver.findElement(By.xpath("//a[@class=' col s12 m12 l12 btn waves-effect waves-light uiColorPrimary center']"))
    download.click()

    start.click()
    
    for(let i = 0; i < data.length; i++) {
        await driver.wait(until.elementLocated(By.xpath("//input[@class='btn uiColorButton']")),15000)
        firstName = driver.findElement(By.xpath("//input[@ng-reflect-name='labelFirstName']"))
        lastName = driver.findElement(By.xpath("//input[@ng-reflect-name='labelLastName']"))
        company = driver.findElement(By.xpath("//input[@ng-reflect-name='labelCompanyName']"))
        role = driver.findElement(By.xpath("//input[@ng-reflect-name='labelRole']"))
        address = driver.findElement(By.xpath("//input[@ng-reflect-name='labelAddress']"))
        email = driver.findElement(By.xpath("//input[@ng-reflect-name='labelEmail']"))
        phone = driver.findElement(By.xpath("//input[@ng-reflect-name='labelPhone']"))
        
        firstName.sendKeys(data[i].First_Name)
        lastName.sendKeys(data[i].Last_Name)
        company.sendKeys(data[i].Company_Name)
        role.sendKeys(data[i].Role_in_Company)
        address.sendKeys(data[i].Address)
        email.sendKeys(data[i].Email)
        phone.sendKeys(data[i].Phone_Number)

        driver.findElement(By.xpath("//input[@class='btn uiColorButton']")).click()
     }
    
}

busca()