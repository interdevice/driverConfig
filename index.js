const { Builder, Key, By, until } = require('selenium-webdriver')
const fs = require('fs')
const lerXLS = require('xlsx')
const firefox = require('selenium-webdriver/firefox')
const proxy = require('selenium-webdriver/proxy')
const { time } = require('console')
const options = new firefox.Options()

options.setProfile('./SeleniumProfile')
//const proxyServer='	140.227.62.35:58888'

options.setPreference("browser.download.dir", "C:\\Users\\luizh\\www\\selenium\\driverConfig")
options.setPreference("browser.download.folderList", 2)
options.setPreference("browser.helperApps.neverAsk.saveToDisk","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

const driver = new Builder()
    .forBrowser('firefox')
    .setFirefoxOptions(options)
    /*.setProxy(proxy.manual({
        http: proxyServer,
        https: proxyServer
    }))*/
    .build()

async function download(){
    await driver.get('http://rpachallenge.com/')
    download = await driver.findElement(By.xpath("//a[@class=' col s12 m12 l12 btn waves-effect waves-light uiColorPrimary center']"))
    download.click()
    var css = 'col s12 m12 l12 btn waves-effect waves-light uiColorPrimary center'
    await driver.wait(until.elementLocated({css: css}),10000)
    
    busca()
    
}

async function busca() {
    var workbook = lerXLS.readFile('challenge.xlsx')
    var sheet_name_list = workbook.SheetNames
    let data = []
    var xlData = lerXLS.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]], {header:"A"})
    
    xlData.forEach((res) => {
        data.push(res)
     })
    
    
    await driver.get('http://rpachallenge.com/')

    start = await driver.findElement(By.xpath("//button[@class='waves-effect col s12 m12 l12 btn-large uiColorButton']"))
    await driver.wait(until.elementLocated(By.xpath("//a[@class=' col s12 m12 l12 btn waves-effect waves-light uiColorPrimary center']")),10000)
    

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
        
        firstName.sendKeys(data[i].A)
        lastName.sendKeys(data[i].B)
        company.sendKeys(data[i].C)
        role.sendKeys(data[i].D)
        address.sendKeys(data[i].E)
        email.sendKeys(data[i].F)
        phone.sendKeys(data[i].G)

        driver.findElement(By.xpath("//input[@class='btn uiColorButton']")).click()
     }
    
     let xpath = "//div[@class='congratulations col s8 m8 l8']"
     await driver.wait(until.elementLocated({xpath: xpath}), 10000)
     
     const encodedString = await driver.takeScreenshot()
     fs.writeFileSync('./result.png', encodedString, 'base64')

}

download()