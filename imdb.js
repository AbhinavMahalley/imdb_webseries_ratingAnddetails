const fs = require("fs");
const xlsx = require("xlsx");
const request = require("request");
const puppeteer = require("puppeteer");
const cheerio = require("cheerio");
const path=require("path");
const pdfkit=require("pdfkit");
const siteURL = "https://www.imdb.com/";

let page, browser;

(async function fn() {
  try {
    //launch chromium browser
    let browserStartPromise = await puppeteer.launch({
      // visible
      headless: false,
      // type 1sec // slowMo: 1000,
      defaultViewport: null,
      // browser setting
      args: ["--start-maximized", "--disable-notifications"],
    });

    let browserObj = browserStartPromise;
    console.log("Browser opened");
    browser = browserObj;
    // new tab
    page = await browserObj.newPage();
    await page.goto(siteURL);

    //excel file
    let webSeries = excelReader("./webseries.xlsx", "Sheet1");
    for (let key in webSeries) {
      webSeriesName = webSeries[key].Web_Series;
      console.log(webSeriesName);

      await page.type("input[placeholder='Search IMDb']", webSeriesName);
      console.log("search");

      await page.click("button[id='suggestion-search-button']");
      await waitAndClick("td[class='result_text'] a", page);
      const url = await page.url({ delay: 1000 });
      console.log("Page URL : " + url);//url of current page

      //Auto Scroll Web-page
      await autoScroll(page);

      // screenshot of current web page
      await page.screenshot({
        path: webSeriesName+'_screenshot.png',
        fullPage: true
      });

      // request call back function
      request(url, cb);

      await page.waitForTimeout(5000);
      await page.goBack();//Back to previous webpage
    }

    await browser.close();//close browser

  } catch (err) {
    console.log(err);
  }
})();


async function cb(error, response, html) {
  if (error) {
    console.log(error); // Print the error if one occurred
  } else if (response.statusCode == 404) {
    console.log("Page Not Found");//page not found
  } else {
    //function to extract data from current page 
    dataExtracter(html);
    console.log( "...........................................................................................");
  }
}

function dataExtracter(html) {
  let searchTool = cheerio.load(html);

  let seriesName = searchTool( "h1[data-testid='hero-title-block__title']").text();
  let rating = searchTool(".AggregateRatingButton__Rating-sc-1ll29m0-2.bmbYRW").text().split("/");
  let userReviews = searchTool('span[class="less-than-three-Elements"]').text().split("U");
  let totalEpisods = searchTool('div[data-testid="episodes-header"] span.ipc-title__subtext').text();
  let storyLine = searchTool("section[data-testid='Storyline'] .ipc-html-content.ipc-html-content--base").text();
  let imageUrl = searchTool(".ipc-media.ipc-media--poster.ipc-image-media-ratio--poster.ipc-media--baseAlt.ipc-media--poster-l.ipc-poster__poster-image.ipc-media__img  .ipc-image").attr('src');
 
  let  details = searchTool(
    'section[data-testid="Details"] .ipc-metadata-list.ipc-metadata-list--dividers-all.ipc-metadata-list--base .ipc-metadata-list-item__list-content-item.ipc-metadata-list-item__list-content-item--link'
  );
  let releaseDate=searchTool(details[0]).text();
  let Country=searchTool(details[1]).text();
  let officialSites=searchTool(details[2]).text();
  let Language=searchTool(details[3]).text();

  console.log("\n"+seriesName +" || Rating- " + rating[0] +"/10  || " + userReviews[0] +" User Reviews || Total Episodes " + totalEpisods+"\n");
  console.log("Storyline:-  " + storyLine+"\n");
  console.table([{"releaseDate":releaseDate,"Country":Country,"officialSite":officialSites,"Language":Language}]);
  
  let SeriesDetails="Web Series:- "+seriesName +" || Rating:- " + rating[0] +"/10   || " + userReviews[0] +" User Reviews || Total Episodes:- " + totalEpisods;
  let Storydetails="Storyline:-  " + storyLine;
  let webSeriesDetails="Release Date:- "+releaseDate+" || Country:- "+Country+" || OfficialSite:- "+officialSites+" || Language:- "+Language;

//Create Directory
let folderpath=path.join(__dirname,"IMDB_WEB_SERIES_RATINGS_&_DETAILS");
dirCreater(folderpath);
//pdf file path
let filePath=path.join(folderpath,seriesName+".pdf");
//screenshot image path
let screenshotPath=path.join(__dirname,seriesName+"_screenshot.png");
//url of current webpage
let url =  page.url();

//Excel Creation
processWebseries(seriesName,rating[0],userReviews[0],totalEpisods,storyLine,releaseDate,Country,officialSites,Language,url);

// PDF Creation
   let pdfDoc=new pdfkit;
   pdfDoc.pipe(fs.createWriteStream(filePath));
   pdfDoc.font('Helvetica-Bold').text(SeriesDetails,7,15);
   pdfDoc.fontSize(10).font('Helvetica').fillColor("#F74621").text(Storydetails,5,30);
   pdfDoc.font('Helvetica').fillColor("#5C5A59").text(webSeriesDetails);
   pdfDoc.fontSize(8).fillColor("#031CFF ").text('IMDB Link',{link:url});
   pdfDoc.image(screenshotPath, { width: 350, height: 890});
   pdfDoc.end();
 
}



//wait for selector and click on it.
async function waitAndClick(selector, cPage) {
  try {
    await cPage.waitForSelector(selector, { visible: true });
    await cPage.click(selector);
  } catch (err) {
    return new Error(err);
  }
}

//process web-series data into excel file
function processWebseries(seriesName,rating,userReviews,totalEpisods,storyLine,releaseDate,Country,officialSites,Language,url){
  let dirPath=path.join(__dirname,"IMDB_WEB_SERIES_RATINGS_&_DETAILS");
  dirCreater(dirPath);
  let filePath=path.join(dirPath,"Web_Series_Ratings_&_Details.xlsx");
  let content=excelReader(filePath,"Sheet1");
  let webseriesObj={
    "SeriesName":seriesName,
    "Rating":rating+"/10",
    "UserReviews":userReviews,
    "Episods":totalEpisods,
    "StoryLine":storyLine,
    "ReleaseDate":releaseDate,
    "Country":Country,
    "OfficialSite":officialSites,
    "Language":Language,
    "Link":url
  }
  content.push(webseriesObj);
  excelWriter(filePath,content,"Sheet1");
  }
  
  // function to scroll the whole web-page
  async function autoScroll(page){
    await page.evaluate(async () => {
        await new Promise((resolve, reject) => {
            var totalHeight = 0;
            var distance = 100;
            var timer = setInterval(() => {
                var scrollHeight = document.body.scrollHeight;
                window.scrollBy(0, distance);
                totalHeight += distance;
  
                if(totalHeight >= scrollHeight){
                    clearInterval(timer);
                    resolve();
                }
            }, 100);
        });
    });
  }

// excel file reader
function excelReader(filePath, sheetName) {
  if (fs.existsSync(filePath) == false) {
    return [];
  }
  let wb = xlsx.readFile(filePath);
  let excelData = wb.Sheets[sheetName];
  let ans = xlsx.utils.sheet_to_json(excelData);
  return ans;
}

// excel file writer
function excelWriter(filePath,json,sheetName){
  let newWB=xlsx.utils.book_new();
  let newWS=xlsx.utils.json_to_sheet(json);
  xlsx.utils.book_append_sheet(newWB,newWS,sheetName);
  xlsx.writeFile(newWB,filePath);
}

// create directory
function dirCreater(folderpath){
    if(fs.existsSync(folderpath)==false){
        fs.mkdirSync(folderpath);
    }
}