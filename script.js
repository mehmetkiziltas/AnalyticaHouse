const { GoogleSpreadSheet } = require('google-spreadsheet');
const xlsx = require('xlsx');
const axios = require('axios');
const { promisify } = require('util');
const cheerio = require('cheerio');
const request_promise = require('request-promise');
const json2csv = require('json2csv').Parser;
const fs = require('fs');
const { google } = require('googleapis');
const creds = require('./client_secret.json');


const web_site_url = "https://www.markastok.com";
const spreadsheetId = "1FFEl_xGOf_XhKYuYwmZXz_olaKiLG2MkxEdBghj7yPI";


const URLs = [];


const googleSheetAuth = async () => {
    const client = new google.auth.JWT(
        creds.client_email,
        null,
        creds.private_key,
        ['https://www.googleapis.com/auth/spreadsheets']
    );
    client.authorize(async (err, tokens) => {
        if (err) {
            if (fs.existsSync('error.json')) {
                const data = fs.readFileSync('error.json');
                const json = JSON.parse(data);
                json.push(errorBody);
                fs.writeFileSync('error.json', JSON.stringify(json));
            } else {
                fs.writeFileSync('error.json', JSON.stringify(errorBody));
            }
            return;
        } else {
            await writeToGoogleSheet(client);
            await writeErrorToGoogleSheet(client);
        }

    });
};

const writeErrorToGoogleSheet = async (cl) => {
    const erorJson = fs.readFileSync('error.json');
    const errorData = JSON.parse(erorJson);
    const gsapiError = google.sheets({ version: 'v4', auth: cl });
    errorData.forEach(async (error) => {
        const errorBody = {
            range: 'A1',
            majorDimension: 'ROWS',
            values: [
                [error.url, error.error]
            ]
        };
        gsapiError.spreadsheets.values.append({
            spreadsheetId: spreadsheetId,
            range: 'Error!A1',
            valueInputOption: 'USER_ENTERED',
            resource: errorBody
        }, (err, res) => {
            if (err) {
                console.log(err);
            } else {
                console.log(res.data);
            }
        });
    });
};

const writeToGoogleSheet = async (cl) => {
    const json = fs.readFileSync('urunler.json');
    const data = await JSON.parse(json);
    const gsapi = google.sheets({ version: 'v4', auth: cl });
    data.forEach(async (item) => {
        const params = {
            spreadsheetId: spreadsheetId,
            range: 'Data!A:E',
            valueInputOption: 'RAW',
            insertDataOption: 'INSERT_ROWS',
            resource: {
                values: [

                    // name: item.marka,
                    // urun_adi: item.urun_adi,
                    // urun_url: item.urun_url,
                    // eski_fiyat: item.eski_fiyat,
                    // indirimli_fiyat: item.indirimli_fiyat,
                    // indirim_oran: item.indirim_oran
                    item.marka,
                    item.urun_adi,
                    item.urun_url,
                    item.eski_fiyat,
                    item.indirimli_fiyat,
                    item.indirim_oran,
                    item.urun_kodu
                ]
            }
        };
        gsapi.spreadsheets.values.append(params).then(async (res) => {
            console.log(res.data, "Success");
        }).catch(err => {
            console.log(err, "Error");
        });
    });
};

const getDataFromWebSite = async () => {
    for (let url of URLs) {
        try {
            await axios.get(url).then(response => {
                const $ = cheerio.load(response.data);
                const urunler = [];
                $('div[class="fl col-12 product-item-content"]').each((i, el) => {
                    const urun = {};
                    urun.marka = $(el).find('a > span').text();
                    urun.urun_adi = $(el).find('a').attr('title');
                    urun.urun_url = web_site_url + $(el).find('a').attr('href');
                    urun.eski_fiyat = $(el).find('div.product-item-price span.discountedPrice').text().replaceAll('\n', ' ').replaceAll('\t', '').replaceAll('\r', '');
                    urun.indirimli_fiyat = $(el).find('div.product-item-price span.currentPrice').text().replaceAll('\n', ' ').replaceAll('\t', '').replaceAll('\r', '');
                    urun.indirim_oran = $(el).find('div.product-item-discount span').text().replaceAll('\n', ' ').replaceAll('\t', '').replaceAll('\r', '');
                    urun.urun_kodu = $(el).find('div.product-item-code').text().replaceAll('\n', ' ').replaceAll('\t', '').replaceAll('\r', '');
                    urunler.push(urun);
                });
                if (fs.existsSync('urunler.json')) {
                    const data = fs.readFileSync('urunler.json');
                    const json = JSON.parse(data);
                    json.push(urunler);
                    fs.writeFileSync('urunler.json', JSON.stringify(json));
                } else {
                    fs.writeFileSync('urunler.json', JSON.stringify(urunler));
                }
                // const urunler_json = JSON.stringify(urunler);
                // fs.writeFileSync('urunler.json', urunler_json);
                const urunler_csv = new json2csv(urunler, {
                    json: true,
                    fields:
                        [
                            'marka',
                            'urun_adi',
                            'urun_url',
                            'eski_fiyat',
                            'indirimli_fiyat',
                            'indirim_oran',
                            'urun_kodu'
                        ]
                });
                fs.writeFileSync('urunler.csv', urunler_csv.toString());
                console.log('Dosya oluşturuldu.');
                const urunler_xlsx = xlsx.utils.json_to_sheet(urunler);
                console.log(url);
                const urunler_xlsx_workbook = xlsx.utils.book_new();
                xlsx.utils.book_append_sheet(urunler_xlsx_workbook, urunler_xlsx, 'urunler');
                xlsx.writeFile(urunler_xlsx_workbook, 'urunler.xlsx');
            });
        } catch (error) {
            const errorBody = [
                {
                    error: error.message,
                    url: url
                }
            ]
            if (fs.existsSync('error.json')) {
                const data = fs.readFileSync('error.json');
                const json = JSON.parse(data);
                json.push(errorBody);
                fs.writeFileSync('error.json', JSON.stringify(json));
            } else {
                fs.writeFileSync('error.json', JSON.stringify(errorBody));
            }

        }
    }
    googleSheetAuth();
}

var workbook = xlsx.readFile('URL.xlsx');
var sheet_name_list = workbook.SheetNames;
sheet_name_list.forEach(function (y) {
    var worksheet = workbook.Sheets[y];
    var headers = {};
    var data = [];
    for (z in worksheet) {
        if (z[0] === '!') continue;
        //parse out the column, row, and value
        var tt = 0;
        for (var i = 0; i < z.length; i++) {
            if (!isNaN(z[i])) {
                tt = i;
                break;
            }
        };
        var col = z.substring(0, tt);
        var row = parseInt(z.substring(tt));
        var value = worksheet[z].v;

        //store header names
        if (row == 1 && value) {
            headers[col] = value;
            continue;
        }

        if (!data[row]) data[row] = {};
        data[row][headers[col]] = value;
    }
    //drop those first two rows which are empty
    data.shift();
    data.shift();
    data.forEach(row => {
        URLs.push(web_site_url + row['/']);
    });
    getDataFromWebSite();
});