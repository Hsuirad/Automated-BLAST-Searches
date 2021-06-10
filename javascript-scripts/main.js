const puppeteer = require('puppeteer');
const fs = require('fs')
const XLSX = require('xlsx');

let argv = require('yargs/yargs')(process.argv.slice(2)).argv;

var workbook = XLSX.readFile(argv.path, {
    type: "string"
});
/* DO SOMETHING WITH workbook HERE */

var first_sheet_name = workbook.SheetNames[0];
var address_of_cell = 'A1';

/* Get worksheet */
var worksheet = workbook.Sheets[first_sheet_name];

/* Find desired cell */
var desired_cell = worksheet[address_of_cell];

let entryNames = []
let hasGene = []

let quickLookup = {}

let cellMax = 0

for(cell in worksheet){
    if(cell.substring(1) * 1 > cellMax){
        cellMax = cell.substring(1) * 1
    }
}

for(let i = 2; i < 56; i++){
    let cell = "J" + i
    console.log(worksheet["A1"].v)
    console.log(worksheet[cell].f)
    // console.log(worksheet[cell].l)
    // console.log(cell)
    let value = worksheet[cell].f.substring(worksheet[cell].f.indexOf("./"), worksheet[cell].f.indexOf('Link to FASTA') - 3)
    console.log(value, worksheet[cell].f.substring(worksheet[cell].f.indexOf("./"), worksheet[cell].f.indexOf('Link to FASTA') - 3))
    let part = worksheet[cell].f.substring(worksheet[cell].f.indexOf("./output-data/") + "./output-data/".length, worksheet[cell].f.indexOf('/","'))
    let data = fs.readFileSync(value + part + ".fasta", {encoding:'utf8', flag:'r'});
    console.log(data)
    entryNames.push(data)
    quickLookup[data] = "J" + cell.substring(1);

    hasGene.push(worksheet["I"+i].v == "N/A" ? true : false);
}

// for(cell in worksheet){
//     // console.log(cell)
//     // console.log(worksheet[cell].v)

//     if(/[A-Z]{1}[0-9]{1,}/.test(cell)){
//         //you could check the E value from the C cell idiot
//         switch(cell[0]){
//             case "E":
//                 console.log(worksheet[cell].v + ".")
//                 hasGene.push(worksheet[cell].v == "N/A" && cell != "E1" ? true : false)
//                 break;
//             default:
//                 break;
//         }
//     }
// }

const filename = 'forlab.xlsx';

let entriesToBlast = []

for(let i = 0; i < hasGene.length; i++){
    if(hasGene[i] == true && i != 0 && entryNames[i] !== undefined){
        entriesToBlast.push(entryNames[i])
    }
}

console.log(quickLookup)
//XLSX.utils.sheet_add_aoa(worksheet, [['NEW VALUE from NODE']], {origin: 'D4'});
// console.log(XLSX.utils.sheet_add_aoa(worksheet, [['TEST']], {origin: 'G38'}))
// console.log(worksheet['G38'])

const URL = 'https://blast.ncbi.nlm.nih.gov/Blast.cgi?PROGRAM=blastp&PAGE_TYPE=BlastSearch&LINK_LOC=blasthome'
let organism, query;

let scrape = async(entry) =>  {
    const browser = await puppeteer.launch({headless: true});
    try{
        const page = await browser.newPage();
        await page.setViewport({
            width: 1000,
            height: 700,
            deviceScaleFactor: 1
        })
        await page.goto(URL);

        await page.waitForFunction("document.querySelector('[id=qorganism]')", {timeout: 20000})
        

        const typeIn = await page.evaluate(() => {
            organism = document.querySelector('[id=qorganism]')//.value = "human (taxid:9606)"
            query = document.querySelector('[name=QUERY]')//.value = 'A0A0D9SAW5_CHLSB'

            return [organism, query]
        });
        await page.type('[id=qorganism]', "human (taxid:9606)")//, {delay:10})
        // await page.type('[id=qorganism]', String.fromCharCode(13));
        await page.type('[name=QUERY]', entry)//, {delay: 100})
        
        await page.click('.blastbutton');

        await page.waitForFunction("document.querySelector('[id=dscTable]')", {timeout: 3*60*1000})
        console.log('HIT')
        
        console.log('PINGED')

        const evalPage = await page.evaluate(() => {
            let xx = [...document.querySelector('[id=dscTable]').children[2].children].map(e => e.innerText)
            let yy = xx.map((e, i) => i < 6 ? e.split('\n').map(e => e.trim()).filter(e => e.length > 0) : "").filter(e => e)
            let eVals = []
            let accs = []
            let descs = []
            let iterator = 1

            yy.forEach(e => {
                //accesion, e-value, description
                // final.push([
                accs.push([`Accession #${iterator}: ${e.map(r => r.trim()).filter(r => r).pop().split('\t').pop()}`])
                eVals.push([`E-value #${iterator}: ${e.map(r => r.trim()).filter(r => r).pop().split('\t')[4]}`])
                descs.push([`Description #${iterator}: ${e.map(r => r.trim()).filter(r => r)[1]}`])
                iterator++
                // ])
            })
            return [accs, eVals, descs]
        })


        console.log('PINGED LAST')


        console.log(evalPage)
        
        await browser.close();

        return [evalPage, entry]
    }
    catch(err){
        await browser.close()

        return ['None Found', entry]
    }
    
}

let running = 0;

let workbookAdditions = []

let timer = setInterval(() => {
    if(running < 8 && entriesToBlast.length > 0){
        console.log(entriesToBlast.length)
        scrape(entriesToBlast.pop()).then(e => {
            console.log(e)
            workbookAdditions.push(e)
            running--; //current problem is, i need try catch to make sure running-- happens even if not properly run
        })
        running++;
    }

    else if(entriesToBlast.length == 0 && running == 0){
        saveToSheet()
        clearInterval(timer)
    }
}, 500)

let saveToSheet = () => {
    console.log('done')
    workbookAdditions.forEach(e => {
        console.log(e)
        console.log('\n\n\n\n\n\n\n\n\n\n\n' + e)
        console.log('\n\n' + e[1] + '\n\n')
        console.log(quickLookup[e[1]])
        if(e[0] !== "None Found"){
            console.log(e[0][0])
            fs.writeFileSync(`./output-data/${worksheet["B" + quickLookup[e[1]].substring(1)].v}/BLAST_Results.txt`, [e[0][0]].join(', ') + "\n\n" + [e[0][1]].join(', ') + "\n\n" + [e[0][2]].join(', '))
            XLSX.utils.sheet_add_aoa(worksheet, [[[e[0][0]].join(', ')]], {origin: "G" + quickLookup[e[1]].substring(1)});
            XLSX.utils.sheet_add_aoa(worksheet, [[[e[0][1]].join(', ')]], {origin: "H" + quickLookup[e[1]].substring(1)});
            XLSX.utils.sheet_add_aoa(worksheet, [[[e[0][2]].join(', ')]], {origin: "F" + quickLookup[e[1]].substring(1)});
        } else{
            XLSX.utils.sheet_add_aoa(worksheet, [["None Found"]], {origin: "G" + quickLookup[e[1]].substring(1)});
            XLSX.utils.sheet_add_aoa(worksheet, [["None Found"]], {origin: "H" + quickLookup[e[1]].substring(1)});
            XLSX.utils.sheet_add_aoa(worksheet, [["None Found"]], {origin: "F" + quickLookup[e[1]].substring(1)});
        }
    })
    XLSX.writeFile(workbook, 'out.xlsx');
}
