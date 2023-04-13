import * as XLSX from 'xlsx/xlsx.mjs'
import * as fs from 'fs'
import path from 'path'
import { fileURLToPath } from 'url'

XLSX.set_fs(fs)

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)
const filepath = process.env.FILEPATH
const workbook = XLSX.readFile(filepath)
const sheetName = process.env.sheetName
const sheet = workbook.Sheets[sheetName]
const deaconDuties = XLSX.utils.sheet_to_json(sheet)

let html = ''

for ( const duty of deaconDuties ) {
    const deaconJobDescription = duty['Deacon Job Description']
    const title = deaconJobDescription.substring(0, deaconJobDescription.indexOf(':'))
    const description = deaconJobDescription.substring(deaconJobDescription.indexOf(':') + 1).trim()
    const lead = duty.Lead
    const assist = duty.Assist
    const nonDeaconAssist = duty['Non-Deacon Assist']
    const elders = duty['Go-to Elders']

    let men = ''

    if (lead) men += `<div><b style="margin-right: 8px;">Lead:</b><span>${lead.split('\r\n').join(', ')}</span></div>`
    if (assist) men += `<div><b style="margin-right: 8px;">Assist:</b><span>${assist.split('\r\n').join(', ')}</span></div>`
    if (nonDeaconAssist) men += `<div><b style="margin-right: 8px;">Non-Deacon Assist:</b><span>${nonDeaconAssist.split('\r\n').join(', ')}</span></div>`
    if (elders) men += `<div><b style="margin-right: 8px;">Go-to Elders:</b><span>${elders.split('\r\n').join(', ')}</span></div>`

    html += `<h3>${title}</h3><p>${description}</p>${men}<p>&nbsp;</p>`
}

fs.writeFileSync(path.join(__dirname, `${new Date().toISOString().substring(0, 10)}_Deacon-Duties.html`), html)