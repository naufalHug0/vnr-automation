const constants = {
  CONFIGURATION_TEMPLATE: {
      totalRespondents: 50,
      variables: [
          {
              indicatorsCount: 6,
              reversed: false,
              scale: { low: 1, high: 5 }
          }
      ]
  },
  VALIDITY_SHEET_NAME: "Validity & Multicol",
  RESULT_SHEET_NAME: "Hasil Validity & Reliability",
  COLORS: {
      VALIDITY: { 
    NOT_VALID_RESULT: "#ea9999" ,
    RESULT_TABLE: {
      COLUMN_NAMES: {
        TEXT: ['white','black','black','white'],
        BACKGROUND: ['#ed7d31','#f7caac','#f7caac','#ed7d31']
      },
      ROWS: { BACKGROUND: ['white','white','white','#b6d7a7'] }
    }
  },
      RELIABILITY: {
          RESULT_RENDAH: { TEXT: "white", BACKGROUND: "red" },
          RESULT_SEDANG: { BACKGROUND: "#ff9900" },
          RESULT_TINGGI: { BACKGROUND: "yellow" },
      }
  },
  TEXTS: {
      VALIDITY: { 
    NOT_VALID_RESULT: "Tidak Valid",
    RESULT_TABLE: {
      TITLE: "Ringkasan Hasil Uji Validitas",
      COLUMN_NAMES: ["No Soal", "rxy", "rtabel", "Status"]
    }
  },
      RELIABILITY: {
          RESULT_RENDAH: "Rendah",
          RESULT_SEDANG: "Sedang",
          RESULT_TINGGI: "Tinggi",
      }
  }
}

function main() {
console.log("BARU")
  new AutomationApp().runScreenDetector()
}

function openConfiguration(width, height) {
  new AutomationApp().start(width, height)
}

function processAutomation(configuration) {
new VnrTemplate(configuration)
}

class AutomationApp {
  constructor() {
      this.ui = new UI()
  }

  runScreenDetector() {
      this.ui.detectScreenSize()
  }

  start(width, height) {
      this.ui.openModal(width, height)
  }
}

class UI {
  constructor() {
      this.screenDetector = {
          fileName: "ScreenSizeDetector",
          modalTitle: "Loading...",
      }
      this.vnr = {
          fileName: "Config",
          modalTitle: "Konfigurasi Automasi KX",
      }
  }

  detectScreenSize() {
      const html = HtmlService.createHtmlOutputFromFile(this.screenDetector.fileName).setWidth(10).setHeight(10)
      SpreadsheetApp.getUi().showModalDialog(html, this.screenDetector.modalTitle)
  }

  openModal(width, height) {
      const html = HtmlService.createHtmlOutputFromFile(this.vnr.fileName)
      .setWidth(width)
      .setHeight(height);
      SpreadsheetApp.getUi().showModalDialog(html, this.vnr.modalTitle)
  }
}

class SpreadsheetUtils {
  constructor(totalRespondents) {
      this.totalRespondents = totalRespondents
  }

  setupVnrInstance(context, resultSheet) {
      this.context = context
  }

  columnNumberToLetter(column) {
      var letter = "";
      while (column > 0) {
          var temp = (column - 1) % 26;
          letter = String.fromCharCode(temp + 65) + letter
          column = Math.floor((column - 1) / 26)
      }
      return letter
  }

  getIndicatorColumnRange(col_index) {
      const colLetter = this.columnNumberToLetter(col_index)
      return `${colLetter}2:${colLetter}${1+this.totalRespondents}`
  }

  setupColorConditionals() {
      const startRange = 2
      const reliabilityStartCol = 2
      const reliabilityStartRow = 5 + this.context.totalIndicators

      const rules = [
          SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(constants.TEXTS.VALIDITY.NOT_VALID_RESULT)
          .setBackground(constants.COLORS.VALIDITY.NOT_VALID_RESULT)          
          .setRanges([this.context.resultSheet.getRange(4, startRange + 3, this.context.totalIndicators, 1)])
          .build(),
          ...[
              SpreadsheetApp.newConditionalFormatRule()
              .whenTextContains(constants.TEXTS.RELIABILITY.RESULT_RENDAH)
              .setBackground(constants.COLORS.RELIABILITY.RESULT_RENDAH.BACKGROUND)
              .setFontColor(constants.COLORS.RELIABILITY.RESULT_RENDAH.TEXT),
              SpreadsheetApp.newConditionalFormatRule()
              .whenTextEqualTo(constants.TEXTS.RELIABILITY.RESULT_SEDANG)
              .setBackground(constants.COLORS.RELIABILITY.RESULT_SEDANG.BACKGROUND),
              SpreadsheetApp.newConditionalFormatRule()
              .whenTextContains(constants.TEXTS.RELIABILITY.RESULT_TINGGI)
              .setBackground(constants.COLORS.RELIABILITY.RESULT_TINGGI.BACKGROUND)
          ].map(rule => rule.setRanges([ this.context.resultSheet.getRange(reliabilityStartRow + 6 + this.totalRespondents, reliabilityStartCol+1) ]).build())
      ]

      const conditionalFormatRules = this.context.resultSheet.getConditionalFormatRules()
      rules.forEach(rule => conditionalFormatRules.push(rule))
      this.context.resultSheet.setConditionalFormatRules(conditionalFormatRules)
  }
}

class VnrTemplate extends SpreadsheetUtils {
  constructor(configuration = constants.CONFIGURATION_TEMPLATE) {
      super(configuration.totalRespondents)

      this.configuration = configuration
      this.spreadSheetApp = SpreadsheetApp.getActiveSpreadsheet()
      this.totalIndicators = configuration.variables.reduce((sum, v) => sum + v.indicatorsCount, 0)

      this.resultSheet = this.spreadSheetApp.getSheetByName(constants.RESULT_SHEET_NAME)
      this.cellAddresses = { correls: [], scales: [] }

      this.setupVnrInstance(this)
      this.createTemplate()
      this.setupColorConditionals(this.resultSheet)
  }

createTemplate() {
  new ValidityAnalysisTemplateBuilder(this)
  new ValidityResultBuilder(this)
  new ReliabilityResultBuilder(this)
}
}

class ValidityAnalysisTemplateBuilder {
constructor(context) {
  this.context = context
  this.configuration = this.context.configuration
  this.validitySheet = this.context.spreadSheetApp.getSheetByName(constants.VALIDITY_SHEET_NAME)

  this.cellAddresses = this.context.cellAddresses
  this.totalIndicators = this.context.totalIndicators

  this.createTemplate()
}

createTemplate() {
      const totalCols = this.totalIndicators + (2 * this.configuration.variables.length)
      let counter = 0
      let j = 0
      let currCol = ""
      let sumCol = ""

      for (let i = 1; i <= totalCols; i++) {
          currCol = this.context.columnNumberToLetter(i - this.configuration.variables[j].indicatorsCount + 1)
          sumCol = this.context.columnNumberToLetter(i+1)
          counter++

          this.fillIndicatorCol(i, j)
          this.addCorrelCellAddress(i)
          this.addScaleCellAddress(i)
          
          if (counter == this.configuration.variables[j].indicatorsCount) {
              counter = 0

              this.addGapBetweenCol(i)
              this.fillSumCol(i, currCol)
              this.fillCorrelCell(i, j, currCol, sumCol)

              i+=2
              j++
          }
      }

      this.addNumberCol()
  }

  getRandomizedScale() {
      const randBetween = "4,5"
      return Array.from({ length: this.context.totalRespondents }, () => [`=RANDBETWEEN(${randBetween})`])
  }

  fillIndicatorCol(colNumber, currVariableIndex) {
      const scale = this.getRandomizedScale()
      const highestScale = this.configuration.variables[currVariableIndex].scale.high

      this.validitySheet.getRange(2, colNumber, this.configuration.totalRespondents).setValues(scale)

      this.validitySheet.getRange(this.configuration.totalRespondents + 4, colNumber, highestScale + 1, 1)
          .setValues([["COUNTS"], ...Array.from({ length: highestScale }, (_, scaleIndex) => [`=${scaleIndex+1}&" = "&COUNTIF(${this.context.getIndicatorColumnRange(colNumber)}, ${scaleIndex+1})`])])
          .setHorizontalAlignment('center')
          .setFontWeights([["bold"], ...Array(highestScale).fill(["normal"])])
  }

  addCorrelCellAddress(colNumber) {
      this.cellAddresses.correls.push(`${this.context.columnNumberToLetter(colNumber+1)}${this.configuration.totalRespondents+2}`)
  }

  addScaleCellAddress(colNumber) {
      let col = []

      for (let row = 2; row <= this.configuration.totalRespondents + 1; row++) {
          col.push(`='${constants.VALIDITY_SHEET_NAME}'!${this.context.columnNumberToLetter(colNumber+1)}${row}`)
      }

      this.cellAddresses.scales.push(col)
  }

  addGapBetweenCol(colNumber, colCount = 2) {
      this.validitySheet.insertColumnsAfter(colNumber, colCount)
  }

  fillSumCol(colNumber, currCol) {
      this.validitySheet.getRange(1, colNumber+1)
          .setValue("SUM")
          .setFontWeight('bold')
          .setHorizontalAlignment('center')

      this.validitySheet.getRange(1, colNumber + 1, 1, 2).setBackground('white')

      this.validitySheet.getRange(2, colNumber+1, this.configuration.totalRespondents).setFormulas(
          Array.from({ length: this.configuration.totalRespondents }, (_, index) => [`=SUM(${currCol}${2+index}:${this.context.columnNumberToLetter(colNumber)}${2+index})`])
      ).setHorizontalAlignment('center')
  }

  fillCorrelCell(colNumber, currVariableIndex, currCol, sumCol) {
      this.validitySheet.getRange(2 + this.configuration.totalRespondents, colNumber - this.configuration.variables[currVariableIndex].indicatorsCount + 1, 1, this.configuration.variables[currVariableIndex].indicatorsCount).setFormula(
          `=CORREL(${currCol}2:${currCol}${1 + this.configuration.totalRespondents}, $${sumCol}$2:$${sumCol}$${1 + this.configuration.totalRespondents})`
      )
      .setNumberFormat("0.00")
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBorder(true,true,true,true,true,true)
  }

  addNumberCol() {
    const sequenceOfNumber = Array.from({ length: this.configuration.totalRespondents }, (_, i) => [i + 1])

    this.validitySheet.insertColumnsBefore(1, 1)
    this.validitySheet.getRange("A1")
      .setValue("No")
      .setHorizontalAlignment('center')
      .setVerticalAlignment('center')
      .setBackground('white')
      .setFontWeight('bold')
      
    this.validitySheet.getRange(2, 1, this.configuration.totalRespondents)
      .setValues(sequenceOfNumber)
      .setHorizontalAlignment('center')
      .setFontWeight('bold')
  }
}

class ValidityResultBuilder {
constructor(context) {
  this.context = context

  this.resultSheet = this.context.resultSheet
  this.minValidity = 0.75
  this.maxValidity = 0.89

  this.createTable()
}

createTable() {
  const startCol = 2
  let rowValues = []
  let rowFontWeights = []
  let rowBackgrounds = []
  let dataSheetName = `'${constants.VALIDITY_SHEET_NAME}'!`

  for (let i = 0; i < this.context.totalIndicators; i++) {
    rowValues[i] = [i+1, `=${dataSheetName + this.context.cellAddresses.correls[i]}`, this.minValidity.toString(), `=IF(AND(${dataSheetName+this.context.cellAddresses.correls[i]} > ${this.minValidity.toString()}, ${dataSheetName+this.context.cellAddresses.correls[i]} <= ${this.maxValidity.toString()}), "Valid","Tidak Valid")`]
    rowFontWeights[i] = ['normal','normal','normal','bold']
    rowBackgrounds[i] = constants.COLORS.VALIDITY.RESULT_TABLE.ROWS.BACKGROUND
  }

  this.createTableTitle(startCol)

  this.createColHeaders(startCol)

  this.createRows(startCol, rowValues, rowFontWeights, rowBackgrounds)
}

createTableTitle(startCol) {
  this.resultSheet.getRange(2, startCol, 1, 4).setValue(constants.TEXTS.VALIDITY.RESULT_TABLE.TITLE).merge().setFontWeight("bold").setHorizontalAlignment('center').setBorder(true,true,true,true,true,true)
}

createColHeaders(startCol) {
  this.resultSheet.getRange(3, startCol, 1, 4).setValues([
    constants.TEXTS.VALIDITY.RESULT_TABLE.COLUMN_NAMES
  ]).setHorizontalAlignment('center')
  .setFontWeight('bold')
  .setBorder(true,true,true,true,true,true)
  .setBackgrounds([constants.COLORS.VALIDITY.RESULT_TABLE.COLUMN_NAMES.BACKGROUND])
  .setFontColors([constants.COLORS.VALIDITY.RESULT_TABLE.COLUMN_NAMES.TEXT])
}

createRows(startCol, rowValues, rowFontWeights, rowBackgrounds) {
  this.resultSheet.getRange(4, startCol, this.context.totalIndicators, 4).setValues(rowValues).setHorizontalAlignment('center').setBorder(true,true,true,true,true,true).setFontWeights(rowFontWeights).setBackgrounds(rowBackgrounds)
}
}

class ReliabilityResultBuilder {
constructor(context) {
  this.context = context
  this.resultSheet = this.context.resultSheet

  this.createTable()
}

createTable() {
  const startCol = 2
  const startRow = 5 + this.context.totalIndicators

  this.createTableHeader(startRow, startCol)

  this.createTableRows(startRow, startCol)

  this.createVariansButirRow(startRow, startCol)

  const r11Cell = `${this.context.columnNumberToLetter(startCol+1)}${startRow + 5 + this.context.configuration.totalRespondents}`

  this.createTableFooter(startRow, startCol, r11Cell)

  this.createFooterMerge(startRow, startCol)
}

createTableHeader(startRow, startCol) { 
  this.resultSheet.getRange(startRow, startCol, 1, 2).setValues([["No. Responden", "Nomor Butir Angket"]]).setHorizontalAlignment('center').setBorder(true,true,true,true,true,true)

  this.resultSheet.getRange(startRow + 1, startCol + 1, 1, this.context.totalIndicators).setValues([
    Array.from({ length: this.context.totalIndicators }, (_,i) => [i+1])
  ]).setHorizontalAlignment('center').setBorder(true,true,true,true,true,true).setBackground('#adb9ca')

  this.resultSheet.getRange(startRow, startCol + this.context.totalIndicators + 1).setValue('Total').setHorizontalAlignment('center').setBorder(true,true,true,true,true,true)

  // merging
  this.resultSheet.getRange(startRow, startCol, 2).merge().setBackground('#1f3864').setFontColor('white')
  this.resultSheet.getRange(startRow, startCol + 1, 1, this.context.totalIndicators).merge().setBackground('#2f5496').setFontColor('white')
  this.resultSheet.getRange(startRow, startCol + this.context.totalIndicators + 1, 2).merge().setBackground('#1f3864').setFontColor('white')
}

createTableRows(startRow, startCol) {
  const firstCol = this.context.columnNumberToLetter(startCol+1)
  const lastCol = this.context.columnNumberToLetter(startCol+this.context.totalIndicators)
  
  this.resultSheet.getRange(startRow + 2, startCol, this.context.configuration.totalRespondents, this.context.totalIndicators + 2).setValues(
    Array.from({ length: this.context.configuration.totalRespondents }, (_,i) => [
      i+1, 
      ...Array.from({ length: this.context.totalIndicators}, (_,col) => this.context.cellAddresses.scales[col][i]),
      `=SUM(${firstCol}${startRow + 2 + i}:${lastCol}${startRow + 2 + i})`
    ]))
    .setHorizontalAlignment('center').setBorder(true,true,true,true,true,true).setBackgrounds(
    Array.from({ length: this.context.configuration.totalRespondents }, () => [
      'white', 
      ...Array.from({ length: this.context.totalIndicators }, () => '#fff2cc'),
      'white'
    ])
  )
}

createVariansButirRow(startRow, startCol) {
  this.resultSheet.getRange(startRow + 2 + this.context.configuration.totalRespondents, startCol, 1, this.context.totalIndicators + 2).setValues([
    [
    "Varians Butir",
    ...Array.from({ length: this.context.totalIndicators + 1 }, (_,index) => `=VAR(${this.context.columnNumberToLetter(startCol + index + 1)}${startRow + 2}:${this.context.columnNumberToLetter(startCol + index + 1)}${startRow + this.context.configuration.totalRespondents + 1})`)
    ]
  ]).setHorizontalAlignment('center').setBorder(true,true,true,true,true,true).setBackgrounds([
    [
    '#7b7b7b',
    ...Array.from({ length: this.context.totalIndicators }, () => 'white'),
    '#f9cb9c'
    ]
  ]).setNumberFormats([
    [
    "", ...Array.from({ length: this.context.totalIndicators + 1 }, () => "0.000")
    ]
  ]).setFontColors([
    [
    "white", ...Array.from({ length: this.context.totalIndicators + 1 }, () => "black")
    ]
  ])
}

createTableFooter(startRow, startCol, r11Cell) {
  this.resultSheet.getRange(startRow + 3 + this.context.configuration.totalRespondents, startCol, 4, 2).setValues([
    ["Jumlah Varians Butir", `=SUM(${this.context.columnNumberToLetter(startCol + 1)}${startRow+2+this.context.configuration.totalRespondents}:${this.context.columnNumberToLetter(startCol + this.context.totalIndicators)}${startRow + 2 + this.context.configuration.totalRespondents})`],
    ["Varians Total", `=${this.context.columnNumberToLetter(startCol + this.context.totalIndicators + 1)}${startRow + 2 + this.context.configuration.totalRespondents}`],
    ["r11", `=(1)*(1-(${this.context.columnNumberToLetter(startCol+1)}${startRow + 3 + this.context.configuration.totalRespondents}/${this.context.columnNumberToLetter(startCol+1)}${startRow + 4 + this.context.configuration.totalRespondents}))`],
    ["Reliabilitas", `=IF(${r11Cell}<0.2,"Sangat Rendah",IF(${r11Cell}<=0.4,"Rendah",IF(${r11Cell}<=0.6,"Sedang",IF(${r11Cell}<=0.8,"Tinggi","Sangat Tinggi"))))`]
  ]).setHorizontalAlignment('center').setBorder(true,true,true,true,true,true).setBackgrounds([
    ["#7f6000",'white'],
    ["#1e4e79",'white'],
    ["#ffc000",'white'],
    ["#00b0f0",'white'],
  ]).setFontColors([
    ...Array.from({ length: 2 }, () => ["white","black"]),
    ...Array.from({ length: 2 }, () => ["black","black"])
  ]).setFontWeights([
    ...Array.from({ length: 2 }, () => ["normal","normal"]),
    ...Array.from({ length: 2 }, () => ["bold","bold"])
  ]).setNumberFormat("0.000")
}

createFooterMerge(startRow, startCol) {
  this.resultSheet.getRange(startRow + 3 + this.context.configuration.totalRespondents, startCol + 2, 4, this.context.totalIndicators).setBorder(true,true,true,true,true,true).merge()
}
}