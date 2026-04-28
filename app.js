import * as XLSX from 'xlsx'

let workbook = null
let currentSheetIndex = 0
let allData = null
let headers = []
let rows = []
let usageData = []
const VISIBLE_ROWS = 30
const START_COLUMN = 6

const uploadArea = document.getElementById('uploadArea')
const fileInput = document.getElementById('fileInput')
const fileInfo = document.getElementById('fileInfo')
const fileName = document.getElementById('fileName')
const fileSize = document.getElementById('fileSize')
const btnRemove = document.getElementById('btnRemove')
const uploadProgress = document.getElementById('uploadProgress')
const progressFill = document.getElementById('progressFill')
const progressText = document.getElementById('progressText')
const btnParse = document.getElementById('btnParse')
const navItems = document.querySelectorAll('.nav-item')
const sections = document.querySelectorAll('.section')
const sheetSelector = document.getElementById('sheetSelector')
const sheetInfo = document.getElementById('sheetInfo')
const tableContainer = document.getElementById('tableContainer')
const dataStats = document.getElementById('dataStats')
const rowCount = document.getElementById('rowCount')
const colCount = document.getElementById('colCount')
const cellCount = document.getElementById('cellCount')
const btnRefresh = document.getElementById('btnRefresh')
const btnExportJSON = document.getElementById('btnExportJSON')
const btnExportCSV = document.getElementById('btnExportCSV')
const btnExportTXT = document.getElementById('btnExportTXT')
const previewContent = document.getElementById('previewContent')
const toast = document.getElementById('toast')
const toastMessage = document.getElementById('toastMessage')

function showToast(message, type = 'success') {
  toastMessage.textContent = message
  toast.className = `toast ${type} show`
  setTimeout(() => {
    toast.classList.remove('show')
  }, 3000)
}

function formatFileSize(bytes) {
  if (bytes < 1024) return bytes + ' B'
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(2) + ' KB'
  return (bytes / (1024 * 1024)).toFixed(2) + ' MB'
}

function updateProgress(text, percent) {
  progressText.textContent = text
  progressFill.style.width = percent + '%'
}

function decodeText(buffer, encoding) {
  try {
    const decoder = new TextDecoder(encoding)
    return decoder.decode(buffer)
  } catch {
    return null
  }
}

function fixEncoding(text) {
  if (!text || typeof text !== 'string') return text
  
  const originalText = text
  
  if (/[\u0080-\u00FF]/.test(text)) {
    const bytes = new Uint8Array(text.split('').map(c => c.charCodeAt(0)))
    
    const encodings = ['GBK', 'GB2312', 'Big5', 'GB18030']
    for (const encoding of encodings) {
      try {
        const decoded = decodeText(bytes, encoding)
        if (decoded && /[\u4e00-\u9fff]/.test(decoded)) {
          return decoded
        }
      } catch {}
    }
  }
  
  try {
    const utf8Text = decodeURIComponent(escape(text))
    if (/[\u4e00-\u9fff]/.test(utf8Text)) {
      return utf8Text
    }
  } catch {}
  
  return originalText
}

function toNumber(value) {
  if (value === undefined || value === null || value === '') return 0
  if (typeof value === 'number') return value
  const str = String(value).trim()
  if (str === '') return 0
  const num = parseFloat(str.replace(/,/g, ''))
  return isNaN(num) ? 0 : num
}

function parseExcelFileWithCodepage(data) {
  const options = { type: 'array' }
  
  try {
    return XLSX.read(data, { ...options, codepage: 936 })
  } catch {
    try {
      return XLSX.read(data, { ...options, codepage: 65001 })
    } catch {
      return XLSX.read(data, options)
    }
  }
}

function asyncParseExcel(data, fileExtension) {
  return new Promise((resolve) => {
    setTimeout(() => {
      let workbook
      
      if (fileExtension === '.xls') {
        workbook = parseExcelFileWithCodepage(data)
      } else {
        workbook = XLSX.read(data, { type: 'array' })
      }
      
      resolve(workbook)
    }, 0)
  })
}

async function parseExcelFile(file) {
  uploadProgress.style.display = 'block'
  fileInfo.style.display = 'none'
  btnParse.disabled = true
  uploadArea.classList.remove('dragover')

  updateProgress('正在读取文件...', 10)

  try {
    const arrayBuffer = await readFileAsArrayBuffer(file)
    updateProgress('正在解析数据...', 30)

    const extension = file.name.toLowerCase().substring(file.name.lastIndexOf('.'))
    workbook = await asyncParseExcel(new Uint8Array(arrayBuffer), extension)
    updateProgress('正在处理数据...', 60)

    await populateSheetSelectorAsync()
    await renderTableAsync(0)

    updateProgress('完成', 100)

    setTimeout(() => {
      uploadProgress.style.display = 'none'
      fileInfo.style.display = 'flex'
      btnParse.disabled = false
      showToast('文件解析成功！')
    }, 300)

  } catch (error) {
    uploadProgress.style.display = 'none'
    fileInfo.style.display = 'flex'
    showToast('文件解析失败: ' + error.message, 'error')
    btnParse.disabled = false
  }
}

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = (e) => resolve(e.target.result)
    reader.onerror = () => reject(new Error('文件读取失败'))
    reader.readAsArrayBuffer(file)
  })
}

async function populateSheetSelectorAsync() {
  return new Promise((resolve) => {
    setTimeout(() => {
      sheetSelector.innerHTML = ''
      
      workbook.SheetNames.forEach((name, index) => {
        const option = document.createElement('option')
        option.value = index
        option.textContent = fixEncoding(name)
        sheetSelector.appendChild(option)
      })

      sheetSelector.disabled = false
      sheetSelector.selectedIndex = 0
      sheetInfo.innerHTML = `工作表: <span>${fixEncoding(workbook.SheetNames[0])}</span>`
      resolve()
    }, 0)
  })
}

async function renderTableAsync(sheetIndex) {
  return new Promise((resolve) => {
    setTimeout(() => {
      renderTable(sheetIndex)
      resolve()
    }, 0)
  })
}

function renderTable(sheetIndex) {
  if (!workbook) return

  const sheetName = workbook.SheetNames[sheetIndex]
  const worksheet = workbook.Sheets[sheetName]
  allData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false })

  if (allData.length === 0) {
    tableContainer.innerHTML = `
      <div class="empty-state">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
          <rect x="3" y="3" width="18" height="18" rx="2" ry="2"/>
          <line x1="9" y1="3" x2="9" y2="21"/>
          <line x1="15" y1="3" x2="15" y2="21"/>
        </svg>
        <p>该工作表为空</p>
      </div>
    `
    dataStats.style.display = 'none'
    return
  }

  const originalHeaders = allData[0]
  const dataRows = allData.slice(1)

  if (originalHeaders.length <= START_COLUMN) {
    tableContainer.innerHTML = `
      <div class="empty-state">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
          <rect x="3" y="3" width="18" height="18" rx="2" ry="2"/>
          <line x1="9" y1="3" x2="9" y2="21"/>
          <line x1="15" y1="3" x2="15" y2="21"/>
        </svg>
        <p>数据不足，需要至少G列（第7列）的数据</p>
      </div>
    `
    dataStats.style.display = 'none'
    return
  }

  usageData = []
  
  for (let col = START_COLUMN; col < originalHeaders.length; col++) {
    const name = fixEncoding(String(originalHeaders[col] || `列 ${col + 1}`)).trim()
    
    if (!name || name === '' || /^[\s\t]*$/.test(name)) {
      continue
    }
    
    let sum = 0
    for (const row of dataRows) {
      const cellValue = row[col]
      sum += toNumber(cellValue)
    }
    
    usageData.push({
      name: name,
      usage: sum
    })
  }

  headers = ['名称', '使用量']
  rows = usageData.map(item => [item.name, item.usage.toLocaleString()])

  const totalRows = usageData.length

  rowCount.textContent = totalRows
  colCount.textContent = 2
  cellCount.textContent = totalRows * 2
  dataStats.style.display = 'flex'

  renderVirtualTable()
}

function renderVirtualTable() {
  const tableHTML = `
    <div class="virtual-table-container">
      <div class="virtual-table-header">
        <table class="data-table-header">
          <thead>
            <tr>
              <th style="width: 60px;">#</th>
              ${headers.map((header, index) => `<th data-index="${index}">${header}</th>`).join('')}
            </tr>
          </thead>
        </table>
      </div>
      <div class="virtual-table-body" id="virtualBody">
        ${renderVisibleRows(0)}
      </div>
      <div class="table-footer">
        <div class="pagination-info">
          <span>共 ${rows.length} 项数据</span>
        </div>
        <div class="pagination-controls" id="paginationControls">
          ${createPagination()}
        </div>
      </div>
    </div>
  `

  tableContainer.innerHTML = tableHTML

  setupPagination()
}

function renderVisibleRows(startIndex) {
  const endIndex = Math.min(startIndex + VISIBLE_ROWS, rows.length)
  const visibleData = rows.slice(startIndex, endIndex)

  let rowsHTML = ''
  for (let i = 0; i < visibleData.length; i++) {
    const rowIndex = startIndex + i
    const row = visibleData[i]
    
    rowsHTML += `<tr data-row="${rowIndex + 1}">`
    rowsHTML += `<td>${rowIndex + 1}</td>`
    
    for (let j = 0; j < headers.length; j++) {
      const cellValue = row[j] || ''
      rowsHTML += `<td title="${cellValue}">${cellValue}</td>`
    }
    
    rowsHTML += '</tr>'
  }

  return `<table class="data-table-body"><tbody>${rowsHTML}</tbody></table>`
}

let currentPage = 1

function createPagination() {
  const totalPages = Math.ceil(rows.length / VISIBLE_ROWS)
  if (totalPages <= 1) return ''

  let paginationHTML = ''
  
  paginationHTML += `<button class="page-btn" data-page="prev" disabled>上一页</button>`
  
  for (let i = 1; i <= totalPages; i++) {
    if (i <= 3 || i > totalPages - 2 || (i >= currentPage - 1 && i <= currentPage + 1)) {
      paginationHTML += `<button class="page-btn ${i === 1 ? 'active' : ''}" data-page="${i}">${i}</button>`
    } else if (i === 4 && totalPages > 6) {
      paginationHTML += `<span class="page-dots">...</span>`
    }
  }
  
  paginationHTML += `<button class="page-btn" data-page="next">下一页</button>`

  return paginationHTML
}

function setupPagination() {
  const container = document.getElementById('paginationControls')
  if (!container) return

  container.addEventListener('click', (e) => {
    const target = e.target
    if (!target.classList.contains('page-btn')) return

    const page = target.dataset.page
    const totalPages = Math.ceil(rows.length / VISIBLE_ROWS)

    if (page === 'prev') {
      if (currentPage > 1) currentPage--
    } else if (page === 'next') {
      if (currentPage < totalPages) currentPage++
    } else {
      currentPage = parseInt(page)
    }

    updatePagination(totalPages)
    updateVisibleRows()
  })
}

function updatePagination(totalPages) {
  const container = document.getElementById('paginationControls')
  if (!container) return

  const buttons = container.querySelectorAll('.page-btn')
  buttons.forEach(btn => {
    const page = btn.dataset.page
    if (page === 'prev') {
      btn.disabled = currentPage === 1
    } else if (page === 'next') {
      btn.disabled = currentPage === totalPages
    } else {
      btn.classList.toggle('active', parseInt(page) === currentPage)
    }
  })
}

function updateVisibleRows() {
  const startIndex = (currentPage - 1) * VISIBLE_ROWS
  const body = document.getElementById('virtualBody')
  if (body) {
    body.innerHTML = renderVisibleRows(startIndex)
  }
}

function exportToJSON() {
  if (!workbook || usageData.length === 0) {
    showToast('请先上传文件并解析数据', 'warning')
    return
  }

  const sheetName = workbook.SheetNames[currentSheetIndex]
  const exportData = usageData.map(item => ({
    name: item.name,
    usage: item.usage
  }))
  
  downloadFile(JSON.stringify(exportData, null, 2), `${fixEncoding(sheetName)}_使用量统计.json`, 'application/json')
  showToast('JSON文件导出成功！')
}

function exportToCSV() {
  if (!workbook || usageData.length === 0) {
    showToast('请先上传文件并解析数据', 'warning')
    return
  }

  const sheetName = workbook.SheetNames[currentSheetIndex]
  let csvData = '名称,使用量\n'
  
  usageData.forEach(item => {
    const name = String(item.name).replace(/,/g, '，')
    csvData += `${name},${item.usage}\n`
  })
  
  const blob = new Blob(['\uFEFF' + csvData], { type: 'text/csv;charset=UTF-8' })
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = `${fixEncoding(sheetName)}_使用量统计.csv`
  document.body.appendChild(link)
  link.click()
  document.body.removeChild(link)
  URL.revokeObjectURL(url)
  
  showToast('CSV文件导出成功！')
}

function exportToTXT() {
  if (!workbook || usageData.length === 0) {
    showToast('请先上传文件并解析数据', 'warning')
    return
  }

  const sheetName = workbook.SheetNames[currentSheetIndex]
  let txtData = `名称\t使用量\n`
  
  usageData.forEach(item => {
    txtData += `${item.name}\t${item.usage}\n`
  })
  
  downloadFile(txtData, `${fixEncoding(sheetName)}_使用量统计.txt`, 'text/plain;charset=UTF-8')
  showToast('TXT文件导出成功！')
}

function downloadFile(content, filename, mimeType) {
  const blob = new Blob([content], { type: mimeType })
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  
  link.href = url
  link.download = filename
  document.body.appendChild(link)
  link.click()
  document.body.removeChild(link)
  URL.revokeObjectURL(url)
}

uploadArea.addEventListener('click', () => {
  fileInput.click()
})

uploadArea.addEventListener('dragover', (e) => {
  e.preventDefault()
  uploadArea.classList.add('dragover')
})

uploadArea.addEventListener('dragleave', () => {
  uploadArea.classList.remove('dragover')
})

uploadArea.addEventListener('drop', (e) => {
  e.preventDefault()
  uploadArea.classList.remove('dragover')
  
  const files = e.dataTransfer.files
  if (files.length > 0) {
    handleFile(files[0])
  }
})

fileInput.addEventListener('change', (e) => {
  if (e.target.files.length > 0) {
    handleFile(e.target.files[0])
  }
})

function handleFile(file) {
  const validExtensions = ['.xlsx', '.xls', '.csv']
  const extension = file.name.toLowerCase().substring(file.name.lastIndexOf('.'))
  
  if (!validExtensions.includes(extension)) {
    showToast('不支持的文件格式，请上传Excel文件', 'error')
    return
  }

  if (file.size > 50 * 1024 * 1024) {
    showToast('文件大小超过50MB限制', 'error')
    return
  }

  fileName.textContent = file.name
  fileSize.textContent = formatFileSize(file.size)
  
  fileInfo.style.display = 'flex'
  btnParse.disabled = false
}

btnRemove.addEventListener('click', () => {
  fileInput.value = ''
  fileInfo.style.display = 'none'
  btnParse.disabled = true
  workbook = null
  allData = null
  headers = []
  rows = []
  usageData = []
  currentPage = 1
  sheetSelector.innerHTML = ''
  sheetSelector.disabled = true
  sheetInfo.innerHTML = '工作表: <span>未选择文件</span>'
  tableContainer.innerHTML = `
    <div class="empty-state">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
        <rect x="3" y="3" width="18" height="18" rx="2" ry="2"/>
        <line x1="9" y1="3" x2="9" y2="21"/>
        <line x1="15" y1="3" x2="15" y2="21"/>
      </svg>
      <p>请先上传Excel文件以查看数据</p>
    </div>
  `
  dataStats.style.display = 'none'
  previewContent.textContent = ''
})

btnParse.addEventListener('click', async () => {
  if (fileInput.files.length > 0) {
    currentPage = 1
    await parseExcelFile(fileInput.files[0])
    
    setTimeout(() => {
      const dataNav = document.querySelector('[data-section="data"]')
      if (dataNav) {
        dataNav.click()
      }
    }, 400)
  }
})

navItems.forEach(item => {
  item.addEventListener('click', () => {
    const targetSection = item.dataset.section
    
    navItems.forEach(nav => nav.classList.remove('active'))
    item.classList.add('active')
    
    sections.forEach(section => {
      section.classList.remove('active')
      if (section.id === `${targetSection}-section`) {
        section.classList.add('active')
      }
    })

    if (targetSection === 'data' && workbook && rows.length > 0) {
      currentPage = 1
      renderVirtualTable()
    }

    if (targetSection === 'export' && workbook && usageData.length > 0) {
      const previewData = usageData.slice(0, 5).map(item => ({
        name: item.name,
        usage: item.usage
      }))
      previewContent.textContent = JSON.stringify(previewData, null, 2) + (usageData.length > 5 ? '\n... 更多数据' : '')
    }
  })
})

sheetSelector.addEventListener('change', (e) => {
  currentSheetIndex = parseInt(e.target.value)
  currentPage = 1
  const sheetName = workbook.SheetNames[currentSheetIndex]
  sheetInfo.innerHTML = `工作表: <span>${fixEncoding(sheetName)}</span>`
  renderTable(currentSheetIndex)

  if (usageData.length > 0) {
    const previewData = usageData.slice(0, 5).map(item => ({
      name: item.name,
      usage: item.usage
    }))
    previewContent.textContent = JSON.stringify(previewData, null, 2) + (usageData.length > 5 ? '\n... 更多数据' : '')
  }
})

btnRefresh.addEventListener('click', async () => {
  if (fileInput.files.length > 0) {
    currentPage = 1
    await parseExcelFile(fileInput.files[0])
  }
})

btnExportJSON.addEventListener('click', exportToJSON)
btnExportCSV.addEventListener('click', exportToCSV)
btnExportTXT.addEventListener('click', exportToTXT)