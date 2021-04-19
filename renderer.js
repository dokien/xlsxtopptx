const importElm = document.getElementById('import')
const exportElm = document.getElementById('export')
const status = document.getElementById('status')
status.style.color = 'white'
importElm.addEventListener('click', async () => {
  try {
    const result = await window.electron.importExcel();
    console.log(result);
    status.textContent = 'import success!'
  } catch (e) {
    status.textContent = `import fail : ${e.message}`
  }
})
exportElm.addEventListener('click', async () => {
  try {
    const result = await window.electron.exportPptx();
    status.textContent = result ? 'export success!' : 'export fail!'
  } catch (e) {
    status.textContent = `export fail : ${e.message}`
  }
})
