function exportXlsx (data, filename = 'data') {
  function _s2ab (s) {
    const buf = new ArrayBuffer(s.length)
    const view = new Uint8Array(buf)
    for (let i = 0; i !== s.length; ++i) {
      view[i] = s.charCodeAt(i) & 0xFF
    }
    return buf
  }

  function _toBlob (data) {
    data.unshift(['材料欠料清单'])
    data.unshift(['爱科通电子（深圳）有限公司'])
    const ws = XLSX.utils.aoa_to_sheet(data)
    ws['!merges'] = [
      {
        s: {
          c: 0,
          r: 0
        },
        e: {
          c: 8,
          r: 0
        }
      },
      {
        s: {
          c: 0,
          r: 1
        },
        e: {
          c: 8,
          r: 1
        }
      }
    ]
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'mysheet')
    const wbout = XLSX.write(wb, {type: 'binary', bookType: 'xlsx'})
    return new Blob([_s2ab(wbout)], {type: 'application/octet-stream'})
  }

  const a = document.createElement('a')
  const blobObj = _toBlob(data)
  a.download = `${filename}.xlsx` // 下载名称
  a.href = URL.createObjectURL(blobObj) // 创建对象超链接绑定a标签
  document.body.appendChild(a)
  a.click() // 模拟点击实现下载
  setTimeout(_ => {
    URL.revokeObjectURL(blobObj)
    document.body.removeChild(a)
  }, 100) // 延时释放, 用URL.revokeObjectURL()来释放这个object URL
}
