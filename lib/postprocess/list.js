module.exports = (doc) => {
  const lists = doc.getElementsByTagName('jsreport:list')
  for (let i = 0; i < lists.length; i++) {
    const wpElement = lists[i].parentNode.parentNode.parentNode
    const items = lists[i].getElementsByTagName('jsreport:item')
    const insertBoreElement = wpElement.nextSibling

    for (let j = 1; j < items.length; j++) {
      const clonedItem = wpElement.cloneNode(true)
      const listNode = clonedItem.getElementsByTagName('jsreport:list')[0]
      listNode.parentNode.textContent = items[j].textContent
      wpElement.parentNode.insertBefore(clonedItem, insertBoreElement)
    }

    lists[i].parentNode.textContent = items[0].textContent
  }
}
