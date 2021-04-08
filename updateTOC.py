import win32com.client


def update_toc(file_name):
    word = win32com.client.DispatchEx("Word.Application")
    doc = word.Documents.Open(file_name)
    toc_count = doc.TablesOfContents.Count
    if toc_count == 1:
        toc = doc.TablesOfContents(1)
        toc.Update()
        print('TOC should have been updated.')
    else:
        print('TOC has not been updated for sure...')

    doc.Close(SaveChanges=True)
    word.Quit()
