import win32com.client

def build_doc(template, outfile, actions):
    wdApp = win32com.client.Dispatch("Word.Application")
    wdDoc = wdApp.Documents.Open(template)

    for i in actions:
        if i == 'add_doc_var':
            wdDoc.Variables.Add(actions[i]['name'], actions[i]['value'])
            wdDoc.Fields.Update

        if i == 'add_vba_module':
            vbaModule = wdDoc.VBProject.VBComponents.Add(1)
            vbaModule.CodeModule.AddFromString(actions['add_vba_module'])

    wdDoc.SaveAs(outfile, 0)
    wdDoc.Close()
    wdApp.Quit()
