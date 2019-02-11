import win32com.client
from pathlib import Path
import os

def build_doc(template, outfile, actions):
    wdApp = win32com.client.Dispatch("Word.Application")
    wdDoc = wdApp.Documents.Open(template)

    save_format = 0  # office97 OLE Compound File - standard .doc
    #  add a 'set_save_format' entry to the playbook with docm or flatxml to change the format
    #  for other formats, see https://docs.microsoft.com/en-us/office/vba/api/word.wdsaveformat

    for i in actions:
        if 'add_doc_var' in i:
            wdDoc.Variables.Add(i['add_doc_var']['name'], i['add_doc_var']['value'])
            wdDoc.Fields.Update

        if 'add_vba_module' in i:
            vbaModule = wdDoc.VBProject.VBComponents.Add(1)
            vbaModule.CodeModule.AddFromString(i['add_vba_module'])

        if 'set_save_format' in i:
            if i['set_save_format'].lower() == 'docm':
                save_format = 13
            if i['set_save_format'].lower() == 'flatxml':
                save_format = 20

        if 'set_author' in i:
            wdAuthor = wdDoc.BuiltinDocumentProperties("Author")  # this is not working
            wdAuthor.Value = i['set_author']
            wdDoc.Fields.Update


    wdDoc.SaveAs(outfile, save_format)
    wdDoc.Close()
    wdApp.Quit()

    for i in actions:
        if 'change_extension_after_save' in i:
            renamed_file_path = os.path.join(Path(outfile).parent, Path(outfile).stem + "." + i['change_extension_after_save'])
            os.rename(outfile, renamed_file_path)
