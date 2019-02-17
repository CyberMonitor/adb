import winreg

def set_doc_author_keys(userinitials, username):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Office\Common\UserInfo', 0,
                             winreg.KEY_ALL_ACCESS) as key:
            winreg.SetValueEx(key, "UserName", 0, winreg.REG_SZ, username)
            winreg.SetValueEx(key, "UserInitials", 0, winreg.REG_SZ, userinitials)
    except FileNotFoundError:
        print("[!] Office may not be installed and initially run to create the necessary key locations. "
              "Please install and open Office once to set it up.")
        raise FileNotFoundError

#  returns the list of playbook items to allow for the creation of the file, and if necessary, renaming it to the
#  desired extension
#
# set_save_format determines the file format
# set_save_extension will be the extension compatible with the type of file you're saving.
# docm format can't be named .doc when saved by word, but it will work if renamed afterward
#
# set_extension_after_save will indicate the file needs renamed after it is saved
#
def reconcile_extension_and_format(extension, filetype):
    extension = extension.lower()
    filetype = filetype.lower()
    if extension == "doc" and filetype == "flatxml":
        return {'set_save_format': 'flatxml',
                'set_save_extension': 'doc'}
    elif extension == "doc" and filetype == "docm":
        return {'set_save_format': 'docm',
                'set_save_extension': 'docm',
                'change_extension_after_save': 'doc'}
    elif extension == "docm" and filetype == "docm":
        return {'set_save_format': 'docm',
                'set_save_extension': 'docm'}
    elif extension == "doc" and filetype == "doc":
        return {'set_save_format': 'doc',
                'set_save_extension': 'doc'}
    elif extension == "rtf" and filetype == "docm":
        return {'set_save_format': 'docm',
                'set_save_extension': 'docm',
                'change_extension_after_save': 'rtf'}
    elif extension == "rtf" and filetype == "doc":
        return {'set_save_format': 'doc',
                'set_save_extension': 'doc',
                'change_extension_after_save': 'rtf'}
    elif extension == "rtf" and filetype == "flatxml":
        return {'set_save_format': 'flatxml',
                'set_save_extension': 'doc',
                'change_extension_after_save': 'rtf'}
    else:
        print("\n[!] Combination of format {} and extension {} is invalid.\n".format(extension, filetype))
        raise ValueError
# tests


if __name__ == '__main__':
    set_doc_author_keys("John Doe", "JD")
