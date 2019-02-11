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

# tests


if __name__ == '__main__':
    set_doc_author_keys("John Doe", "JD")
