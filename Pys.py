import win32com.client
import win32com

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
accounts = win32com.client.Dispatch("Outlook.Application")

for i in accounts.Session.AddressLists:
    if str(i) == 'Lista de Endereços Global':
        myAddressList = accounts.Session.AddressLists(str(i))
        print(i)
        myFolder = outlook.GetDefaultFolder(myAddressList)
        print(myFolder)

        break
