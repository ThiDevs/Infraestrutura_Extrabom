import win32com.client
import win32com

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
accounts = win32com.client.Dispatch("Outlook.Application").Session.Accounts

print(win32com.client.Dispatch("Outlook.Application").Session.AddressLists("Personal Address Book"))


def email_messages(folder):
    messages = folder.Items
    a = len(messages)
    dic = {}
    if a > 0:
        for message in messages:
            try:
                dic.update({message.Body.split("\n")[2].split("|")[1]: "0"})
            except:
                message.Close(0)
    return dic


def main():
    folder = outlook.Folders(accounts[0].DeliveryStore.DisplayName).Folders[1]
    Demissao = outlook.Folders(accounts[0].DeliveryStore.DisplayName).Folders(folder.name).Folders[1]
    arq = open("usuariofluig.csv", "rt")
    lines = arq.readlines()[1:]
    dic = email_messages(Demissao)
    i = 0
    for line in lines:
        if dic.get(line.split(";")[1].upper()) == "0":
            if line.split(";")[3].upper() == 'ACTIVE':
                print(line.split(";")[1].upper())
                i += 1
    print(i)
    arq.close()


main()
