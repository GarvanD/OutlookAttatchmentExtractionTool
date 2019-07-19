import eel
import win32com.client, os


eel.init('web')

@eel.expose
def getAttatchments(in0,in1,in2,in3,chk0,chk1,chk2,chk3,key_word):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    if chk0 and not chk1 and not chk2 and not chk3:
        inbox = outlook.Folders[in0]
    if chk0 and chk1 and not chk2 and not chk3:
        inbox = outlook.Folders[in0].Folders[in1]
    if chk0 and chk1 and chk2 and not chk3:
        inbox = outlook.Folders[in0].Folders[in1].Folders[in2]
    if chk0 and chk1 and chk2 and chk3:
        inbox = outlook.Folders[in0].Folders[in1].Folders[in2].Folders[in3]

    messages = inbox.Items
    message = messages.GetFirst()
    subject = message.Subject

    newpath = os.getcwd() + '//Attatchments'
    if not os.path.exists(newpath):
        os.makedirs(newpath)
    else:
        shutil.rmtree(newpath)
        os.makedirs(newpath)
    for m in messages:
        if key_word in m.Subject:
            attachments = message.Attachments
            num_attach = len([x for x in attachments])
            for x in range(1, num_attach+1):
                attachment = attachments.Item(x)
                attachment.SaveASFile(os.path.join(newpath,attachment.FileName))
        message = messages.GetNext()


eel.start('index.html')