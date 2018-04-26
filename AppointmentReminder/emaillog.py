
def emaillog(title, msg):
    import win32com.client
    from win32com.client import Dispatch, constants

    const=win32com.client.constants
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = title
    newMail.Body = msg
    newMail.To = "mrwilliams@seton.org"
    # newMail.CC = "dparsons@seton.org"
    # attachment1 = r"E:\test\logo.png"

    # newMail.Attachments.Add(Source=attachment1)
    newMail.display()
    try:
        newMail.send()
    except:
        pass