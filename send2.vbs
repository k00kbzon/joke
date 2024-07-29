Set MyApp = CreateObject("Outlook.Application")
Set MyItem = MyApp.CreateItem(0)
With MyItem
    .To = "finapps_dev_mx_grp@oracle.com"
    .Subject = "Donas!!"
    .ReadReceiptRequested = False
    .HTMLBody = "Donas para todos"
End With
MyItem.Send
