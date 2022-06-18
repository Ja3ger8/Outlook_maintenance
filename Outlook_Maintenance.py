import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
#test_folder_original =  outlook.Folders.Item("sa_test1@outlook.com")
#inbox_original = test_folder_original.Folders.Item("folder2")
#msg = inbox_original.Items
#msgs = msg.GetLast()
#print (msgs)
#print (msgs.Subject)


#root_folder = outlook.Folders.Item(1)
#subfolder = root_folder.Folders['folder1'].Folders['folder2']


#your_folder = outlook.Folders['Outlook_Mails'].Folders['Inbox'].Folders['folder2']
#for message in your_folder.Items:
#    print(message.Subject)


def cleanup_inbox(outlook):
    inbox_original_mailbox = outlook.Folders.Item("sa_test1@outlook.com")
    inbox_original = inbox_original_mailbox.Folders.Item("Inbox")

    inbox_target_mailbox = outlook.Folders.Item("sa_test2@outlook.com")
    inbox_target = inbox_target_mailbox.Folders.Item("Inbox")



    original_mails = inbox_original.items
    for i in reversed(original_mails):
        print(i.Subject)
        if i.Subject == 'Fw: SA Rare Bird News Report - 13 June 2022' :
            i.Move(inbox_target)

    #for message in inbox_target.Items:
    #    print(message.Subject)


    #msg = inbox_original.Items
    #root_folder_original = outlook.Folders.Item(1)
    #root_folder_original = test_folder_original.Folders.Item("Inbox")
    #outlook_inbox_original = root_folder_original.Folders['Inbox']

    #test_folder_target = outlook.Folders.Item("sa_test2@outlook.com")
    #root_folder_target = outlook.Folders.Item(1)
    #outlook_inbox_original = root_folder_original.Folders['Inbox']

    #inbox_original = test_folder_original.Folders.Item("Inbox")


    #root_folder = outlook.Folders.Item(1)
    #outlook_inbox = root_folder.Folders['folder1'].Folders['folder2']

def cleanup_subfolders(outlook):

    subfolder_list = ['folder2']

    for folder in subfolder_list:

        subfolder_original = outlook.Folders.Item("sa_test1@outlook.com").Folders['folder1'].Folders[folder]

        subfolder_target = outlook.Folders.Item("sa_test2@outlook.com").Folders['folder1'].Folders[folder]

        #subfolder_original = outlook.Folders.Item("sa_test1@outlook.com").Folders['folder1'].Folders['folder2']

        #subfolder_target = outlook.Folders.Item("sa_test2@outlook.com").Folders['folder1'].Folders['folder2']

        original_mails = subfolder_original.items
        for i in reversed(original_mails):
            print(i.Subject)
            if i.Subject == 'Fw: SA Rare Bird News Report - 13 June 2022':
                i.Move(subfolder_target)


if __name__ == '__main__':
    #cleanup_inbox(outlook)
    cleanup_subfolders(outlook)