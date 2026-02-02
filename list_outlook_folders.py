import win32com.client

def list_folders(folder, indent=0):
    prefix = "  " * indent
    print(f"{prefix}- {folder.Name}")
    try:
        for sub in folder.Folders:
            list_folders(sub, indent + 1)
    except:
        pass

def main():
    print("Listing Outlook Folders Structure:")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        
        # Print Inbox structure (where 'דוחות' likely resides)
        list_folders(inbox)
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
