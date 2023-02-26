$config = @{
    settings = @{
        cloudFolderPath = "<your local cloud sync folder>"
        destinationFolderPath = "<where to copy missing files to>"
    }
    phones = @{
        oneplus = @{
            name = "<your phone's name as it is shown in file explorer>"
            folder = "<the full path (without the phone name) on your phone to"`
                     " the folder where you want to sync files from>"
        }
        samsung = @{
            # ...
        }
        iphone = @{
            # ...
        }
    }
}