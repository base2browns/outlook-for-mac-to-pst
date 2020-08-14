# Introduction

A friend of mine accidentally deleted a set of email folders and wanted to get them back. They had a time machine backup of their Outlook for Mac profiles but couldn't work out how to recover the emails.

To save the next developer the pain, I have provided the source code of what I put together here.

## How it works

Outlook for Mac appears to store it's emails in a profile folder. I was able to deduce that the files in the `Data\Message Sources` folder were essentially RFC822 (*.eml) files and could be read by both Outlook (for Windows) and other tools. 
In conjunction with this, the program uses a Sqlite database to hold it's state data. 

With these together I was able to extract the data and create a PST file with their emails. I attempted to create threads but never got it work quite right. The code is still there but disabled by default.

To run the tool:

```batch
dotnet run -- -p <path to the profile folder> [-s <the target store> -f <the target folder>] [-t]
```

I used the [Outlook Redemption](http://www.dimastr.com/redemption/home.htm) library to load everything into Outlook. The first time you run it, you can select the target folder. I then output the store and folder identifiers so you can pass them in next time.

A general thank you to the authors of the various libraries I used.

This software is provided without any warranty of any kind. Also, I have no intention of providing any support for it, just feel free to fork the repository.