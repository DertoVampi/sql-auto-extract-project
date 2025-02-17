I use this program daily to automatically extract data and perform correction to obtain one or two Excel files to be delivered via FTP to a customer.

Previously, this process was made manually, by downloading PowerBI tables as .xlsx and adding them with copy-paste to a single file. This process was clearly outdated, and I first implemented a quick Python script to automatically extract data from PowerBI
and concatenate it in a single Excel, then I upgraded it to extract data directly from our SQL Server database hosted on Azure. Doing so reduced the time needed from at least 1 hour to around 5-6 minutes.
