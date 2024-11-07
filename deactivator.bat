Open cmd as administrator

Type Locate the location of MS Office, mine is installed in c:\Program files\Microsoft Office or in c:\Program files (x86)

Copy the MS office location address.

paste it on cmd and type cscript before the address line and press ENTER

Put double quote before and after the address line..

cscript "C:\Program Files\Microsoft Office\Office15\ospp.vbs" /dstatus and Press ENTER

OR

cscript "C:\Program Files\Microsoft Office\Office16\ospp.vbs" /dstatus and Press ENTER

There will be one text written as (Last 5 character of installed product key : XXXXX) the poduct key displays the last 5 digit of product key

Now, press arrow up key to redo the first command.

Delete dstatus and type unpkey: <Enter the Last 5 character of installed product key> and Press ENTER