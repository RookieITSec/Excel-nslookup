# Excel-nslookup
Here are some Excel macros that I modified from others.  I wanted a place to save them for later use.  NSlookup, internal DNS lookup, External DNS lookup.


2019-04-22 RookieITSec- These notes are largely for my own future reference, but if you can also use them, awesome!  The scripts are rough, but they worked for what I needed to do (validate the configurations of a ton of (sub)domains we own).

### How to make this stuff work.  

Depending on your use case, you may only need a few of the modules, but for my use, I have them all in my spreadsheet.
I am using Excel 2010 on Windows 10 for this.

The results can be a bit odd if there are more-than-basic DNS records.  CNAME's or multiple IPv4's or IPV6 make things a bit odd, but for my usage, this was not a problem.

## Here is what the MS VB structure looks like - 
![](https://github.com/RookieITSec/Excel-nslookup/blob/master/ss-VB.png)

## Here are the formulas in Excel -
![](https://github.com/RookieITSec/Excel-nslookup/blob/master/ss-Formulas.png)

## Here are the calculated values in Excel - 
![](https://github.com/RookieITSec/Excel-nslookup/blob/master/ss-calced.png)


This may take some time to calculate and Excel will occasionally diem, so save often.  A good coder could probably make a few adjustments to fix this.



