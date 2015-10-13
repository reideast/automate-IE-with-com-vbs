# automate-IE-with-com-vbs
A VBScript automation script to scrape data from and submit data to Internet Explorer

This script was created to automate access to a web-based database. The requirement was to search for a numeric ID, update some properties in the resultant HTML form, save, and print the results. Initially developed with a co-worker, the finalized script was revised, debugged, commented, and tested for final use by the team on a whole. VBScript was chosen for compatibility with Windows.

Functions include: connect to an Internet Explorer COM object, load a page into IE, wait for pages to load and properly reconnect to page, read information from an HTML form, submit information into the page, write to log files, read and write to CSV files, print web pages, & get verified input from user.

(Script has been modified to remove identifying information, but core VBScript methods to work with IE COM objects and the HTML DOM are left intact for reference.)
