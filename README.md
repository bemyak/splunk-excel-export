# What is this?
This is the [Splunk](https://www.splunk.com/) plugin used to export multiple reports in one Excel template.  
The plugin consistently fills XLSX sheets with reports starting from the first sheet.  
If the next sheet in template is missing, new sheet will be created with name `SheetX` where X in a report serial number.  

So if you have 4 reports you can create template with 4 empty sheets and on the fifth one you can create some analytical reports, diagrams, etc.

# Installation Instructions
## 1. Prepare template

Sheets that are expected to be filled is recommended to keep empty, any data on them may be lost.

> **Important!** The template must be in `xlsx` format.

Upload prepared template to Splunk server into some directory (ex. `/opt/splunk/templates/template.xlsx`).
Don't use `/tmp/` dir case it may be cleaned up during reboot.

## 2. Installing dependencies
The only dependency is [openpyxl](https://pypi.org/project/openpyxl/).  
It must be installed to the Splunk python dir (`/opt/splunk/lib/python2.7/site-packages/` in my case).

The easiest way is to use `epel` repo on CentOs:
```bash
# Plugin epel repo
yum install -y epel-release
# Install pip
yum install -y python2-pip
# Install openpyxl
pip install openpyxl -t /opt/splunk/lib/python2.7/site-packages/
```

> **Important!** If you are using manual installation don't forget to install openpyxl's dependency [jdcal](https://pypi.org/project/jdcal/) 

## 3. Plugin installation
1. Delete directory `/opt/splunk/etc/apps/excel_export` if it exists
1. Create directory `/opt/splunk/etc/apps/excel_export`
2. Copy all *directories* from this repo into created dir.
3. Restart Splunk using `/opt/splunk/bin/splunk restart`

## 4. Adding plugin to your report page
### Filling a variable on search complete
Add the following block into each `<search/>` section:
```xml
<done>
    <set token="search1_sid">$job.sid$</set>
</done>
```
Notice that each `token` must be unique (ex. for first search use `search1_sid`, for second - `search2_sid` and so on)

### Add export widget
Add one extra panel on report page:
```xml
<row>
    <panel>
        <html depends="$search1_sid$,$search2_sid$">
            <a href='/custom/excel_export/excel?filename=test.xlsx&amp;template=/opt/splunk/templates/template.xlsx&amp;sid1=$search1_sid$&amp;sid2=$search2_sid$'>Export to Excel</a>
        </html>
    </panel>
</row>
```
The `depends` attribute contains all comma-separated tokens from the previous step. They must be wrapped in `$` signs.  
This section needs for splunk not to render it until all search jobs complete.

The request parameters are:
* `filename` - user will download file with this name
* `template` - path to template on filesystem
* `sid1`, `sid2` and so on - all of the tokens from the previous section wrapped in `$` sign  
The attributes are separated using the `&amp;` separator.
