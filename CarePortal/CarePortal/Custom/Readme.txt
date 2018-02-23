CUSTOMISING THE LOOK AND FEEL OF THE PORTAL PAGES

The display of the active contact from the portal can be customised such that it can appear within the framework of the overall website presentation by customising the head and body sections of the web pages produced by the portal to match the look and file of a web site.

The head and body sections may be customised in one of two ways to match the look and feel of the web site.

1/ Files may be created in the custom directory to provide a static presentation of the frame in which the portal pages are presented, which will be the same for each page. For this method three files should be created; head.htm, bodystart.htm and bodyend.htm

The minimum data required in each of these files would be as follows;

head.htm
  <head>
  </head>

bodystart.htm
  <body>

bodyend.htm
  </body>

Normally of course these files would contain information to provide the web page with customised page content.


2/ A dynamic method may be used to provide the presentation of the frame in which the portal pages are presented. This may differ for each page.

This method requires that a configuration file (custom.config) be placed in the custom directory. This file specifies a CustomURL from which the dynamic content will be retrieved. The minimum content of the configuration file is as follows;

<?xml version="1.0"?>
<CustomConfiguration>
  <CustomURL>
    https://localhost/CarePortal/custom/{0}.htm?CN={1}&amp;QSRT={2}&amp;FURL={3}
  </CustomURL>
</CustomConfiguration>

The CustomURL item shown above should be replaced with a URL capable of providing dynamic content. In the CustomURL as specified above the {0} item will be replaced by the portal with either head, bodystart or bodyend dependant on which section of the page is required. The {1} item will be replaced by the contact number of the currently logged in contact (This will be 0 if the user is not logged in), {2} item will be replaced by the query string of the page being served and {3} item will be replaced by the Friendly Url specified in the WPD for the served page.


STYLESHEET REFERENCES

If customisation is provided as described above and it is required that the default stylesheet for the portal still be used the following line should be added to the customised head.htm file;

<link href="Styles.css" type="text/css" rel="stylesheet" />

The following classes are used in the default stylesheet and if required may be overwritten in a customised head section;

HeaderItem
LeftItem
CenterItem
RightItem
FooterItem

Table
TableSelectedData
TableAlternateData
TableData
TableHeader
TableFooter

MainTable
MainTableRow
InnerTable
InnerTableRow

DataEntryLabel
DataEntryItem
DataEntryHelp
DataEntryViewItem
DataEntryCheckBox
DataMessage
DataEntryItemMandatory
DataValidator

DatePicker
OtherMonthDayStyle
TodayDayStyle
SelectedDayStyle
DatePickerTitle		


ReadOnlyNumber
Button
ReadOnly


3/ Portal may be configured to access the business functionality by making calls to the webservices (Default) or by calling them directly. Businesses can choose whichever option is best suited with the infrastructure.  Direct calls should never be used if the Portal server is accesssible from outside the firewall as the database should not be direct accessible from a machine that has public access.
 
 This method requires that a configuration file (custom.config) be placed in the custom directory. This file specifies UseWebServices option with possible value(s) of 'Y' or 'N'
 
 Y - will make portal call webservices for accessing business functionality.
 N - will make direct calls to the business functionality.  

<?xml version="1.0"?>
<CustomConfiguration>
  <UserWebServices>
    <value>Y</value>
  </UserWebServices>    
</CustomConfiguration>  

If the custom.config file is not found under custom directory then portal will use web service calls to implement business functionality.  


4/ A dynamic method may be used for a Single sign-on and sign-off for 3rd party systems and NG Portal.

This method requires that a configuration file (custom.config) be placed in the custom directory. This file specifies SingleSignOnURL, SingleSignOffURL and SingleSignOnKey. The minimum content of the configuration file is as follows;

<?xml version="1.0"?>
<CustomConfiguration>
  <SingleSignOnURL>https://localhost/CarePortal/SingleSignOn.aspx?CN={0}&amp;HS={1}&amp;ReturnURL={2}</SingleSignOnURL>
  <SingleSignOffURL>https://localhost/CarePortal/SingleSignOff.aspx?ReturnURL={0}</SingleSignOffURL>
  <SingleSignOnKey>A shared secret key to create the Hash Value</SingleSignOnKey>
</CustomConfiguration>

The SingleSignOnURL item shown above should be replaced with a 3rd Party Sign-On URL. In the SingleSignOnURL as specified above the {0} item will be replaced by the contact number of the registered user or the user name when logged in as back end user. The {1} item will be replaced by the MD5 Hash Value generated from the SingleSignOnKey and the User Id (Contact Number/User Name) and the {2} item will be replaced by the URL that NG portal would have navigated to if Single Sign On was not in use.
The SingleSignOffURL item shown above should be replaced with a 3rd Party Sign-Off URL. In the SingleSignOffURL as specified above the {0} item will be replaced by the URL that NG portal would have navigated to if Single Sign Off was not in use.
The SingleSignOnKey item shown above should be replaced with a secret key which only the 3rd Party should know. NG Portal will use this key to generate the MD5 Hash Value for Single Sign-On. The MD5 Hash Value will be generated by concatenating this Key value with the User Id (Contact Number/User Name) e.g. if the Key is TESTZ and Contact Number is 456 then the MD5 Hash Value will be generated for the text TESTZ456.

Note: The MD5 hash value will be a 32-character, hexadecimal-formatted string.
