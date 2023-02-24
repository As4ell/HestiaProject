import win32com.client
import pyodbc

out_App = win32com.client.gencache.EnsureDispatch("Outlook.Application")
gal = out_App.Session.GetGlobalAddressList()
entries = gal.AddressEntries

global_address_items = []
# print(type(gal))
for entry in entries:
    user = entry.GetExchangeUser()
    try:
        str_name = user.Name
    except:
        str_name = 'brak danych'
    try:
        str_last_name = user.LastName
    except:
        str_last_name = 'brak danych'
    try:
        str_first_name = user.FirstName
    except:
        str_first_name = 'brak danych'
    try:
        str_alias = user.Alias
    except:
        str_alias = 'brak danych'
    try:
        str_mail_address = user.PrimarySmtpAddress
    except:
        str_mail_address = 'brak danych'
    try:
        str_mobilephone = user.MobileTelephoneNumber
    except:
        str_mobilephone = 'brak danych'
    try:
        str_businessPhone = user.BusinessTelephoneNumber
    except:
        str_businessPhone = 'brak danych'
    try:
        str_department = user.Department
    except:
        str_department = 'brak danych'
    try:
        str_job_title = user.JobTitle
    except:
        str_job_title = 'brak danych'
    try:
        office_location = user.OfficeLocation
    except:
        office_location = 'brak danych'


    #except:
        #str_location
    # print(str_name, str_alias, str_mail_address, str_mobilephone, str_businessPhone, str_department, str_job_title)

# SQL_CNXN = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};SERVER=PROC2016005\PROC2016005; '
 #                         'DATABASE=CIC_raportowanie; Trusted_Connection=yes', autocommit=True, timeout=0)

# print(type(SQL_CNXN))




