import win32com.client
import datetime
import sql_server_tools

out_App = win32com.client.gencache.EnsureDispatch("Outlook.Application")
gal = out_App.Session.GetGlobalAddressList()
entries = gal.AddressEntries

#wyczyść tabele roboczą
trunc_comm = 'TRUNCATE TABLE [dbo].[Outlook_Address_Book_buff]'
sql_server_tools.sql_server_trusted_conn_insert_stmt('PROC2016005\\PROC2016005', 'michal_work', trunc_comm)


def variable_input(var):
    if var == '':
        output = 'brak danych'
    else:
        output = var
    return output


for entry in entries:
    user = entry.GetExchangeUser()
    try:
        name = variable_input(user.Name)
    except:
        name = 'brak danych'
    try:
        last_name = variable_input(user.LastName)
    except:
        last_name = 'brak danych'
    try:
        first_name = variable_input(user.FirstName)
    except:
        first_name = 'brak danych'
    try:
        alias = variable_input(user.Alias)
    except:
        alias = 'brak danych'
    try:
        mail_address = variable_input(user.PrimarySmtpAddress)
    except:
        mail_address = 'brak danych'
    try:
        mobilePhone = variable_input(user.MobileTelephoneNumber)
    except:
        mobilePhone = 'brak danych'
    try:
        businessPhone = variable_input(user.BusinessTelephoneNumber)
    except:
        businessPhone = 'brak danych'
    try:
        department = variable_input(user.Department)
    except:
        department = 'brak danych'
    try:
        job_title = variable_input(user.JobTitle)
    except:
        job_title = 'brak danych'
    try:
        office_location = variable_input(user.OfficeLocation)
    except:
        office_location = 'brak danych'

    comm_str = f"""INSERT INTO [dbo].[Outlook_Address_Book_buff] (
       [Name]
      ,[FirstName]
      ,[LastName]
      ,[Alias]
      ,[MailAddress]
      ,[MobilePhone]
      ,[BusinessPhone]
      ,[Department]
      ,[JobTitle]
      ,[OfficeLocation]
      ,[data_raportu])
VALUES
('{name}', '{last_name}', '{first_name}', '{alias}', '{mail_address}', '{mobilePhone}', '{businessPhone}', '{department}', '{job_title}', '{office_location}', '{datetime.date.today()}')"""
    if name not in ('WOLF Administracja', 'brak danych'):
        sql_server_tools.sql_server_trusted_conn_insert_stmt('PROC2016005\\PROC2016005', 'michal_work', comm_str)
