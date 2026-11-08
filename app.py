files = ['ProductData.xlsx','MasterInvoice.pdf','Order1.xlsx','Order2.xlsx','Order3.xlsx','Order4.xlsx']
while True:
    directory = single_input("Please enter the directory where the ProductData.xslx and other files are located.")
    directory = directory.strip('"')
    if directory[-1] != '\\':
        directory = directory + '\\'
    else:
        directory = directory
    filenames = os.listdir(directory)
    if 'ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' and 'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames :
        break
    else:
        show_message("Error", "The directory you entered is not recognized. Please try again.")

product_descriptions = {}
product_prices = {}
df = pd.read_excel(directory + 'ProductData.xlsx', sheet_name = 'Descriptions')
df1 = pd.read_excel(directory + 'ProductData.xlsx',sheet_name = 'Prices')
df.head()
for index,row in df.iterrows():
    product_descriptions[row['Unit number']] = row['Description']
for index,row in df1.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']
def get_order_data(invoice_num,filename,directory):
    order_data = {}
    template = openpyxl.load_workbook(directory + filename)
    selected_sheet = template['OrderForm']
    order_data['InvoiceNum'] = invoice_num
    order_data['Date'] = selected_sheet['C9'].value
    order_data['Date'] = order_data['Date'].strftime('%m/%d/%Y')
    order_data['Total'] = 0
  
    if selected_sheet['B24'].value is not None:
        order_data['Price0'] = product_prices[selected_sheet['B24'].value]
        order_data['Quantity0'] = selected_sheet['C24'].value
        order_data['Amount0'] = product_prices[selected_sheet['B24'].value] * selected_sheet['C24'].value
        order_data['Total'] = (order_data['Price0'] * order_data['Quantity0'])
        order_data['Description0'] = product_descriptions[selected_sheet['B24'].value]

    if selected_sheet['B25'].value is not None:
        order_data['Price1'] = product_prices[selected_sheet['B25'].value]
        order_data['Quantity1'] = selected_sheet['C25'].value
        order_data['Amount1'] = product_prices[selected_sheet['B25'].value] * selected_sheet['C25'].value
        order_data['Total'] = (order_data['Price0'] * order_data['Quantity0']) + (order_data['Price1'] * order_data['Quantity1'])
        order_data['Description1'] = product_descriptions[selected_sheet['B25'].value]

    if selected_sheet['B26'].value is not None:
        order_data['Price2'] = product_prices[selected_sheet['B26'].value]
        order_data['Quantity2'] = selected_sheet['C26'].value
        order_data['Amount2'] = product_prices[selected_sheet['B26'].value] * selected_sheet['C26'].value
        order_data['Total'] = (order_data['Price0'] * order_data['Quantity0']) + (order_data['Price1'] * order_data['Quantity1']) + (order_data['Price2'] * order_data['Quantity2'])
        order_data['Description2'] = product_descriptions[selected_sheet['B26'].value]

    if selected_sheet['B27'].value is not None:
        order_data['Price3'] = product_prices[selected_sheet['B27'].value]
        order_data['Quantity3'] = selected_sheet['C27'].value
        order_data['Amount3'] = product_prices[selected_sheet['B27'].value] * selected_sheet['C27'].value
        order_data['Total'] = (order_data['Price0'] * order_data['Quantity0']) + (order_data['Price1'] * order_data['Quantity1']) + (order_data['Price2'] * order_data['Quantity2']) + (order_data['Price3'] * order_data['Quantity3'])
        order_data['Description3'] = product_descriptions[selected_sheet['B27'].value]

    if selected_sheet['B28'].value is not None:
        order_data['Price4'] = product_prices[selected_sheet['B28'].value]
        order_data['Quantity4'] = selected_sheet['C28'].value
        order_data['Amount4'] = product_prices[selected_sheet['B28'].value] * selected_sheet['C28'].value
        order_data['Total'] = (order_data['Price0'] * order_data['Quantity0']) + (order_data['Price1'] * order_data['Quantity1']) + (order_data['Price2'] * order_data['Quantity2']) + (order_data['Price3'] * order_data['Quantity3']) + (order_data['Price4'] * order_data['Quantity4'])
        order_data['Description4'] = product_descriptions[selected_sheet['B28'].value]

    if selected_sheet['B24'].value is not None:
        order_data['Price0'] = f'${order_data['Price0']:,.2f}'
        order_data['Amount0'] = f'${order_data['Amount0']:,.2f}'
        order_data['Quantity0'] = str(order_data['Quantity0'])

    if selected_sheet['B25'].value is not None:
        order_data['Price1'] = f'${order_data['Price1']:,.2f}'
        order_data['Amount1'] = f'${order_data['Amount1']:,.2f}'
        order_data['Quantity1'] = str(order_data['Quantity1'])

    if selected_sheet['B26'].value is not None:
        order_data['Price2'] = f'${order_data['Price2']:,.2f}'
        order_data['Amount2'] = f'${order_data['Amount2']:,.2f}'
        order_data['Quantity2'] = str(order_data['Quantity2'])

    if selected_sheet['B27'].value is not None:
        order_data['Price3'] = f'${order_data['Price3']:,.2f}'
        order_data['Amount3'] = f'${order_data['Amount3']:,.2f}'
        order_data['Quantity3'] = str(order_data['Quantity3'])

    if selected_sheet['B28'].value is not None:
        order_data['Price4'] = f'${order_data['Price4']:,.2f}'
        order_data['Amount4'] = f'${order_data['Amount4']:,.2f}'
        order_data['Quantity4'] = str(order_data['Quantity4'])

    order_data['Total'] = f'${order_data['Total']:,.2f}'

    order_data['CustomerAddress1'] = selected_sheet['C16'].value
    order_data['CustomerName'] = selected_sheet['C14'].value
    order_data['CustomerCompany'] = selected_sheet['C15'].value
    order_data['CustomerAddress2'] = selected_sheet['C17'].value + ", " + selected_sheet['C18'].value + " " + str(selected_sheet['C19'].value)

    name = order_data['CustomerName']
    email = selected_sheet['C13'].value
    date = order_data['Date']

    return order_data, name, email, date 

order_data, name, email, date = get_order_data(invoice_num= '8515' , filename = 'Order1.xlsx' , directory = directory)
def create_invoice(invoice_num,order_data,directory):
    create_pdf(directory + 'MasterInvoice.pdf', directory + f'Invoice {invoice_num}.pdf', order_data)

test2 = create_invoice(invoice_num=8515,order_data = order_data, directory = directory)
gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()
def send_email(invoice_num,name,email,date,directory):

    gmail_user = "siennajbranch@gmail.com"
    recipients = email
    subject = f'Your order from PythonTech Co - Invoice {invoice_num}'
    body = f"""Dear {name},
    
Attached is your invoice dated {date}.

Thank you for your business.

PythonTech Co.
10 PYthon Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)
    path_to_file = f'{directory}\\Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_file, message) 
    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)
    mail_server.send_message(message)
    mail_server.quit()
filenames = os.listdir(directory)
invoice_num = 8515 
for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        order_data, name, email, date = get_order_data(invoice_num = invoice_num,filename = filename, directory = directory)
        create_invoice(invoice_num = invoice_num, order_data = order_data, directory = directory)
        send_email(invoice_num = invoice_num, name = name, email = email, date = date, directory = directory)
        invoice_num = invoice_num + 1