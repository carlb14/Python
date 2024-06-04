from openpyxl import load_workbook
import dumb_menu
from prompt_toolkit import prompt
from openpyxl.styles import Font, Alignment
from prompt_toolkit.completion import WordCompleter
from prompt_toolkit.history import InMemoryHistory
import datetime

print('Welcome to Python Automated Tool')

# Function to get the worksheet and workbook from the user
def Requirements():
    Path = input(
        '\nEnter the Location or PATH of the file (e.g: /home/carl/workloads.xlsx): ')
    Choice_sheet = input(
        '\nEnter the name of the sheet you want to access (e.g: sheet1): ')

    wb = load_workbook(Path)
    ws = wb[Choice_sheet]

    return ws, wb, Path

# Function to provide autocompletion for user input
def autocomplete_input(prompt_text, options):
    completer = WordCompleter(options, ignore_case=True)
    history = InMemoryHistory()

    print(f'\n{prompt_text}\n')

    while True:
        user_input = prompt("Search anything (If cannot find what you searching for. You can you alternative way. Enter none, to put it by yourself and contact the Developer.): ", completer=completer, history=history)
        if user_input.lower() == 'none':
            selected_option = input('Enter the text: ')
            return selected_option
        else:
            suggestions = [option for option in options if user_input.lower() in option.lower()]
            if suggestions:
                terminal_menu = dumb_menu.get_menu_choice(suggestions)
                menu_entry_index = terminal_menu
                selected_option = suggestions[menu_entry_index]
                print('>',selected_option)
                return selected_option
            else:
                print(f"\033[91mInvalid input. Please try again.\033[00m")

# Function to get the company name with autocompletion
def CompanyNameD():
    options = [
        "AFTERSIX DESSET&COFFEE INC", "AMICI FOODSERVICE VENTURES INC", "ANTHEM SHOPPES INC", "ASHLAR SECURITY SERVICES", "BIR", "BISTRO AMERICANO CORP", "BREADTALK PHILIPPINES INC", "BU PARTNERS INC", "CAFE VIA MARE", "CAFFEFRANCE CORP.", "CARA MIA SALES AND SERVICES INC", "CDC RESOURCES, INC", "CFAL OASIS DEVELOPMENT CORPORATION", "CK OF ASIA INC", "COFFEE MASTERS INC", "DECATHLON PHILIPPINES INC", "ENCHANTED KONGDOM INC", "ESPRESSO CONCEPTS INC", "EST. NAMNAM, INC", "EXA FUEL MARKETING OPC", "FOOGY MOUNTAIN COOKHOUSE", "GENMAICHA TEA CORPORATION", "GINGERSNAPS MARKETING IL CONIGLIO BIANCO CORPORATION", "GOLDEN DONUTS INC", "GOLDEN DONUTS, INC", "GOLDEN TEA LEAF INC", "GT & T GASOLINE STATION", "GTI", "HAPPY LEMON GROUP PHILIPPINES INC", "HEALTHY OPTION CORPORATION", "HOLDINGS INC.", "ILAOLLAO", "INTERBIKE MARKETING CORPORATION", "INTERNATIONAL TOYWORLD, INC", "IPPUDO PHILIPPINES, INC", "IT'S TIME FOR TIMS COFFEE INC.", "JBPX FOODS INC.", "JOLLIBEE BAGUIO BONIFACIO NIDEL COUNTRY LINE FOODS CORPORATION", "JOLLIBEE FOODS CORP.", "JOLLIBEE FOODS CORPORATION", "JOLLY888 FOOD CORP.", "JWB GASULINA CORP.", "KAREILA MGNT CORP", "LE PATISSIER INC",
        "LOLOMBOY WATERWORKS ASSOCIATION, INC", "LUGGAGE TRADING SERVICES INC", "MAJOR SHOPPING MANAGEMENT CORPORATION", "MARIE GALANTE FOOD SERVICES INC.", "MARY GRACE FOODS INC", "MAX'S GROUP INC", "MCDELIGHT SOUTH INC.", "MERCURY DRUG", "MERALCO", "MULTISPORTS INC", "MUNICIPALITY OF BOCAUE", "NATIONAL BOOK STORE, INC", "NOODLERAMA GROUP INC", "NOODLERAMA GROUP INC.", "OMAKASE INC.", "ORTIGAS HOME DEPOT", "OUTLOOK RESIDENCES INC", "PEPER BROERS, INC", "PRETIOLAS PHILIPPINES, INC", "REBEL ALLIANCE FOOD SUPPLY, INC", "REVSTAR TRADING", "RICHPRIME SALES INCORPORATED", "ROTI BOY, INC", "RUSTAN COFFEE CORP.", "RUSTAN COFFEE CORPORATION", "SAN ROQUE SHELL SERVICE STATION", "SEATTLE'S BEST COFFEE MASTER INC", "SPECIALTY FOOD RETAILERS INC", "SPECIALTY LIFESTYLE CONCEPTS, INC", "SUNDAYS EST CAFE AND RESTAURANT INC", "SUNDAYS EST CAFE AND RESTAURANT INC.", "SWITCHSTART INC", "TABLE BLACK KITCHEN CORP.", "TEAMBAKEMARKETING CORP", "THE FRENCH BAKER INC", "THE TWELVE 28TH INC", "TOBISTO FOODSO PHI INC", "TOBISTRO FOODS, INC", "UNIGLIPPLOBE TRAVELWARE CO., INC", "UNIMART", "URBAN TASTE INC", "VIVA INTERNATIONAL FOOD & RESTAURANT, INC", "WILD FLOUR BAKERY+CAFE CORP", "WILDFLOUR BAKERY+CAFE CORP"]

    return autocomplete_input("Company name:", options)

# Function to get the address with autocompletion
def AddressD():
    options = [
        "QUEZON CITY",
        "MAKATI CITY",
        "BAGUIO CITY",
        "GUIGUINTO BULACAN",
        "SANJUAN CITY",
        "PASIG CITY",
        "APALIT PAMPANGA",
        "MARIKINA CITY",
        "MARILAO BULACAN",
        "MARILAO EXTENTION OFFICE",
        "PLARIDEL BULACAN",
        "AYALA TRINOMA",
        "TAGUIG CITY",
        "MANDALUYONG CITY",
        "BOCAUE BULACAN",
        "SAN PEDRO LAGUNA",
        "ANTIPOLO RIZAL",
        "STA MARIA BULACAN",
        "LOLOMBOY BOCAUE BULACAN",
        "SANTAROSA LAGUNA",
        "ANTIPOLO CITY"
    ]

    return autocomplete_input("Address:", options)

# Function to get the description with autocompletion
def Description():
    options = ['MEDICINE', 'GASOLINE', 'MEALS', 'DRINKS', 'SNACKS', 'SUPPLIES',
               'SECURITY FEE', 'WATER', 'ELECTRICITY', '2550- Q', '2307', 'OFFICE SUPPLIES']

    return autocomplete_input("Description:", options)

# Function to convert input date string to a formatted date string
def DateTime(input_str):
    input_str = input_str.replace('.', '')
    month_number = int(input_str[:2])
    day = int(input_str[2:4])
    year = int(input_str[4:])
    if year >= 0 and year <= 21:
        year += 2000
    else:
        year += 1900
    year %= 100  # Get only the last two digits
    month_full_name = datetime.date(year, month_number, day).strftime("%B")
    month_abbr = datetime.date(year, month_number, day).strftime("%b")

    datetime_number = f"{day}-{month_abbr}-{year}"
    return datetime_number, input_str

# Function to get the TIN (Taxpayer Identification Number)
def TIN_D():
    while True:
        tin_number = input("\nEnter TIN number: ")
        if tin_number == "":
            print("\n\033[91mFill the blank!\033[00m")
        else:
            TIN = tin_number
            return TIN

# Function to input details and calculate Cost of Products & Services
def Input_details(input_str):
    datetime_number = DateTime(input_str)
    company_name = CompanyNameD()  
    address_name = AddressD()
    TIN = TIN_D()
    description_name = Description()
    TotalAmountPaid = float(input('\nEnter the Total Amount Paid: '))
    InputVat = float(input('\nEnter the Input Vat: '))
    CPS = TotalAmountPaid - InputVat

    return datetime_number, company_name, address_name, TIN, description_name, TotalAmountPaid, InputVat, CPS, input_str

# Function to add details to Excel worksheet
def excel(ws, wb, datetime_number, company_name, address_name, TIN, description_name, TotalAmountPaid, InputVat, CPS, Path):

    font = Font(name="Arial", size="10",)
    alignment = Alignment(horizontal='center', vertical='center')

    ws.append([datetime_number, company_name, address_name, TIN,
               description_name, TotalAmountPaid, InputVat, CPS])

    for row in ws.iter_rows(min_row=ws.max_row, max_row=ws.max_row):
        for cell in row:
            cell.font = font
            cell.alignment = alignment
    wb.save(Path)

# Function to confirm details with the user and encode them into Excel
def Confirm_details(datetime_number, company_name, address_name, TIN, description_name, TotalAmountPaid, InputVat, CPS, ws, wb, Path, input_str):
    print('\nConfirm Details\n')

    print(f'Date: {datetime_number}')
    print(f'Company_name: {company_name}')
    print(f'Address: {address_name}')
    print(f'TIN: {TIN}')
    print(f'Description: {description_name}')
    print(f'TotalAmountPaid: {TotalAmountPaid}')
    print(f'InputVat: {InputVat}')
    print(f'Cost of Product & Services: {CPS}')

    response = input(
        '\nPlease enter either y or n. If details are correct enter y, while for n if not: ')

    while response.lower() not in ['y', 'n']:
        print("\033[91mInvalid input. Enter either y or n.\033[0m")
        response = input("Check details. Do you want to proceed? (y/n): ")

    if response.lower() == 'y':
        excel(ws, wb, datetime_number, company_name, address_name, TIN,
              description_name, TotalAmountPaid, InputVat, CPS, Path)

        print('\nEncoded successfully')
    else:
        print('\nRetyping...')
        Input_details(input_str)

# Function to ask if user wants to input more data
def AnotherOne(input_str):
    response = input(
        '\nDo you want to create more? Enter y for YES, n for NO:  ')

    while response.lower() not in ['y', 'n']:
        print("\033[91mInvalid input. Enter either y or n.\033[0m")
        response = input(" Do you want to proceed? (y/n): ")

    if response.lower() == 'y':
        Input_details(input_str)

    else:
        print('Good Bye! ')
        exit()

# Function to execute the encoding process
def Method():
    input_str = input("\nEnter the date in MMDDYY format (e.g. 051523): ")
    datetime_number, company_name, address_name, TIN, description_name, TotalAmountPaid, InputVat, CPS, input_str = Input_details(input_str)
    
    Confirm_details(datetime_number[0], company_name, address_name, TIN, description_name, TotalAmountPaid, InputVat, CPS, ws, wb, Path, input_str)
    return input_str

if __name__ == "__main__":

    ws, wb, Path = Requirements()
    input_str = Method()

    AnotherOne(input_str)
