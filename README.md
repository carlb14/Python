import os
import re

#region windows
def get_menu_choice1(options):
    shortcuts = scan_short_cuts(options) # scan for shortcuts
    selected_index = 0
    while True:
        os.system("cls" if os.name == "nt" else "clear") #clear screen
        show_menu(options,selected_index)

        key = b'\a' #default to error
        try:
            import msvcrt
            char = msvcrt.getch() #get keypress
        except :
            pass

        if key == b'\x1b':  # Esc key to exit
            return -1
        elif key == b'\r':  # Enter key to select
            return selected_index
        elif key in (b'\x48', b'\x50'):  # Up or Down arrow
            selected_index = (selected_index + (1 if key == b'\x50' else -1) + len(options)) % len(options)
        elif key in shortcuts: # Shortcut key
            return shortcuts[key]
        elif key == b'\a':
            print('error , may not support your system')
            exit()
# endregion
def get_key(): #get keypress using getch, msvcrt = windows
    flag_have_getch = False
    flag_have_msvcrt = False
    try :
        import getch
        flag_have_getch = True
        first_char = getch.getch()
        if first_char == '\x1b': #arrow keys
            a=getch.getch()
            b=getch.getch()
            return {'[A': 'up', '[B': 'down', '[C': 'right', '[D': 'left' }[a+b]
        if ord(first_char) == 10:
            return  'enter'
        if ord(first_char) == 32:
            return  'space'
        else:
            return first_char #normal keys like abcd 1234
    except :
        pass
    
    try:
        import msvcrt
        flag_have_msvcrt = True
        key = msvcrt.getch()  # get keypress
        if key == b'\x1b':  # Esc key to exit
            return 'esc'
        elif key == b'\r':  # Enter key to select
            return 'enter'
        elif key == b'\x48':  # Up or Down arrow
            return  'up'
        elif key == b'\x50':  # Up or Down arrow
            return 'down'
        else:
            return key.decode('utf-8')
    except:
        pass
    
    if flag_have_getch == False and flag_have_msvcrt == False:
        print('\nErr:\tcan\'t get input \nFix:\tpip install getch')
        exit()

# isclean gives a clean meun without hint text
# give_key_str gives the key pressed instead of the index
def get_menu_choice(options,isclean = False,give_key_str = False):

    selected_index = 0

    while True:

        key = get_key()
        if key == 'enter':  # Enter key to select
            return selected_index
        elif key in ('up','down'):  # Up or Down arrow
            selected_index = (selected_index + (1 if key == 'down' else -1) + len(options)) % len(options)
        elif key in shortcuts:  # Shortcut key
            show_menu(options, shortcuts[key],isclean) #show selected option when using shortcut
            if(give_key_str):
                return key
            else:
                return shortcuts[key]




