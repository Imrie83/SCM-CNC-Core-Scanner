import pyodbc
import os
import sys
import re
import logging
from time import sleep, strftime, localtime, asctime
from datetime import datetime
from pywinauto.application import Application
import pywinauto.mouse as mouse
import pywinauto.keyboard as keyboard
import PySimpleGUI as sg


def except_log(err):
    logging.basicConfig(filename='myLog.log', level=logging.DEBUG,
                        format='%(asctime)s %(levelname)s %(name)s %(message)s')
    logger = logging.getLogger(__name__)
    logger.error(err)


def action_log(action):
    with open('action_log.txt', 'a+') as file:
        file.write(action)
    file.close()


def format_order(order):
    pattern = r'^([joswJOWS]{1})[a-zA-Z]*?([\d]*)-([\d]{1,})$'
    result = ''
    try:
        search = re.search(pattern, order)
        if search is None:
            result = 'error'
            pass
        elif search[1].upper() == 'J':
            result = 'JOB'+search[2].rjust(7, '0')+'-'+search[3].rjust(4, '0')
        elif search[1].upper() == 'O':
            result = 'ORD'+search[2].rjust(7, '0')+'-'+search[3].rjust(4, '0')
        elif search[1].upper() == 'S':
            result = 'SLB'+search[2].rjust(7, '0')+'-'+search[3].rjust(4, '0')
        elif search[1].upper() == 'W':
            result = 'WAR'+search[2].rjust(7, '0')+'-'+search[3].rjust(4, '0')
        return result
    except Exception as err:
        except_log(err)
        pass


def coordinates_read():
    coord_pattern = r'(^#[a-z]{2,10}#)x\ ([0-9]{1,5})\ y\ ([0-9]{1,5})'
    x_width = 0
    y_width = 0
    x_height = 0
    y_height = 0
    x_tp = 0
    y_tp = 0
    x_mul = 0
    y_mul = 0
    x_ok = 0
    y_ok = 0

    try:
        # open text file with coordinates for 'mfcgridctrl' fields to allow
        # for input of data in to non standard win32 control.
        with open('coordinates.txt', 'r') as f:
            reader = f.readlines()
            for item in reader:
                find_coord = re.search(coord_pattern, item)
                # allocating coordinates to variables
                if find_coord[1] == '#width#':
                    x_width = int(find_coord[2].strip())
                    y_width = int(find_coord[3].strip())
                elif find_coord[1] == '#height#':
                    x_height = int(find_coord[2].strip())
                    y_height = int(find_coord[3].strip())
                elif find_coord[1] == '#tp#':
                    x_tp = int(find_coord[2].strip())
                    y_tp = int(find_coord[3].strip())
                elif find_coord[1] == '#mullaley#':
                    x_mul = int(find_coord[2].strip())
                    y_mul = int(find_coord[3].strip())
                elif find_coord[1] == '#ok#':
                    x_ok = int(find_coord[2].strip())
                    y_ok = int(find_coord[3].strip())

        return x_width, y_width, x_height, y_height, x_tp, y_tp, x_mul,\
            y_mul, x_ok, y_ok

    except Exception as err:
        except_log(err)
        pass


def win_pos():

    coord = r'x\ ([0-9]{1,8})\ y\ ([0-9]{1,8})\ width\ ([0-9]{1,4})\ height\ ([0-9]{1,4})'

    try:
        with open('pos.txt', 'r') as f:
            reader = f.readlines()
            for item in reader:
                find_coord = re.search(coord, item)
                x_pos = int(find_coord[1].strip())
                y_pos = int(find_coord[2].strip())
                win_width = int(find_coord[3].strip())
                win_height = int(find_coord[4].strip())
        return x_pos, y_pos, win_width, win_height

    except Exception as err:
        except_log(err)
        pass


def file_path():

    pattern = r'(^#[a-z0-9_]{5,16}#)([A-Z]\:\\\(?\)?.*)'
    core_path = ''
    c3_panel = ''
    c4_panel = ''
    c5_panel = ''
    c6_panel = ''
    flush = ''
    c3_panel_ce = ''
    c4_panel_ce = ''
    c5_panel_ce = ''
    c6_panel_ce = ''
    flush_ce = ''
    c3_panel_ce_large = ''
    c4_panel_ce_large = ''
    c5_panel_ce_large = ''
    c6_panel_ce_large = ''
    flush_ce_large = ''

    try:
        # open text file with paths to the cores programmes
        # to be used in the script
        # allows to change the paths with no need to edit code
        with open('paths.txt', 'r') as f:
            reader = f.readlines()
            for item in reader:
                find = re.search(pattern, item)
                # allocate programmes paths to variables
                if find[1] == '#core_path#':
                    core_path = find[2]
                elif find[1] == '#3_panel#':
                    c3_panel = find[2]
                elif find[1] == '#4_panel#':
                    c4_panel = find[2]
                elif find[1] == '#5_panel#':
                    c5_panel = find[2]
                elif find[1] == '#6_panel#':
                    c6_panel = find[2]
                elif find[1] == '#flush#':
                    flush = find[2]
                elif find[1] == '#3_panel_ce#':
                    c3_panel_ce = find[2]
                elif find[1] == '#4_panel_ce#':
                    c4_panel_ce = find[2]
                elif find[1] == '#5_panel_ce#':
                    c5_panel_ce = find[2]
                elif find[1] == '#6_panel_ce#':
                    c6_panel_ce = find[2]
                elif find[1] == '#flush_ce#':
                    flush_ce = find[2]
                elif find[1] == '#3_panel_ce_large#':
                    c3_panel_ce_large = find[2]
                elif find[1] == '#4_panel_ce_large#':
                    c4_panel_ce_large = find[2]
                elif find[1] == '#5_panel_ce_large#':
                    c5_panel_ce_large = find[2]
                elif find[1] == '#6_panel_ce_large#':
                    c6_panel_ce_large = find[2]
                elif find[1] == '#flush_ce_large#':
                    flush_ce_large = find[2]

        return core_path, c3_panel, c4_panel, c5_panel, c6_panel, flush, \
            c3_panel_ce, c4_panel_ce, c5_panel_ce, c6_panel_ce, flush_ce, \
            c3_panel_ce_large, c4_panel_ce_large, c5_panel_ce_large, \
            c6_panel_ce_large, flush_ce_large
    except Exception as err:
        except_log(err)
        pass


def auto_load(order, manual_ce, manual_std, manual_mul):
    err = 0
    width, height, style, tp, hinge_no, error, formated, specifier, timber =\
        data_retrieval(order)
    core_path, c3_panel, c4_panel, c5_panel, c6_panel, flush, c3_panel_ce, \
        c4_panel_ce, c5_panel_ce, c6_panel_ce, flush_ce, c3_panel_ce_large, \
        c4_panel_ce_large, c5_panel_ce_large, c6_panel_ce_large, \
        flush_ce_large = file_path()
    x_width, y_width, x_height, y_height, x_tp, y_tp, x_mul, y_mul, \
        x_ok, y_ok = coordinates_read()
    ce = 0
    if order == '':
        pass
    elif error == '1':
        pass
    else:
        app = Application().connect(title_re='.*Machine*')
        # this_app = Application().connect(title_re='.*Cores Scanner*')
        # create dialogs for main window, file selection window, header
        # and prog type window
        main_dlg = app.window(title_re='.*Machine*')
        select_dlg = app.window(title_re='^Select*')
        header_dlg = app.window(title_re='.*Change header*')
        prog_dlg = app.window(title_re='^Select automatic*')
        # this_app_dlg = this_app.window(title_re='Cores Scanner')

        try:
            # try pressing shift+F8 to load the programme - works only on
            # first running the machine!
            main_dlg.wait('enabled', timeout=0.1).type_keys('+{VK_F8}')
            prog_dlg.wait('enabled', timeout=0.1).type_keys('{ENTER}')
        except:
            pass

        try:
            # press shift+f2 main dialog to open file selection window
            main_dlg.wait('enabled', timeout=0.1).type_keys('+{VK_F2}')
        except Exception as err:
            except_log(err)
            pass
        try:
            # in file selection window type file name followed by ENTER to
            # open scanned file
            # "style" numbers are hard coded in to WinMan DB
            if style == '681' or style == '683' or style == '691':
                if ("SMC REGENCY EN" not in specifier and manual_ce is False and
                        timber != 43) or manual_std is True:
                        select_dlg.Edit.wait('enabled', timeout=1)\
                            .type_keys(c5_panel+'{ENTER}', with_spaces=True)
                elif ("SMC REGENCY EN" in specifier and manual_std is False) \
                        or manual_ce is True or timber == 43:
                    ce = 1
                    if int(width < 849) and int(height < 2067):
                        select_dlg.Edit.wait('enabled', timeout=1)\
                            .type_keys(c5_panel_ce+'{ENTER}',
                                       with_spaces=True)
                    else:
                        select_dlg.Edit.wait('enabled', timeout=1) \
                            .type_keys(c5_panel_ce_large+'{ENTER}',
                                       with_spaces=True)
                        sg.popup('Use Large Cores', keep_on_top='true')
                else:
                    pass
                header_dlg.Parameters.click()
                mouse.click(coords=(x_width, y_width))
                keyboard.send_keys(str(width))
                mouse.click(coords=(x_height, y_height))
                keyboard.send_keys(str(height))
                mouse.click(coords=(x_tp, y_tp))
                keyboard.send_keys(tp)
                if manual_mul is True:
                    mouse.click(coords=(x_mul, y_mul))
                    keyboard.send_keys('1')
            elif style == '686':
                ce = 1
                if ("SMC REGENCY EN" not in specifier and manual_ce is False and
                        timber != 43) or manual_std is True:
                    select_dlg.Edit.wait('enabled', timeout=1)\
                        .type_keys(c4_panel+'{ENTER}', with_spaces=True)
                elif ("SMC REGENCY EN" in specifier and manual_std is False)\
                        or manual_ce is True or timber == 43:
                    if int(width < 849) and int(height < 2067):
                        select_dlg.Edit.wait('enabled', timeout=1)\
                            .type_keys(c4_panel_ce+'{ENTER}',
                                       with_spaces=True)
                    else:
                        select_dlg.Edit.wait('enabled', timeout=1)\
                            .type_keys(c4_panel_ce_large+'{ENTER}',
                                       with_spaces=True)
                        sg.popup('Use Large Cores', keep_on_top='true')
                else:
                    pass
                header_dlg.Parameters.click()
                mouse.click(coords=(x_width, y_width))
                keyboard.send_keys(str(width))
                mouse.click(coords=(x_height, y_height))
                keyboard.send_keys(str(height))
                mouse.click(coords=(x_tp, y_tp))
                keyboard.send_keys(tp)
                if manual_mul is True:
                    mouse.click(coords=(x_mul, y_mul))
                    keyboard.send_keys('1')

            elif style == '684' or style == '685' or style == '699':
                if ("SMC REGENCY EN" not in specifier and manual_ce is False and
                        timber != 43) or manual_std is True:
                    ce = 1
                    select_dlg.Edit.wait('enabled', timeout=1)\
                        .type_keys(c3_panel+'{ENTER}', with_spaces=True)
                elif ("SMC REGENCY EN" in specifier and manual_std is False)\
                        or manual_ce is True or timber == 43:
                    if int(width < 849) and int(height < 2067):
                        select_dlg.Edit.wait('enabled', timeout=1)\
                            .type_keys(c3_panel_ce+'{ENTER}',
                                       with_spaces=True)
                    else:
                        select_dlg.Edit.wait('enabled', timeout=1)\
                            .type_keys(c3_panel_ce_large+'{ENTER}',
                                       with_spaces=True)
                        sg.popup('Use Large Cores', keep_on_top='true')
                else:
                    pass
                header_dlg.Parameters.click()
                mouse.click(coords=(x_width, y_width))
                keyboard.send_keys(str(width))
                mouse.click(coords=(x_height, y_height))
                keyboard.send_keys(str(height))
                mouse.click(coords=(x_tp, y_tp))
                keyboard.send_keys(tp)
                if manual_mul is True:
                    mouse.click(coords=(x_mul, y_mul))
                    keyboard.send_keys('1')

            elif style == '688' or style == '687':
                if ("SMC REGENCY EN" not in specifier and manual_ce is False and
                        timber != 43) or manual_std is True:
                    ce = 1
                    select_dlg.Edit.wait('enabled', timeout=1)\
                        .type_keys(c6_panel+'{ENTER}', with_spaces=True)
                elif ("SMC REGENCY EN" in specifier and manual_std is False)\
                        or manual_ce is True or timber == 43:
                    if int(width < 849) and int(height < 2067):
                        select_dlg.Edit.wait('enabled', timeout=1)\
                            .type_keys(c6_panel_ce+'{ENTER}',
                                       with_spaces=True)
                    else:
                        select_dlg.Edit.wait('enabled', timeout=1)\
                            .type_keys(c6_panel_ce_large+'{ENTER}',
                                       with_spaces=True)
                        sg.popup('Use Large Cores', keep_on_top='true')
                else:
                    pass
                header_dlg.Parameters.click()
                mouse.click(coords=(x_width, y_width))
                keyboard.send_keys(str(width))
                mouse.click(coords=(x_height, y_height))
                keyboard.send_keys(str(height))
                mouse.click(coords=(x_tp, y_tp))
                keyboard.send_keys(tp)
                if manual_mul is True:
                    mouse.click(coords=(x_mul, y_mul))
                    keyboard.send_keys('1')

            elif style == '529' or style == '629':
                if ("SMC REGENCY EN" not in specifier and manual_ce is False and
                        timber != 43) or manual_std is True:
                    ce = 1
                    select_dlg.Edit.wait('enabled', timeout=1)\
                        .type_keys(flush+'{ENTER}', with_spaces=True)
                elif ("SMC REGENCY EN" in specifier and manual_std is False)\
                        or manual_ce is True or timber == 43:
                    if int(width < 849) and int(height < 2067):
                        select_dlg.Edit.wait('enabled', timeout=1)\
                            .type_keys(flush_ce+'{ENTER}', with_spaces=True)
                    else:
                        select_dlg.Edit.wait('enabled', timeout=1)\
                            .type_keys(flush_ce_large+'{ENTER}',
                                       with_spaces=True)
                        sg.popup('Use Large Cores', keep_on_top='true')
                else:
                    pass
                print(timber)
                header_dlg.Parameters.click()
                mouse.click(coords=(x_width, y_width))
                keyboard.send_keys(str(height))
                mouse.click(coords=(x_height, y_height))
                keyboard.send_keys(str(width))

            return width, height, style, tp, hinge_no, ce, error, formated
        except Exception as err:
            except_log(err)
            pass


def data_retrieval(order):
    width = 0
    height = 0
    style = 0
    tp = 0
    hinge_no = 0
    error = 0
    formated = ''
    specifier = ''
    writeToDB = ''
    timber = 0

    try:
        time_now = strftime("%Y-%m-%d %H:%M", localtime())
        writeToDB = time_now+":10"
        conn = pyodbc.connect('Driver={SQL Server};'
                              'Server=DIV-SQL\WINMAN;'
                              'Database=WinManLive;'
                              'uid=winman;pwd=winman')
        cursor = conn.cursor()

        if order == '':
            order = ''
            pass
        else:
            order = format_order(order)
            formated = order
            # order = 'JOB0157520-0001'
            cursor.execute("SELECT Bespoke_DoorEntry.DoorLeafWidth,\
                Bespoke_DoorEntry.DoorLeafHeight,Bespoke_DoorEntry.CNCDoorStyle,\
                Bespoke_DoorEntry.ProductionComments,Bespoke_DoorEntry.Comments,\
                Bespoke_DoorEntry.CNCTopPoly,Bespoke_DoorEntry.HingeQuantity,\
                Bespoke_DoorEntry.Specifier, Bespoke_DoorEntry.TimberType FROM\
                Bespoke_DoorEntry, SalesOrderItems WHERE \
                Bespoke_DoorEntry.SalesOrderItem = \
                SalesOrderItems.SalesOrderItem AND CNCOrderNumber LIKE (?)",
                    (order))
            data = cursor.fetchone()
            if data is None:
                error = 1
                pass
            else:
                width = int(data[0])
                height = int(data[1])
                style = data[2]
                tp = str(data[5])
                hinge_no = int(data[6])
                specifier = str(data[7].upper())
                timber = int(data[8])
            cursor.close()
            #
            # cursor = conn.cursor()
            # cursor.execute("UPDATE dbo.Bespoke_DoorEntry SET CNCScanTime=(?)
            #   WHERE CNCOrderNumber=(?)", (writeToDB,order))
            # conn.commit()
            # cursor.close()
        return width, height, style, tp, hinge_no, error, formated, specifier,\
            timber

    except Exception as err:
        except_log(err)
        sg.popup("Check you network connection!")
        pass


def main_window():

    x_pos, y_pos, win_width, win_height = win_pos()

    # sg.theme('Default1')
    # sg.theme('BrownBlue')
    sg.theme('SystemDefaultForReal')
    # All the stuff inside the window.
    layout = [[sg.Frame(layout=[
                [sg.Text('Enter No.'), sg.InputText(size=(16, 1),
                                                    key="_INPUT_")],
                [sg.T()],
                [sg.Text('Job No:', size=(10, 1)),
                    sg.Text(key='_JOB_', size=(16, 1), justification='right')],
                [sg.Text('Width: ', size=(10, 1)),
                    sg.Text(key='_WIDTH_', size=(16, 1),
                            justification='right')],
                [sg.Text('Height:', size=(10, 1)),
                    sg.Text(key='_HEIGHT_', size=(16, 1),
                            justification='right')],
                [sg.Text('Top Poly:', size=(10, 1)),
                    sg.Text(key='_TP_', size=(16, 1),
                            justification='right')],
                [sg.Checkbox("CE", key='_CE_', enable_events=True),
                    sg.Checkbox("STD", key='_STD_',  enable_events=True),
                    sg.Checkbox("Mullaley", key='_MUL_', enable_events=True)],
                [sg.T(key='_ERROR_', size=(22, 1), justification='centre')],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Text('')],
                [sg.Button('>>GO<<', bind_return_key=True, visible=False)]],
                title='Cores Scan', title_color='black',
                relief=sg.RELIEF_GROOVE)]]

    # Create the Window
    window = sg.Window('Cores Scanner', layout, keep_on_top='true',
                       location=(x_pos, y_pos), finalize='true',
                       size=(win_width, win_height), no_titlebar=False,
                       resizable=True, grab_anywhere=False, alpha_channel=1,
                       return_keyboard_events=True)

    # Event Loop to process "events" and get the "values" of the inputs
    while True:
        event, values = window.read()
        # if user closes window or clicks cancel
        if event == sg.WIN_CLOSED or event == 'Quit':
            break

        if event == "_STD_":
            try:
                now = strftime("%d/%m/%Y %H:%M:%S", localtime())
                action = "***\n" + now + "\nCheckbox 'STD' status changed to:"\
                    + str(values["_STD_"]) + "\n***\n\n"
                action_log(action)
                if values["_STD_"] is True:
                    window.FindElement('_CE_').Update(False)
                    window.FindElement('_MUL_').Update(False)
            except Exception as err:
                except_log(err)

        if event == "_CE_":
            try:
                now = strftime("%d/%m/%Y %H:%M:%S", localtime())
                action = "***\n" + now + "\nCheckbox 'CE' status changed to:"\
                    + str(values["_CE_"]) + "\n***\n\n"
                action_log(action)
                if values["_CE_"] is True:
                    window.FindElement('_STD_').Update(False)
                    window.FindElement('_MUL_').Update(False)
            except Exception as err:
                except_log(err)

        if event == "_MUL_":
            try:
                now = strftime("%d/%m/%Y %H:%M:%S", localtime())
                action = "***\n" + now + "\nCheckbox 'Mullaley' status changed to:"\
                    + str(values["_MUL_"]) + "\n***\n\n"
                action_log(action)
                if values["_MUL_"] is True:
                    window.FindElement('_CE_').Update(False)
                    window.FindElement('_STD_').Update(False)
            except Exception as err:
                except_log(err)

        if event == '>>GO<<':
            order = values["_INPUT_"]
            manual_ce = values["_CE_"]
            manual_std = values["_STD_"]
            manual_mul = values["_MUL_"]

            if order == '':
                window.FindElement('_ERROR_').Update('Enter a Job Number')
                pass
            else:
                try:
                    width, height, style, tp, hinge_no, ce, error, formated =\
                        auto_load(order, manual_ce, manual_std, manual_mul)

                    window.FindElement('_JOB_').Update(formated)
                    window.FindElement('_WIDTH_').Update(width)
                    window.FindElement('_HEIGHT_').Update(height)
                    window.FindElement('_TP_').Update(tp)
                    window.FindElement('_ERROR_').Update('')
                    window.FindElement('_INPUT_').Update('')

                    try:
                        now = strftime("%d/%m/%Y %H:%M:%S", localtime())
                        action = "***\n" + now + "\nSlab scanned:" +\
                            str(formated) + "\nCE box:" + str(manual_ce) +\
                            "\nSTD box:" + str(manual_std) + "\nMullaley box:"\
                            + str(manual_mul) + "\n***\n\n"
                        action_log(action)
                    except Exception as err:
                        except_log(err)

                    if error == 1:
                        window.FindElement('_ERROR_')\
                            .Update('Incorrect Job Number')
                except Exception as err:
                    except_log(err)

                    pass
    window.close()


def main():
    try:
        main_window()

    except Exception as err:
        except_log(err)
        pass

if __name__ == '__main__':
    main()
