#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""

Partner Promotions Automation Bot
Glovo Italy

"""

import json
import pandas as pd
import requests
import logging
import sys
import os
import get_new_token
from datetime import datetime, timedelta
from time import sleep
from concurrent.futures import ThreadPoolExecutor
from subprocess import call
#from regex import sub

class Promobot:
    global bot_name, mode, df_promo, output_excel, output_path, platform
    bot_name = 'promobot'
    '''Init functions'''
    #Step 1: set path
    def set_path():
        global cwd, token_path, input_path, platform
        cwd = os.getcwd()
        '''
        #if sys has attribute _MEIPASS then script launched by bundled exe.
        if getattr(sys, '_MEIPASS', False):
            cwd = os.path.dirname(os.path.dirname(sys._MEIPASS))
        else:
            cwd = os.getcwd()
        else:
            if "SPY_PYTHONPATH" in os.environ:
                cwd = os.getcwd()
            else:
                cwd = os.path.dirname(sys.path[0])
        print('cwd',cwd)
        print('sys._MEIPASS',sys._MEIPASS)
        print('sys.path[0]',sys.path[0])
        print('os.getcwd()',os.getcwd())
        '''
        token_path = os.path.join(cwd,'my_personal_token.json')
        if os.name == 'nt':
            platform = 'windows'
        elif os.name == 'posix':
            platform = 'mac'

    #Step 2: enable Logger
    def logger_start():
        #log config
        logging.basicConfig(filename = os.path.join(cwd,"my_log.log"),
                            level =  logging.INFO,
                            format = "%(levelname)s %(asctime)s %(message)s",
                            datefmt = '%m/%d/%Y %H:%M:%S',
                            filemode = "a")
        #log start
        global logger
        logger = logging.getLogger()
        logger.info(f"Starting log for {bot_name}")
        #print("Logger started")
        if platform == 'mac':
            call(["chflags", "hidden", os.path.join(cwd,"my_log.log")])
            if os.path.isdir('__pycache__'): call(["chflags", "hidden", os.path.join(cwd,"__pycache__")])
        elif platform == 'windows':
            call(["attrib", "+h", os.path.join(cwd,"my_log.log")])
            if os.path.isdir('__pycache__'): call(["attrib", "+h", os.path.join(cwd,"__pycache__")])
    #custom for Step 3: read credentials json
    def read_json():
        global content
        global glovo_email, refresh_token, country
        with open(token_path) as read_file:
            content = json.load(read_file)
        glovo_email = content['glovo_email']
        refresh_token = content['refresh_token']
        country = content['country']

    #Step 3: check login credentials
    def login_check():
        #Check/get login data: check if file 'my personal token' exists and read it to get login data.
        global glovo_email, refresh_token
        #print("Checking login data")
        if os.path.isfile(token_path):
            try:
                Promobot.read_json()
            except Exception:
                    get_new_token.Glovo_token()
            else:
                welcome_name = glovo_email[:glovo_email.find("@")].replace("."," ").title()
                print(f"\nLogged in as {welcome_name}")
        #if file does not exist: lauch file creation
        else:
            get_new_token.Glovo_token()

    #Step 4: get fresh api access token
    def refresh():
        global oauth, access_token
        Promobot.read_json()
        #step 2: make request at oauth/refresh
        oauth_data = {'refreshToken' : refresh_token, 'grantType' : 'refresh_token'}
        oauth_request = requests.post('https://adminapi.glovoapp.com/oauth/refresh', json = oauth_data)
        #print(oauth_request.ok)
        if oauth_request.ok:
            access_token = oauth_request.json()['accessToken']
            new_refresh_token = oauth_request.json()['refreshToken']
            oauth = {'authorization' : access_token}
            #print("Token refreshed")
            logger.info('Access Token Refreshed')
            #saving new refresh token
            content['refresh_token'] = new_refresh_token
            with open(token_path, "r+") as dst_file:
                json.dump(content, dst_file)
            #print("token refreshed")
        else:
            print(f"Token NOT refreshed -> {oauth_request.content}")
            logger.info(f'Access Token NOT Refreshed -> {oauth_request.content}')


    def print_bot_name():
        print('\nEl Promobot')
        print('\nPartners Promotions Automation')

    '''''''''''''''''''''''''''''End Init'''''''''''''''''''''''''''''

    '''''''''''''''''''''''''''Beginning bot'''''''''''''''''''''''''''
    '''Part 1: Get Store Address IDs and set mode'''
    #Set mode(enable/disable)
    def set_mode():
        sleep(0.5)
        global mode
        while True:
            a_or_b = input('\nSelect the type of operation you want to perform:\n[A] - Create Promos\n[B] - Delete Promos\n[C] - Check current Promos status\nPress "A", "B" or "C" then press ENTER:\t').lower().strip()
            if a_or_b in ["a","b","c"]:
                sleep(0.5)
                if a_or_b == 'a':
                    print('\nSelected mode is "Create Promos"')
                    mode = 'create'
                    #print(f'Promos will will be Created for the {len(df_promo)} Store IDs found in {input_name}')
                    #time.sleep(2)
                    #confirm = input('\nProceed with Promo Creation? [yes/no]\t').strip().lower()
                    #if confirm in ['yes','ye','y','si']:
                        #mode = 'create'
                        #break
                elif a_or_b == "b":
                    print('\nSelected mode is "Delete Promos"')
                    mode = 'delete'
                    #print(f'Promos will will be Deleted for the {len(df_promo)} Store IDs found in {input_name}')
                    #time.sleep(2)
                    #confirm = input('\nProceed with Promo Deletion? [yes/no]\t').strip().lower()
                    #if confirm in ['yes','ye','y','si']:
                        #mode = 'delete'
                        #break
                elif a_or_b == "c":
                    print('\nSelected mode is "Promos status check"')
                    mode = 'check'
                    #print(f'A simple check of current Promo status will be done for the {len(df_promo)} Store IDs found in {input_name}')
                    #time.sleep(2)
                    #confirm = input('\nProceed [yes]/[no]:\t').strip().lower()
                    #if confirm in ["yes","y","ye","si"]:
                        #mode = 'check'
                        #break
                break

    #custom function for set_input(): extracts dataframe once input name is set
    def import_data(input_file):
        global df_promo, no_prime, no_products, no_store_address
        #import data
        df_promo = pd.read_excel(input_file, usecols=lambda x: 'Unnamed' not in x, engine='openpyxl')
        #clean empty rows
        df_promo.dropna(how='all', inplace = True)
        #reset index after deleting empty rows
        df_promo.reset_index(drop = True, inplace = True)
        #check if Status column in present
        try: df_promo['Status']
        except KeyError: df_promo.loc[:,'Status']= None
        if mode == 'create':
            #clean str columns
            try:
                df_promo.loc[:,'City_Code'] = df_promo.loc[:,'City_Code'].str.strip()
                df_promo.loc[:,'Promo_Name'] = df_promo.loc[:,'Promo_Name'].str.strip()
            except AttributeError:pass
            try:
                df_promo.loc[:,'Promo_Type ("FLAT"/"FREE"/"XX%")'] = df_promo.loc[:,'Promo_Type ("FLAT"/"FREE"/"XX%")'].str.strip()
            except AttributeError:pass
            #Only Prime
            try:
                df_promo.loc[:,'Only_Prime'] = df_promo.loc[:,'Only_Prime'].str.strip()
            except KeyError:
                no_prime = True
            else:
                if pd.notna(df_promo.loc[:,'Only_Prime']).sum() == 0:
                    no_prime = True
                else:
                    no_prime = False
            #Pructs columns
            try:
                df_promo.loc[:,list(map(lambda x: 'Product' in x, list(df_promo)))]
            except KeyError:
                no_products = True
            else:
                if pd.notna(df_promo.loc[:,list(map(lambda x: 'Product' in x, list(df_promo)))]).sum().sum() ==  0:
                    no_products = True
                else:
                    no_products = False
            #Store_Address columns
            try:
                df_promo.loc[:,list(map(lambda x: 'Store_Address' in x, list(df_promo)))]
            except KeyError:
                no_store_address = True
            else:
                if pd.notna(df_promo.loc[:,list(map(lambda x: 'Store_Address' in x, list(df_promo)))]).sum().sum() ==  0:
                    no_store_address = True
                else:
                    no_store_address = False
            #clean dates
            df_promo.loc[:,"Start_Date (dd/mm/yyyy)"] = pd.to_datetime(df_promo.loc[:,"Start_Date (dd/mm/yyyy)"],dayfirst=True)
            df_promo.loc[:,"End_Date (included)"] = pd.to_datetime(df_promo.loc[:,"End_Date (included)"],dayfirst=True)
            #%glovo
            if df_promo.loc[:,"%GLOVO"].dtype == 'O':
                df_promo.loc[:,"%GLOVO"]= df_promo.loc[:,"%GLOVO"].str.strip('%')
                df_promo.loc[:,"%GLOVO"].astype('int')
        elif mode == 'delete':
            if 'Promo_ID' not in list(df_promo):
                raise KeyError('Promo_ID')
        elif mode == 'check':
            if 'Promo_ID' not in list(df_promo):
                raise KeyError('Promo_ID')
        #print(f'Data succesfully extracted from {os.path.join(os.path.basename(os.path.dirname(input_file)),os.path.basename(input_file))}')

    #custom function for set_input(): find excel file
    def find_excel_file_path(excel_name):
        #walk in cwd -> return excel path or raise error
        for root, dirs, files in os.walk(cwd):
            if excel_name in files:
                for file in files:
                    if file == excel_name:
                        #print(f'\n{excel_name} found in folder {os.path.basename(root)}')
                        return os.path.join(root,file)
        else:
            #print('File not found in current working directory')
            raise NameError

    def set_output_dir():
        global output_path
        upload_identifier = input('Enter folder name for storing results:\t')
        output_path = os.path.join(cwd, upload_identifier)
        try: os.mkdir(output_path)
        except Exception: pass

    #set input
    def set_input():
        global input_name
        while True:
            input_name = input('Enter Excel file name:\t')
            if '.xlsx' not in input_name: input_name = f'{input_name}.xlsx'
            #input_name = f'{bot_name}_input.xlsx'
            try:
                input_path = Promobot.find_excel_file_path(input_name)
            #print('cwd',cwd)
            #print('input_path',input_path)
            except NameError:
                sleep(0.5)
                print(f'\nCould not find {input_name} in {os.path.basename(cwd)}\nPlease try again\n')
                continue
            else:
                try:
                    Promobot.import_data(input_path)
                except KeyError as e:
                    print(f'Column {e} is missing. Unable to import data.\nUpdate file or choose another file.')
                    continue
                else:
                    #print(f'Promos will will be Created for the {len(df_promo)} Store IDs found in {input_name}')
                    Promobot.set_output_dir()
                    sleep(1)
                    confirm_path = input(f'Using file \'{input_name}\' to {mode} promos of {len(df_promo)} Store IDs found in {os.path.join(os.path.basename(os.path.dirname(input_path)),os.path.basename(input_path))}.\nOutput will be saved in folder \'{os.path.basename(output_path)}\'\nContinue? [yes,no]\t')
                    if confirm_path in ["yes","y","ye","si"]:
                        logger.info(f'Using file {input_name} in folder {os.path.basename(os.path.dirname(input_path))}')
                        print('')
                        break

    '''promo creation'''
    def p_type(promo_type):
        if type(promo_type) == 'str':
            promo_type.strip().upper()
        if promo_type == 'FLAT':
            return 'FLAT_DELIVERY'
        elif promo_type == 'FREE':
            return 'FREE_DELIVERY'
        else:
            return 'PERCENTAGE_DISCOUNT'

    def del_fee(promo_type):
        if promo_type == 'FLAT':
            return 100
        if promo_type == 'FREE':
            return None
        else:
            return None

    def perc(promo_type):
        if Promobot.p_type(promo_type) == 'FLAT_DELIVERY' or Promobot.p_type(promo_type) == 'FREE_DELIVERY':
            return None
        elif Promobot.p_type(promo_type) == 'PERCENTAGE_DISCOUNT':
            if type(promo_type) == 'str':
                return int((promo_type).strip('%'))
            else:
                return int(promo_type)

    def strat(subsidy):
        return f'ASSUMED_BY_{subsidy}'

    def paymentStrat(subsidy):
        if Promobot.strat(subsidy) == "ASSUMED_BY_GLOVO" or Promobot.strat(subsidy) == "ASSUMED_BY_PARTNER":
            return Promobot.strat(subsidy)
        elif Promobot.strat(subsidy) == "ASSUMED_BY_BOTH":
            return "ASSUMED_BY_PARTNER"

    def time_code(x, date):
        if x == 'start':
            hours_added = timedelta(hours = 1)
            future_date = date + hours_added
            stamp = datetime.timestamp(future_date)
            return int(stamp*1000)
        if x == 'end':
            hours_added = timedelta(hours = 25)
            future_date = date + hours_added
            stamp = datetime.timestamp(future_date)
            return int(stamp*1000)

    def products_ID_list(n):
        if no_products:
            return None
        else:
            prods_list = []
            for i in range(1,10):
                try:
                    df_promo.at[n,f'Product_ID{i}']
                except KeyError:
                    break
                else:
                    if pd.isna(df_promo.at[n,f'Product_ID{i}']): continue
                    prods_list.append((str(df_promo.at[n,f'Product_ID{i}'])).replace('\ufeff', ''))
            if prods_list == []:
                return None
            else:
                return prods_list

    def store_addresses_ID_list(n):
        if no_store_address:
            return None
        else:
            sa_ID_list = []
            for o in range(1,10):
                try:
                    df_promo.at[n,f'Store_Address{o}']
                except KeyError:
                    break
                else:
                    if pd.isna(df_promo.at[n,f'Store_Address{o}']): continue
                    if type(df_promo.at[n,f'Store_Address{o}']) == str:
                        df_promo.at[n,f'Store_Address{o}'] = df_promo.at[n,f'Store_Address{o}'].replace('\ufeff', '')

                    sa_ID_list.append(int(df_promo.at[n,f'Store_Address{o}']))
            if sa_ID_list == []:
                return None
            else:
                return sa_ID_list

    def subsidyValue(subject, n):
        if Promobot.strat((df_promo.at[n,'Subsidized_By (\"PARTNER\"/\"GLOVO\"/\"BOTH\")']).strip().upper()) == 'ASSUMED_BY_GLOVO':
                if subject == 'glovo':
                    return 100
                if subject == 'partner':
                    return 0
        elif Promobot.strat((df_promo.at[n,'Subsidized_By (\"PARTNER\"/\"GLOVO\"/\"BOTH\")']).strip().upper()) == 'ASSUMED_BY_PARTNER':
            if subject == 'glovo':
                return 0
            if subject == 'partner':
                return 100
        elif Promobot.strat((df_promo.at[n,'Subsidized_By (\"PARTNER\"/\"GLOVO\"/\"BOTH\")']).strip().upper()) == 'ASSUMED_BY_BOTH':
            if subject == 'glovo':
                return df_promo.at[n,"%GLOVO"]
            if subject == 'partner':
                return df_promo.at[n,"%PARTNER"]

    def is_prime(n):
        if no_prime:
            return None
        else:
            if df_promo.at[n,'Only_Prime'] == 'yes':
                return True
            else:
                return False

    def creation(n):
        if df_promo.at[n,'Status'] == 'created':
            print(n,'already created')
        else:
            url = 'https://adminapi.glovoapp.com/admin/partner_promotions'
            payload = {"name": df_promo.at[n,'Promo_Name'],
                        "cityCode": df_promo.at[n,'City_Code'],
                        "type": Promobot.p_type(df_promo.at[n,'Promo_Type ("FLAT"/"FREE"/"XX%")']),
                        "percentage": Promobot.perc(df_promo.at[n,'Promo_Type ("FLAT"/"FREE"/"XX%")']),
                        "deliveryFeeCents": Promobot.del_fee(df_promo.at[n,'Promo_Type ("FLAT"/"FREE"/"XX%")']),
                        "startDate": Promobot.time_code('start',df_promo.at[n,"Start_Date (dd/mm/yyyy)"]),
                        "endDate": Promobot.time_code('end',df_promo.at[n,"End_Date (included)"]),
                        "openingTimes": None,
                        "partners":[{"id": int(df_promo.at[n,'Store_ID']),
                                    "paymentStrategy": Promobot.paymentStrat((df_promo.at[n,'Subsidized_By (\"PARTNER\"/\"GLOVO\"/\"BOTH\")']).strip().upper()),
                                    "externalIds": Promobot.products_ID_list(n),
                                    "addresses": Promobot.store_addresses_ID_list(n),
                                    "commissionOnDiscountedPrice":False,
                                    "subsidyStrategy":"BY_PERCENTAGE",
                                    "sponsors":[{"sponsorId":1,
                                        "sponsorOrigin":"GLOVO",
                                        "subsidyValue":Promobot.subsidyValue("glovo", n)},
                                        {"sponsorId":2,
                                        "sponsorOrigin":"PARTNER",
                                        "subsidyValue":Promobot.subsidyValue("partner", n)}]}],
                        "customerTagId":None,
                        "budget":None,
                        "prime": Promobot.is_prime(n)}
            p = requests.post(url, headers = {'authorization' : access_token}, json = payload)
            if p.ok is False:
                print(f'Promo {n} NOT CREATED')
                if n == 0:
                    print(f'Promo {n} - {p.content}')
                    confirmation = input(f'Continue promo creation? [yes,no]\n')
                    print("\n")
                    if confirmation in ['yes','ye','y','si']:
                        pass
                    else:
                        Promobot.df_to_excel()
                        sys.exit(0)
                try:
                    df_promo.at[n,'Status'] = f"ERROR: {p.json()['error']['message']}"
                except Exception:
                    df_promo.at[n,'Status'] = f"ERROR: {p.content}"
                    if 'Bad request' in str(p.content):
                        print(f'Promo {n} - status: NOT CREATED - INVALID INPUT DATA OR INACTIVE STORE ID')
                else: df_promo.at[n,'Status'] = f"ERROR: {p.json()['error']['message']}"
                finally:
                    print(f'ERROR: {p.text}')
            else:
                df_promo.at[n,'Promo_ID'] = int(p.json()['id'])
                df_promo.at[n,'Status'] = 'created'
                print(f'Promo {n} - status: created; id: {p.json()["id"]}')
                if n == 0:
                    print(f'\nPromo link: https://beta-admin.glovoapp.com/promotions/{p.json()["id"]}')
                    print('Check if promo has been created as expected')
                    confirmation = input(f'Continue promo creation ({len(df_promo)-1} promos left)? [yes,no]\n')
                    print("\n")
                    if confirmation in ['yes','ye','y','si']:
                        pass
                    else:
                        Promobot.df_to_excel()
                        sys.exit(0)

    '''promo deletion'''
    def deletion(n):
        if pd.notna(df_promo.at[n,'Promo_ID']):
            if df_promo.at[n,'Status'] == 'deleted':
                print(n,'already deleted')
            else:
                url = f'https://adminapi.glovoapp.com/admin/partner_promotions/{int(df_promo.at[n,"Promo_ID"])}'
                r = requests.delete(url, headers  = {'authorization' : access_token})
                if r.ok:
                    df_promo.at[n,'Status'] = 'deleted'
                    print(f'Promo {n} - deleted')
                else:
                    print(f'Promo {n} - unable to delete', r.content)
        else:
            print(f'Promo {n} - no promo ID to delete')

    '''promo check'''
    def checker(n):
        if pd.notna(df_promo.at[n,'Promo_ID']):
            if type(df_promo.at[n,'Promo_ID']) == 'str':
                df_promo.at[n,'Promo_ID'] = df_promo.at[n,'Promo_ID'].str.strip()
                #df_promo.at[n,'Promo_ID'] = sub("\D",'',df_promo.at[n,'Promo_ID'])
            url = f"https://adminapi.glovoapp.com/admin/partner_promotions/{int(df_promo.at[n,'Promo_ID'])}"
            p = requests.get(url, headers = {'authorization' : access_token})
            if p.ok is False:
                try:
                    p.json()['error']['message']
                except Exception:
                    df_promo.at[n,'Status'] = p.text
                    print(p.text)
                else:
                    if 'deleted' in p.json()['error']['message']:
                        df_promo.at[n,'Status'] = 'deleted'
                        print(f'Promo {n} - status deleted')
                    else:
                        df_promo.at[n,'Status'] = p.json()['error']['message']
                        print(p.json()['error']['message'])

            else:
                if p.json()['deleted'] == True:
                    df_promo.at[n,'Status'] = 'deleted'
                    print(f'Promo {n} - status deleted')
                else:
                    df_promo.at[n,'Status'] = 'active'
                    print(f'Promo {n} - status active')
        else:
            print(f'Promo {n} - No promo ID to check')

    '''save to excel'''
    def df_to_excel():
        global output_excel
        tz = datetime.now()
        output_excel = os.path.join(output_path, f'{bot_name}_{mode}_{tz.strftime("%Y_%m_%d_(h%H_%M)")}.xlsx')
        try: df_promo.loc[:,["Start_Date (dd/mm/yyyy)","End_Date (included)"]] = df_promo.loc[:,["Start_Date (dd/mm/yyyy)","End_Date (included)"]].apply(lambda x: x.dt.strftime('%d/%m/%Y'))
        except Exception: pass
        #save output
        df_promo.to_excel(output_excel, index = False)
        #with pd.ExcelWriter(output_excel) as writer:
            #df_promo.to_excel(writer, sheet_name = 'Promos', index=False)
            #writer.sheets['Promos'].set_default_row(20)
            #writer.sheets['Promos'].freeze_panes(1, 0)

    '''launcher'''
    def launch(function):
        try:
            for n in df_promo.index[:21]:
                function(n)
        except (KeyboardInterrupt, Exception):
            Promobot.df_to_excel()
        if len(df_promo) > 20:
            try:
                ###using multithreading concurrent futures###
                with ThreadPoolExecutor() as executor:
                    for n in df_promo.index[21:]:
                        executor.submit(function, n)
                ###end multiprocessing process###
            except (KeyboardInterrupt, Exception):
                Promobot.df_to_excel()

    '''main'''
    def main():
        try:
            '''initiation code'''
            Promobot.print_bot_name()
            Promobot.set_path()
            Promobot.logger_start()
            Promobot.login_check()
            Promobot.refresh()
            '''bot code'''
            Promobot.set_mode()
            Promobot.set_input()
            #Promobot.set_output_dir()
            if mode == 'create':
                Promobot.launch(Promobot.creation)
            elif mode == 'delete':
                Promobot.launch(Promobot.deletion)
            elif mode == 'check':
                Promobot.launch(Promobot.checker)
            Promobot.df_to_excel()
            print(f'\n\n{bot_name} has processed {len(df_promo)} Store Addresses\n\nResults are available in file {os.path.relpath(output_excel)}')
        except Exception as e:
            print(repr(e))
            k=input('\nPress Enter x2 to close')
        else:
            k=input('\nPress Enter x2 to close')