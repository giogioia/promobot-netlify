#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""

Partner Promotions Automation Bot
Glovo Italy

version: 1.5

"""

import json
import pandas as pd
import requests
import logging
import sys
import os
import get_new_token
import datetime as dt
from datetime import datetime, timedelta
from time import sleep, time
from concurrent.futures import ThreadPoolExecutor
from subprocess import call
#from regex import sub

class PromoBot:
    global bot_name, exe_version_available, mode, df_promo, output_excel, output_path, platform
    bot_name = 'promobot'
    exe_version_available = 1.1
    '''Init functions'''
    def exe_checker(exe_version):
        try:
            print(exe_version)
            print(exe_version_available)
        except NameError:
            pass
        else:
            if exe_version != exe_version_available:
                PromoBot.set_path()
                conf = input(f'New PromoBot_{exe_version_available} is available. Download new version? [yes/no]\t')
                if conf in ['yes','ye','y','si']:
                    print('Estimated time: ~60 seconds.\nDownload in progress... Do not close the terminal page.')
                    try:
                        r = requests.get(f'https://el-promobot.netlify.app/assets/PromoBot_{exe_version_available}.exe')
                        with open(os.path.join(cwd,f'PromoBot_{exe_version_available}.exe'), 'wb') as file:
                            file.write(r.content)
                    except Exception:
                        print('Something went wrong.\nTry downloading new bot manually from https://el-promobot.netlify.app/')
                        k=input('\nPress Enter x2 to close')
                        sys.exit(0)
                    else: 
                        print(f'\nNew PromoBot_{exe_version_available} successfully downloaded!\nClose this window and lanch PromoBot.exe again')
                        k=input('\nPress Enter x2 to close')

        #Step 1: set path
    def set_path():
        global cwd, token_path, input_path, platform
        cwd = os.getcwd()
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
        #without encryption
        with open(token_path) as read_file:
            content = json.load(read_file)
        glovo_email = content['glovo_email']
        refresh_token = content['refresh_token']
        country = content['country']
        # #with encryption
        # with open(token_path, 'rb') as read_file:
        #     enc_content = read_file.read()
        # dec_content = base64.b64decode(enc_content).decode("utf-32")
        # content = json.loads(dec_content)
        # glovo_email = content['glovo_email']
        # refresh_token = content['refresh_token']
        # country = content['country']
    #Step 3: check login credentials
    def login_check():
        #Check/get login data: check if file 'my personal token' exists and read it to get login data.
        global glovo_email, refresh_token
        #print("Checking login data")
        if os.path.isfile(token_path):
            try:
                PromoBot.read_json()
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
        PromoBot.read_json()
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
            #without enc
            content['refresh_token'] = new_refresh_token
            with open(token_path, "r+") as dst_file:
                json.dump(content, dst_file)
            # #with enc
            # content['refresh_token'] = new_refresh_token
            # json_data = json.dumps(content)  
            # with open(token_path, "wb") as dst_file:
            #     dst_file.write(base64.b64encode(str(json_data).encode("utf-32")))
            #print("token refreshed")
        else:
            print(f"Token NOT refreshed -> {oauth_request.content}")
            logger.info(f'Access Token NOT Refreshed -> {oauth_request.content}')


    def print_bot_name():
        print(f'\nEl PromoBot')
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
                elif a_or_b == "b":
                    print('\nSelected mode is "Delete Promos"')
                    mode = 'delete'
                elif a_or_b == "c":
                    print('\nSelected mode is "Promos status check"')
                    mode = 'check'
                break

    #custom function for set_input(): extracts dataframe once input name is set
    def import_data(input_file):
        global df_promo, no_prime, no_products, no_store_address, no_budget, no_commissionOnDiscountedPrice
        #import data
        df_promo = pd.read_excel(input_file, usecols=lambda x: 'Unnamed' not in x, engine='openpyxl', dtype=object)
        #clean empty rows
        df_promo.dropna(how='all', inplace = True)
        #reset index after deleting empty rows
        df_promo.reset_index(drop = True, inplace = True)
        #check if Status column in present
        if 'Status' not in list(df_promo):
            df_promo.loc[:,'Status']= None
        if mode == 'create':
            #clean str columns
            try:
                df_promo.loc[:,'City_Code'] = df_promo.loc[:,'City_Code'].str.strip()
                df_promo.loc[:,'Promo_Name'] = df_promo.loc[:,'Promo_Name'].str.strip()
            except AttributeError:pass
            # try:
            #     df_promo.loc[:,'Promo_Type ("FLAT"/"FREE"/"XX%"/"2for1")'] = df_promo.loc[:,'Promo_Type ("FLAT"/"FREE"/"XX%"/"2for1")'].str.strip()
            # except AttributeError:pass
            #Only Prime
            if 'Only_Prime' not in list(df_promo):
                no_prime = True
            else:
                if pd.notna(df_promo.loc[:,'Only_Prime']).sum() == 0:
                    no_prime = True
                else:
                    if df_promo.loc[:,'Only_Prime'].dtype != 'O':
                        print('\nColumn Only_Prime must have text format.\nCheck input file and try again.')
                        raise AttributeError
                    else: 
                        no_prime = False
                        df_promo.loc[:,'Only_Prime'] = df_promo.loc[:,'Only_Prime'].str.strip()
            #Budget
            if 'Budget' not in list(df_promo):
                no_budget = True
            else:
                if pd.notna(df_promo.loc[:,'Budget']).sum() == 0:
                    no_budget = True
                else:
                    no_budget = False
            #no_commissionOnDiscountedPrice     
            if 'Commission_On_Discounted_Price' not in list(df_promo):
                no_commissionOnDiscountedPrice = True
            else:
                if pd.notna(df_promo.loc[:,'Commission_On_Discounted_Price']).sum() == 0:
                    no_commissionOnDiscountedPrice = True
                else:
                    if df_promo.loc[:,'Commission_On_Discounted_Price'].dtype != 'O':
                        raise ValueError('Column Commission_On_Discounted_Price must have text format.\nCheck input file and try again.')
                    else: 
                        no_commissionOnDiscountedPrice = False
            #Pructs columns
            if any(list(map(lambda x: 'Product' in x, list(df_promo)))) == False:
                no_products = True
            else:
                if pd.notna(df_promo.loc[:,list(map(lambda x: 'Product' in x, list(df_promo)))]).sum().sum() ==  0:
                    no_products = True
                else:
                    no_products = False
            #Store_Address columns
            if any(list(map(lambda x: 'Store_Address' in x, list(df_promo)))) == False:
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
                try: df_promo.loc[:,"%GLOVO"]= df_promo.loc[:,"%GLOVO"].str.strip('%')
                except AttributeError: pass
                else: 
                    try: df_promo.loc[:,"%GLOVO"].astype('int')
                    except ValueError: pass
            #%partner
            if df_promo.loc[:,"%PARTNER"].dtype == 'O':
                try: df_promo.loc[:,"%PARTNER"]= df_promo.loc[:,"%PARTNER"].str.strip('%')
                except AttributeError: pass
                else: 
                    try: df_promo.loc[:,"%PARTNER"].astype('int')
                    except ValueError: pass
        elif mode == 'delete':
            if ('Promo_ID' not in list(df_promo)) or (pd.notna(df_promo.loc[:,'Promo_ID']).sum() == 0):
                raise KeyError('Promo_ID')
        elif mode == 'check':
            if ('Promo_ID' not in list(df_promo)) or (pd.notna(df_promo.loc[:,'Promo_ID']).sum() == 0):
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
                input_path = PromoBot.find_excel_file_path(input_name)
            #print('cwd',cwd)
            #print('input_path',input_path)
            except NameError:
                sleep(0.5)
                print(f'\nCould not find {input_name} in {os.path.basename(cwd)}\nPlease try again\n')
                continue
            else:
                try:
                    PromoBot.import_data(input_path)
                except KeyError as e:
                    print(f'Column {e} is missing. Unable to import data.\nUpdate file or choose another file.')
                    continue
                else:
                    #print(f'Promos will will be Created for the {len(df_promo)} Store IDs found in {input_name}')
                    PromoBot.set_output_dir()
                    sleep(1)
                    confirm_path = input(f'\nUsing file \'{input_name}\' to {mode} promos of {len(df_promo)} Store IDs found in {os.path.join(os.path.basename(os.path.dirname(input_path)),os.path.basename(input_path))}.\nOutput will be saved in folder \'{os.path.basename(output_path)}\'\nContinue? [yes,no]\t')
                    if confirm_path in ["yes","y","ye","si"]:
                        logger.info(f'Using file {input_name} in folder {os.path.basename(os.path.dirname(input_path))}')
                        print('')
                        break

    def is_number(x):
        if isinstance(x, str):
            if x.isdigit(): return True
            else: return False
        else:
            try: 
                x += 0
            except TypeError: return False
            else: return True

    '''promo creation'''
    def p_type(promo_type):
        if type(promo_type) == 'str':
            promo_type.strip().upper()
        if promo_type == 'FLAT':
            return 'FLAT_DELIVERY'
        elif promo_type == 'FREE':
            return 'FREE_DELIVERY'
        elif promo_type == '2for1':
            return 'TWO_FOR_ONE'
        elif PromoBot.is_number(promo_type) or (isinstance(promo_type,str) and '%' in promo_type):
            return 'PERCENTAGE_DISCOUNT'
        else:
            raise ValueError('INVALID PROMO TYPE')

    def del_fee(promo_type):
        if promo_type == 'FLAT':
            return 100
        if promo_type == 'FREE':
            return None
        else:
            return None

    def perc(promo_type):
        if PromoBot.p_type(promo_type) == 'PERCENTAGE_DISCOUNT':
            if isinstance(promo_type,str):
                return float((promo_type).strip('%'))
            else:
                if promo_type < 1:
                    return float(promo_type * 100)
                else:
                    return float(promo_type)
        else:
            return None

    def strat(subsidy):
        return f'ASSUMED_BY_{subsidy}'

    def paymentStrat(subsidy):
        if PromoBot.strat(subsidy) == "ASSUMED_BY_GLOVO" or PromoBot.strat(subsidy) == "ASSUMED_BY_PARTNER":
            return PromoBot.strat(subsidy)
        elif PromoBot.strat(subsidy) == "ASSUMED_BY_THIRD_PARTY":
            return "ASSUMED_BY_GLOVO"
        else: 
            return "ASSUMED_BY_PARTNER"

    def sponsors(subsidy, n):
        if PromoBot.strat(subsidy) == "ASSUMED_BY_GLOVO":
            return [{"sponsorId":1,
                    "sponsorOrigin":"GLOVO",
                    "subsidyValue": 100}]
                    #"subsidyValue":int(PromoBot.subsidyValue("glovo", n))}]
        if PromoBot.strat(subsidy) == "ASSUMED_BY_PARTNER":
            return [{"sponsorId":2,
                    "sponsorOrigin":"PARTNER",
                    "subsidyValue": 100}]
                    #"subsidyValue":int(PromoBot.subsidyValue("partner", n))}]
        if PromoBot.strat(subsidy) == "ASSUMED_BY_THIRD_PARTY":
            return [{"sponsorId":3,
                    "sponsorOrigin":"THIRD_PARTY",
                    "sponsorName":"Third Party",
                    "subsidyValue": 100}]
                    #"subsidyValue":int(PromoBot.subsidyValue("partner", n))}]
        if PromoBot.strat(subsidy) == "ASSUMED_BY_BOTH":
            return [{"sponsorId":1,
                    "sponsorOrigin":"GLOVO",
                    "subsidyValue":int(PromoBot.subsidyValue("glovo", n))},
                    {"sponsorId":2,
                    "sponsorOrigin":"PARTNER",
                    "subsidyValue":int(PromoBot.subsidyValue("partner", n))}]

    def get_utc_timestamp(local_time):
        return (local_time - datetime(1970, 1, 1)).total_seconds()

    def time_code(x, date):
        if x == 'start':
            hours_added = timedelta(hours = 5)
            future_date = date + hours_added
            stamp = PromoBot.get_utc_timestamp(future_date)
            return int(stamp*1000)
        if x == 'end':
            hours_added = timedelta(hours = 25)
            future_date = date + hours_added
            stamp = PromoBot.get_utc_timestamp(future_date)
            return int(stamp*1000)

    def products_ID_list(n):
        if no_products:
            return None
        else:
            prods_list = []
            for i in range(1,10):
                try:
                    df_promo.loc[n,f'Product_ID{i}']
                except KeyError:
                    break
                else:
                    if pd.isna(df_promo.loc[n,f'Product_ID{i}']): 
                        continue
                    else: 
                        if PromoBot.is_number(df_promo.loc[n,f'Product_ID{i}']):
                            prods_list.append((str(df_promo.loc[n,f'Product_ID{i}'])).replace('\ufeff', ''))
                        else: 
                            prods_list.append((str(df_promo.loc[n,f'Product_ID{i}'])).replace('\ufeff', ''))
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
                    df_promo.loc[n,f'Store_Address{o}']
                except KeyError:
                    break
                else:
                    if pd.isna(df_promo.loc[n,f'Store_Address{o}']): continue
                    if type(df_promo.loc[n,f'Store_Address{o}']) == str:
                        df_promo.loc[n,f'Store_Address{o}'] = df_promo.loc[n,f'Store_Address{o}'].replace('\ufeff', '')
                    sa_ID_list.append(int(df_promo.loc[n,f'Store_Address{o}']))
            if sa_ID_list == []:
                return None
            else:
                return sa_ID_list

    def subsidyValue(subject, n):
        if PromoBot.strat((df_promo.loc[n,'Subsidized_By (\"PARTNER\"/\"GLOVO\"/\"BOTH\")']).strip().upper()) == 'ASSUMED_BY_GLOVO':
                if subject == 'glovo':
                    return 100
                if subject == 'partner':
                    return 0
        elif PromoBot.strat((df_promo.loc[n,'Subsidized_By (\"PARTNER\"/\"GLOVO\"/\"BOTH\")']).strip().upper()) == 'ASSUMED_BY_PARTNER':
            if subject == 'glovo':
                return 0
            if subject == 'partner':
                return 100
        elif PromoBot.strat((df_promo.loc[n,'Subsidized_By (\"PARTNER\"/\"GLOVO\"/\"BOTH\")']).strip().upper()) == 'ASSUMED_BY_THIRD_PARTY':
            if subject == 'glovo':
                return 0
            if subject == 'partner':
                return 100
        elif PromoBot.strat((df_promo.loc[n,'Subsidized_By (\"PARTNER\"/\"GLOVO\"/\"BOTH\")']).strip().upper()) == 'ASSUMED_BY_BOTH':
            if subject == 'glovo':
                if df_promo.loc[n,"%GLOVO"] < 1:
                    return int(df_promo.loc[n,"%GLOVO"] * 100)
                else:
                    return int(df_promo.loc[n,"%GLOVO"])
            if subject == 'partner':
                if df_promo.loc[n,"%PARTNER"] < 1:
                    return int(df_promo.loc[n,"%PARTNER"] * 100)
                else:
                    return int(df_promo.loc[n,"%PARTNER"])

    def is_prime(n):
        if no_prime:
            return None
        else:
            if df_promo.loc[n,'Only_Prime'] == 'yes':
                return True
            else:
                return False

    def with_budget(n):
        if no_budget:
            return None
        else:
            try:
                int(df_promo.loc[n,'Budget'])
            except ValueError:
                return None
            else:
                return int(df_promo.loc[n,'Budget'])
    
    def commissionOnDiscountedPrice(n):
        if no_commissionOnDiscountedPrice:
            return None
        else:
            if df_promo.loc[n,'Commission_On_Discounted_Price'] == 'yes':
                return True
            else:
                return False
  
    def creation(n):
        if df_promo.loc[n,'Status'] == 'created':
            print(n,'already created')
        else:
            url = 'https://adminapi.glovoapp.com/admin/partner_promotions'
            payload = {"name": df_promo.loc[n,'Promo_Name'],
                        "cityCode": df_promo.loc[n,'City_Code'],
                        "type": PromoBot.p_type(df_promo.loc[n,'Promo_Type ("FLAT"/"FREE"/"XX%"/"2for1")']),
                        "percentage": PromoBot.perc(df_promo.loc[n,'Promo_Type ("FLAT"/"FREE"/"XX%"/"2for1")']),
                        "deliveryFeeCents": PromoBot.del_fee(df_promo.loc[n,'Promo_Type ("FLAT"/"FREE"/"XX%"/"2for1")']),
                        "startDate": PromoBot.time_code('start',df_promo.loc[n,"Start_Date (dd/mm/yyyy)"]),
                        "endDate": PromoBot.time_code('end',df_promo.loc[n,"End_Date (included)"]),
                        "openingTimes": None,
                        "partners":[{"id": int(df_promo.loc[n,'Store_ID']),
                                    "paymentStrategy": PromoBot.paymentStrat((df_promo.loc[n,'Subsidized_By (\"PARTNER\"/\"GLOVO\"/\"BOTH\")']).strip().upper()),
                                    "externalIds": PromoBot.products_ID_list(n),
                                    "addresses": PromoBot.store_addresses_ID_list(n),
                                    "commissionOnDiscountedPrice":PromoBot.commissionOnDiscountedPrice(n),
                                    "subsidyStrategy":"BY_PERCENTAGE",
                                    "sponsors":PromoBot.sponsors((df_promo.loc[n,'Subsidized_By (\"PARTNER\"/\"GLOVO\"/\"BOTH\")']).strip().upper(), n)}],
                                    # "sponsors":[{"sponsorId":1,
                                    #     "sponsorOrigin":"GLOVO",
                                    #     "subsidyValue":int(PromoBot.subsidyValue("glovo", n))},
                                    #     {"sponsorId":2,
                                    #     "sponsorOrigin":"PARTNER",
                                    #     "subsidyValue":int(PromoBot.subsidyValue("partner", n))}]}],
                        "customerTagId":None,
                        "budget":PromoBot.with_budget(n),
                        "prime": PromoBot.is_prime(n)}
            print('PAYLOAD TEST:', payload)
            sleep(1)
            p = requests.post(url, headers = {'authorization' : access_token}, json = payload)
            sleep(1)
            PromoBot.refresh()
            if p.ok is False:
                print(f'Promo {n} NOT CREATED')
                if n == 0:
                    print(f'Promo {n} - {p.content}')
                    confirmation = input(f'Continue promo creation? [yes,no]\t')
                    print("\n")
                    if confirmation in ['yes','ye','y','si']:
                        pass
                    else:
                        PromoBot.df_to_excel()
                        sys.exit(0)
                try:
                    df_promo.loc[n,'Status'] = f"ERROR: {p.json()['error']['message']}"
                except Exception:
                    df_promo.loc[n,'Status'] = f"ERROR: {p.content}"
                    if 'Bad request' in str(p.content):
                        print(f'Promo {n} - status: NOT CREATED - INVALID INPUT DATA OR INACTIVE STORE ID')
                else: df_promo.loc[n,'Status'] = f"ERROR: {p.json()['error']['message']}"
                finally:
                    print(f'ERROR: {p.text}')
            else:
                df_promo.loc[n,'Promo_ID'] = int(p.json()['id'])
                df_promo.loc[n,'Status'] = 'created'
                print(f'Promo {n} - status: created; id: {p.json()["id"]}')
                if n == 0:
                    print(f'\nPromo link: https://beta-admin.glovoapp.com/promotions/{p.json()["id"]}')
                    print('Check if promo has been created as expected')
                    confirmation = input(f'Continue promo creation ({len(df_promo)-1} promos left)? [yes,no]\n')
                    if confirmation in ['yes','ye','y','si']:
                        pass
                    else:
                        delete = input(f'\nDelete promo {n}? [yes/no]\t')
                        if delete in ['yes','ye','y','si']:
                            url = f'https://adminapi.glovoapp.com/admin/partner_promotions/{int(df_promo.loc[n,"Promo_ID"])}'
                            sleep(2)
                            r = requests.delete(url, headers  = {'authorization' : access_token})
                            if r.ok:
                                df_promo.loc[n,'Status'] = 'deleted'
                                print(f'Promo {n} - deleted')
                            else:
                                print(f'Promo {n} - unable to delete', r.content)
                        PromoBot.df_to_excel()
                        sys.exit(0)

    '''promo deletion'''
    def deletion(n):
        if pd.notna(df_promo.loc[n,'Promo_ID']):
            if df_promo.loc[n,'Status'] == 'deleted':
                print(n,'already deleted')
            else:
                url = f'https://adminapi.glovoapp.com/admin/partner_promotions/{int(df_promo.loc[n,"Promo_ID"])}'
                sleep(2)
                r = requests.delete(url, headers  = {'authorization' : access_token})
                if r.ok:
                    df_promo.loc[n,'Status'] = 'deleted'
                    print(f'Promo {n} - deleted')
                else:
                    print(f'Promo {n} - unable to delete', r.content)
        else:
            print(f'Promo {n} - no promo ID to delete')

    '''promo check'''
    def checker(n):
        if pd.notna(df_promo.loc[n,'Promo_ID']):
            if isinstance(df_promo.loc[n,'Promo_ID'],str):
                df_promo.loc[n,'Promo_ID'] = df_promo.loc[n,'Promo_ID'].str.strip()
                #df_promo.loc[n,'Promo_ID'] = sub("\D",'',df_promo.loc[n,'Promo_ID'])
            url = f"https://adminapi.glovoapp.com/admin/partner_promotions/{int(df_promo.loc[n,'Promo_ID'])}"
            sleep(2)
            p = requests.get(url, headers = {'authorization' : access_token})
            if p.ok is False:
                try:
                    p.json()['error']['message']
                except Exception:
                    df_promo.loc[n,'Status'] = p.text
                    print(p.text)
                else:
                    if 'deleted' in p.json()['error']['message']:
                        df_promo.loc[n,'Status'] = 'deleted'
                        print(f'Promo {n} - status deleted')
                    else:
                        df_promo.loc[n,'Status'] = p.json()['error']['message']
                        print(p.json()['error']['message'])

            else:
                if p.json()['deleted'] == True:
                    df_promo.loc[n,'Status'] = 'deleted'
                    print(f'Promo {n} - status deleted')
                else:
                    df_promo.loc[n,'Status'] = 'active'
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
        print('\nOutput Excel saved')
        #with pd.ExcelWriter(output_excel) as writer:
            #df_promo.to_excel(writer, sheet_name = 'Promos', index=False)
            #writer.sheets['Promos'].set_default_row(20)
            #writer.sheets['Promos'].freeze_panes(1, 0)

    '''launcher'''
    def launch(function):
        for n in df_promo.index[:21]:
            try:
                function(n)
            except Exception as e:
                print(f'Problem with data of promo {n} - unable to process')
                print(repr(e))
                pass
        if len(df_promo) > 20:
            ###using multithreading concurrent futures###
            with ThreadPoolExecutor() as executor:
                for n in df_promo.index[21:]:
                    try:
                        executor.submit(function, n)
                    except Exception as e:
                        print(f'Problem with data of promo {n} - unable to process')
                        print(repr(e))
                        pass
            ###end multiprocessing process###

    '''driver'''
    def driver():
        try:
            '''initiation code'''
            PromoBot.print_bot_name()
            PromoBot.set_path()
            PromoBot.logger_start()
            PromoBot.login_check()
            PromoBot.refresh()
            '''bot code'''
            PromoBot.set_mode()
            PromoBot.set_input()
            #PromoBot.set_output_dir()
            if mode == 'create':
                PromoBot.launch(PromoBot.creation)
            elif mode == 'delete':
                PromoBot.launch(PromoBot.deletion)
            elif mode == 'check':
                PromoBot.launch(PromoBot.checker)
            PromoBot.df_to_excel()
            print(f'\n\n{bot_name} has processed {len(df_promo)} Store Addresses\n\nResults are available in file {os.path.relpath(output_excel)}')
            k=input('\n\nPress Enter x2 to close')
        except (KeyboardInterrupt, Exception) as e:
            print(repr(e))
            try: PromoBot.df_to_excel()
            except Exception: print('Unable to save output data')
            k=input('\nPress Enter x2 to close')
            
